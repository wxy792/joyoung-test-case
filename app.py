#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
九阳程序用例转换网站后端
支持两种转换：
1. 菜单流程 -> 菜单一致性核对用例
2. 报警文件 -> 报警用例
"""

import os
import io
import json
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from flask import Flask, render_template, request, send_file, flash, redirect, url_for

app = Flask(__name__,
            template_folder=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'templates'),
            static_folder=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static'))
app.secret_key = 'joyoung-test-case-tool-2026'
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0  # 禁用静态文件缓存
app.config['TEMPLATES_AUTO_RELOAD'] = True  # 模板自动重载

UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
OUTPUT_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'outputs')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# ===================== 通用样式 =====================

def apply_header_style(cell):
    """表头样式：深蓝底白字、加粗、居中"""
    cell.font = Font(name='微软雅黑', bold=True, color='FFFFFF', size=11)
    cell.fill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin = Side(style='thin', color='B8CCE4')
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def apply_data_style(cell, is_step_col=False):
    """数据行样式"""
    cell.font = Font(name='微软雅黑', size=10)
    cell.alignment = Alignment(horizontal='center' if is_step_col else 'left',
                                vertical='center', wrap_text=True)
    thin = Side(style='thin', color='B8CCE4')
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def auto_width(sheet, min_width=8, max_width=60):
    """自动调整列宽"""
    for col in sheet.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                val_len = len(str(cell.value))
                # 考虑中文字符
                chinese_count = sum(1 for c in str(cell.value) if '\u4e00' <= c <= '\u9fff')
                effective_len = val_len + chinese_count
                max_len = max(max_len, min(effective_len, max_width))
        adjusted = min(max(max_len + 2, min_width), max_width)
        sheet.column_dimensions[col_letter].width = adjusted


# ===================== 转换功能1：菜单一致性核对 =====================

def convert_menu_consistency(input_path, output_path):
    """
    将菜单流程图转换为菜单一致性核对用例
    处理逻辑：
    1. 遍历所有sheet（跳过名为'目录'的sheet）
    2. 删除D列（实时时间/倒计时）
    3. 在D列填入'一致性评估'表头，各行填入'符合:□     不符合:□'
    4. 检查A列步骤序号是否连续，空单元格也补充数字
    5. 所有单元格水平居中、自动换行
    6. 自动调整行高，避免文字被覆盖
    """
    wb = openpyxl.load_workbook(input_path)
    new_wb = openpyxl.Workbook()
    new_wb.remove(new_wb.active)

    for sheet_name in wb.sheetnames:
        if sheet_name == '目录':
            continue

        src_sheet = wb[sheet_name]
        new_sheet = new_wb.create_sheet(title=sheet_name)

        all_rows = list(src_sheet.iter_rows(values_only=True))
        if not all_rows:
            continue

        # 找到表头行
        header_row_idx = None
        for i, row in enumerate(all_rows):
            row_vals = [str(v).strip() if v else '' for v in row]
            joined = ''.join(row_vals)
            if '步骤' in joined and ('动作' in joined or '标准时间' in joined):
                header_row_idx = i
                break

        if header_row_idx is None:
            header_row_idx = 1

        new_row_idx = 1
        expected_step = 1

        # 处理标题行（表头之前的行）
        for i in range(header_row_idx):
            row_vals = all_rows[i]
            if all(v is None or str(v).strip() == '' for v in row_vals):
                continue
            title_val = next((v for v in row_vals if v is not None and str(v).strip()), '')
            cell = new_sheet.cell(row=new_row_idx, column=1, value=title_val)
            cell.font = Font(name='微软雅黑', bold=True, size=14, color='1e40af')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            new_sheet.merge_cells(f'A{new_row_idx}:E{new_row_idx}')
            new_sheet.row_dimensions[new_row_idx].height = 28
            new_row_idx += 1

        # 处理表头行
        new_header = ['步骤', '动作', '标准时间', '一致性评估', '备注']
        for col_idx, val in enumerate(new_header, 1):
            cell = new_sheet.cell(row=new_row_idx, column=col_idx, value=val)
            apply_header_style(cell)
        new_sheet.row_dimensions[new_row_idx].height = 30
        header_output_row = new_row_idx
        new_row_idx += 1

        # 处理数据行
        for row_idx in range(header_row_idx + 1, len(all_rows)):
            row_vals = list(all_rows[row_idx])

            # 跳过完全空的行
            if all(v is None or str(v).strip() == '' for v in row_vals):
                continue

            # 原数据列: A=步骤, B=动作, C=标准时间, D=实时时间(丢弃), E=备注
            step_val = row_vals[0] if len(row_vals) > 0 else None
            action_val = row_vals[1] if len(row_vals) > 1 else None
            std_time_val = row_vals[2] if len(row_vals) > 2 else None
            remark_val = row_vals[4] if len(row_vals) > 4 else None

            # 需求1：不管A列是否为空，都补充连续的步骤序号
            step_num = expected_step
            expected_step += 1
            new_sheet.cell(row=new_row_idx, column=1, value=step_num)

            # 需求2：B列（动作）为空时，A列原文字放入B列，B、C合并
            # 适用于：原文件某些行A列有说明文字、B/C列为空的情况
            merge_bc = False
            if action_val is None or str(action_val).strip() == '':
                merge_bc = True
                # A列原文字作为B列内容（B、C合并显示），D列留空（下拉选择）
                header_text = str(step_val).strip() if (step_val is not None and str(step_val).strip()) else ''
                new_sheet.cell(row=new_row_idx, column=2, value=header_text)
            else:
                # 正常数据行，D列留空（下拉选择）
                new_sheet.cell(row=new_row_idx, column=2, value=action_val)
                new_sheet.cell(row=new_row_idx, column=3, value=std_time_val)
                new_sheet.cell(row=new_row_idx, column=5, value=remark_val)

            # 需求3+4：所有单元格水平居中、自动换行、自适应行高
            max_lines = 1
            cols_to_style = [1, 2, 4, 5] if merge_bc else [1, 2, 3, 4, 5]
            for col_idx in cols_to_style:
                cell = new_sheet.cell(row=new_row_idx, column=col_idx)
                cell.font = Font(name='微软雅黑', size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                thin = Side(style='thin', color='B8CCE4')
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
                if cell.value:
                    lines = str(cell.value).count('\n') + 1
                    max_lines = max(max_lines, lines)

            if merge_bc:
                new_sheet.merge_cells(f'B{new_row_idx}:C{new_row_idx}')
                # 合并后重新设定B列对齐方式
                b_cell = new_sheet.cell(row=new_row_idx, column=2)
                b_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

            new_sheet.row_dimensions[new_row_idx].height = max(20, max_lines * 18)

            new_row_idx += 1

        # 为D列（一致性评估）添加数据验证：下拉选择「符合/不符合」，二选一
        # 方案：在隐藏列（F列）写入选项，然后引用该列作为下拉源（兼容性最好）
        if new_row_idx > header_output_row + 1:
            # 在F列（第6列）写入选项（隐藏该列）
            new_sheet.cell(row=1, column=6, value='符合')
            new_sheet.cell(row=2, column=6, value='不符合')
            new_sheet.column_dimensions['F'].hidden = True

            # 创建数据验证，引用F1:F2作为选项源
            dv = DataValidation(
                type="list",
                formula1="=$F$1:$F$2",
                allow_blank=True,
                prompt="请选择一致性评估结果",
                promptTitle="一致性评估"
            )
            col_d_range = f'D{header_output_row + 1}:D{new_row_idx - 1}'
            dv.add(col_d_range)
            new_sheet.add_data_validation(dv)

        # 为D列添加条件格式：值为「不符合」时红底强调
        if new_row_idx > header_output_row + 1:
            red_fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
            col_d_range = f'D{header_output_row + 1}:D{new_row_idx - 1}'
            # FormulaRule: 用公式 D3="不符合" 匹配（条件格式会自动按行调整）
            rule = FormulaRule(formula=[f'D{header_output_row + 1}="不符合"'], fill=red_fill)
            new_sheet.conditional_formatting.add(col_d_range, rule)

        # 设置列宽
        col_widths = [8, 40, 12, 20, 30]
        for i, width in enumerate(col_widths, 1):
            new_sheet.column_dimensions[get_column_letter(i)].width = width

        # 冻结首行
        new_sheet.freeze_panes = f'A{header_output_row + 1}'

    new_wb.save(output_path)
    return len(new_wb.sheetnames)


# ===================== 转换功能2：报警用例Word生成 =====================

def convert_alarm_word(template_path, src_path, output_path):
    """
    将线路板功能说明书（docx）转换为报警核对Word文档
    使用用户上传的文件作为模板，保持格式不变
    """
    import shutil
    import copy
    from docx import Document
    import re

    # 1. 复制模板文件
    shutil.copyfile(template_path, output_path)

    # 2. 打开复制后的文件
    doc = Document(output_path)

    # 3. 提取报警数据
    src_doc = Document(src_path)
    alarms = []

    # 查找包含报警数据的表格
    for table in src_doc.tables:
        for row in table.rows:
            cells_text = [cell.text.strip() for cell in row.cells]

            # 查找包含"U"开头的报警代码
            for i, text in enumerate(cells_text):
                if re.match(r'^U\d+', text):
                    alarm_code = text.strip()
                    alarm_code = alarm_code.replace('\n', '').replace('\r', '')

                    # 提取报警类型
                    alarm_type = ""
                    if "(" in alarm_code and ")" in alarm_code:
                        start = alarm_code.find("(")
                        end = alarm_code.find(")")
                        alarm_type = alarm_code[start+1:end]
                        alarm_code = alarm_code[:start] + "：" + alarm_type
                    elif "：" in alarm_code:
                        parts = alarm_code.split("：")
                        if len(parts) > 1:
                            alarm_type = parts[1].strip()
                    elif ":" in alarm_code:
                        parts = alarm_code.split(":")
                        if len(parts) > 1:
                            alarm_type = parts[1].strip()
                            alarm_code = parts[0] + "：" + alarm_type

                    # 查找报警文案
                    alarm_text = ""
                    for j in range(i+1, len(cells_text)):
                        if cells_text[j]:
                            alarm_text = cells_text[j]
                            break

                    alarms.append({
                        'code': alarm_code,
                        'type': alarm_type,
                        'text': alarm_text
                    })

    if len(alarms) == 0:
        return 0

    # 4. 保存第一个报警表格的XML
    if len(doc.tables) <= 1:
        return 0

    template_table_xml = copy.deepcopy(doc.tables[1]._element)

    # 5. 更新第一个报警
    title_cell = doc.tables[1].rows[0].cells[0]
    title_cell.text = alarms[0]['code']

    if len(doc.tables[1].rows) > 2:
        alarm_text_cell = doc.tables[1].rows[2].cells[2]
        alarm_text_cell.text = alarms[0]['text']

    # 6. 为其他报警添加表格
    for i in range(1, len(alarms)):
        alarm = alarms[i]

        # 复制表格XML
        new_xml = copy.deepcopy(template_table_xml)

        # 插入到文档末尾
        body = doc.element.body
        body.append(new_xml)

        # 保存并重新打开
        doc.save(output_path)
        doc = Document(output_path)

        # 更新新表格
        title_cell = doc.tables[-1].rows[0].cells[0]
        title_cell.text = alarm['code']

        if len(doc.tables[-1].rows) > 2:
            alarm_text_cell = doc.tables[-1].rows[2].cells[2]
            alarm_text_cell.text = alarm['text']

    # 7. 保存最终文档
    doc.save(output_path)

    return len(alarms)


# ===================== Flask 路由 =====================

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert/menu', methods=['GET', 'POST'])
def convert_menu():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('请选择文件')
            return redirect(request.url)

        file = request.files['file']
        if file.filename == '':
            flash('请选择文件')
            return redirect(request.url)

        if not file.filename.endswith(('.xlsx', '.xls')):
            flash('仅支持 .xlsx 或 .xls 格式')
            return redirect(request.url)

        # 保存上传文件
        filename = file.filename
        input_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(input_path)

        # 转换
        output_filename = filename.replace('.xlsx', '').replace('.xls', '') + '_一致性核对.xlsx'
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)

        try:
            sheet_count = convert_menu_consistency(input_path, output_path)
            flash(f'转换成功！共处理 {sheet_count} 个sheet。点击下方按钮下载结果。')
            return render_template('result.html',
                                download_url=url_for('download_file', filename=output_filename),
                                filename=output_filename)
        except Exception as e:
            flash(f'转换失败：{str(e)}')
            return redirect(request.url)

    return render_template('upload.html', title='菜单一致性核对转换', action=url_for('convert_menu'))

@app.route('/convert/alarm', methods=['GET', 'POST'])
def convert_alarm_route():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('请选择文件')
            return redirect(request.url)

        file = request.files['file']
        if file.filename == '':
            flash('请选择文件')
            return redirect(request.url)

        if not file.filename.endswith(('.xlsx', '.xls')):
            flash('仅支持 .xlsx 或 .xls 格式')
            return redirect(request.url)

        filename = file.filename
        input_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(input_path)

        output_filename = filename.replace('.xlsx', '').replace('.xls', '') + '_报警用例.xlsx'
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)

        try:
            sheet_count = convert_alarm(input_path, output_path)
            flash(f'转换成功！共处理 {sheet_count} 个sheet。')
            return render_template('result.html',
                                download_url=url_for('download_file', filename=output_filename),
                                filename=output_filename)
        except Exception as e:
            flash(f'转换失败：{str(e)}')
            return redirect(request.url)

    return render_template('upload.html', title='报警用例转换（Excel）', action=url_for('convert_alarm_route'))


@app.route('/convert/alarm-word', methods=['GET', 'POST'])
def convert_alarm_word_route():
    """报警核对Word生成"""
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('请选择文件')
            return redirect(request.url)

        file = request.files['file']
        if file.filename == '':
            flash('请选择文件')
            return redirect(request.url)

        if not file.filename.endswith('.docx'):
            flash('仅支持 .docx 格式（Word文档）')
            return redirect(request.url)

        filename = file.filename
        input_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(input_path)

        # 使用用户上传的文件作为模板
        # 修复路径问题：使用基于app.py所在目录的绝对路径
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(BASE_DIR, '程序/程序/L15-Pxx-样机-报警核对-版本号 - 训练1.docx')

        output_filename = filename.replace('.docx', '') + '_报警核对.docx'
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)

        try:
            alarm_count = convert_alarm_word(template_path, input_path, output_path)
            flash(f'生成成功！共生成 {alarm_count} 个报警用例。')
            return render_template('result.html',
                                download_url=url_for('download_file', filename=output_filename),
                                filename=output_filename)
        except Exception as e:
            flash(f'生成失败：{str(e)}')
            return redirect(request.url)

    return render_template('upload.html', title='报警核对Word生成', action=url_for('convert_alarm_word_route'))

@app.route('/download/<filename>')
def download_file(filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    return send_file(path, as_attachment=True)

@app.route('/preview', methods=['POST'])
def preview():
    """预览转换结果（返回JSON）"""
    if 'file' not in request.files:
        return {'error': '请选择文件'}, 400

    file = request.files['file']
    convert_type = request.form.get('type', 'menu')

    if file.filename == '':
        return {'error': '请选择文件'}, 400

    filename = file.filename
    input_path = os.path.join(UPLOAD_FOLDER, '_preview_' + filename)
    file.save(input_path)

    try:
        wb = openpyxl.load_workbook(input_path)
        sheets_info = []
        for sheet_name in wb.sheetnames:
            if sheet_name == '目录':
                continue
            sheet = wb[sheet_name]
            rows = list(sheet.iter_rows(values_only=True))
            # 找到表头
            header = []
            data_preview = []
            for i, row in enumerate(rows):
                row_vals = [str(v)[:50] if v else '' for v in row]
                if not header and any('步骤' in v or '动作' in v for v in row_vals if v):
                    header = row_vals
                elif header and len(data_preview) < 5:
                    data_preview.append(row_vals)
                elif header and len(data_preview) >= 5:
                    break

            sheets_info.append({
                'name': sheet_name,
                'header': header,
                'preview': data_preview,
                'row_count': max(0, len(rows) - (1 if header else 0))
            })

        os.remove(input_path)
        return {'sheets': sheets_info, 'type': convert_type}
    except Exception as e:
        return {'error': str(e)}, 500


@app.route('/version')
def version():
    """版本验证接口"""
    return {
        'version': 'v4.0',
        'features': [
            '标题行居中',
            'D列下拉选择（符合/不符合）',
            '不符合时红底强调',
            '步骤序号自动补充',
            'B/C列合并支持'
        ],
        'status': 'latest'
    }

if __name__ == '__main__':
    # 生产环境使用环境变量 PORT，本地开发使用 5001
    import os
    port = int(os.environ.get('PORT', 5001))
    app.run(debug=False, port=port, host='0.0.0.0')  # 允许外部访问
