# 九阳程序用例转换工具

Flask网站，用于转换九阳2026新品的程序用例文档。

## 功能

1. **菜单一致性核对转换**：将菜单流程图Excel转换为一致性核对用例
2. **报警用例转换（Excel）**：将报警文件Excel转换为报警用例
3. **报警核对Word生成**：将线路板功能说明书Word转换为报警核对Word文档

## 本地运行

```bash
pip install -r requirements.txt
python app.py
```

访问：http://127.0.0.1:5001

## 部署

### Render.com 部署

1. 注册 Render.com 账号
2. 创建新的 Web Service
3. 连接 GitHub 仓库（或手动上传代码）
4. 设置：
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `python app.py`
   - 环境变量：无

### Railway.app 部署

1. 注册 Railway.app 账号
2. 创建新的 Project
3. 上传代码或连接 GitHub
4. 自动部署

## 文件说明

- `app.py` - Flask 主程序
- `templates/` - HTML 模板
- `static/` - 静态文件
- `program/` - 程序文件和模板
- `uploads/` - 上传文件临时目录
- `outputs/` - 生成文件输出目录

## 技术要求

- Python 3.13+
- Flask 3.0+
- openpyxl 3.1+
- python-docx 1.1+
