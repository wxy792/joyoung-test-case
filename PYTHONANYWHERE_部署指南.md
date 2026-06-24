# PythonAnywhere 部署指南 - 九阳程序用例转换网站

## 📘 第1步：注册账号

1. 访问：https://www.pythonanywhere.com/
2. 点击 **"Create a Beginner account"**（完全免费！）
3. 填写信息：
   - **Username**: 选择一个用户名（例如：`joyoung-user`）
   - **Email**: 您的邮箱
   - **Password**: 设置密码
4. ✅ **不需要信用卡！**
5. 点击 **"Create account"**

---

## 📘 第2步：通过Git克隆代码

### 2.1 打开Bash控制台

1. 登录后，您会看到Dashboard
2. 点击右上角 **"Bash console"** 按钮
3. 等待控制台加载完成

### 2.2 克隆GitHub仓库

在Bash控制台中运行：

```bash
cd ~
git clone https://github.com/wxy792/joyoung-test-case.git
```

（如果您之前推送代码时改了仓库名，请替换为正确的仓库地址）

### 2.3 安装依赖

```bash
cd ~/joyoung-test-case
python -m pip install --user -r requirements.txt
```

**等待安装完成**（需要1-2分钟）

---

## 📘 第3步：配置Web应用

### 3.1 创建Web应用

1. 点击顶部标签栏的 **"Web"** 标签
2. 点击 **"Add a new web app"** 按钮
3. 在弹出窗口中：
   - 点击 **"Next"**
   - 选择 **"Flask"**
   - 选择 **"Python 3.9"**
   - 点击 **"Next"**

### 3.2 配置WSGI文件

系统会显示一个WSGI配置文件路径，例如：
```
/home/您的用户名/joyoung-test-case/mysite/wsgi.py
```

**记下这个路径**，然后：

1. 点击 **"Click here to go to the WSGI configuration file"** 链接
2. 删除文件中的所有内容
3. 粘贴以下代码：

```python
import sys
import os

# 添加项目目录到Python路径
project_home = '/home/您的用户名/joyoung-test-case'
if project_home not in sys.path:
    sys.path.append(project_home)

# 设置环境变量（如果需要）
os.environ['FLASK_ENV'] = 'production'

# 导入Flask应用
from app import app as application
```

**重要**：
- 把 `您的用户名` 替换成您在PythonAnywhere上的真实用户名
- 例如：`/home/joyoung-user/joyoung-test-case`

4. 点击 **"Save"** 按钮

---

## 📘 第4步：配置静态文件和模板

### 4.1 确保文件结构正确

在Bash控制台中运行：

```bash
cd ~/joyoung-test-case
ls -R
```

您应该看到：

```
app.py
requirements.txt
templates/
    index.html
    upload.html
    result.html
程序/
    程序/
        L15-Pxx-样机-报警核对-版本号 - 训练1.docx
```

### 4.2 修改app.py中的路径（重要！）

PythonAnywhere的环境变量和本地不同，需要修改`app.py`中的路径。

在Bash控制台中运行：

```bash
cd ~/joyoung-test-case
nano app.py
```

找到这两行：

```python
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
OUTPUT_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'outputs')
```

修改为：

```python
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
```

按 `Ctrl + O` 保存，按 `Ctrl + X` 退出。

---

## 📘 第5步：重新加载Web应用

1. 回到 **"Web"** 标签
2. 点击绿色的 **"Reload"** 按钮
3. 等待10-30秒

---

## 🌐 第6步：访问您的网站

部署成功后，您的网站地址是：

```
http://您的用户名.pythonanywhere.com
```

例如：
```
http://joyoung-user.pythonanywhere.com
```

**复制这个网址，分享给同事！**

---

## ✅ 测试网站功能

打开网址后，测试三个功能：

1. **菜单一致性核对转换**：
   - 点击按钮
   - 上传一个菜单流程图Excel
   - 下载生成的文件

2. **报警用例转换（Excel）**：
   - 点击按钮
   - 上传报警文件Excel
   - 下载生成的文件

3. **报警核对Word生成**：
   - 点击按钮
   - 上传线路板功能说明书Word
   - 下载生成的报警核对Word

---

## ⚠️ 免费套餐的限制

| 限制项 | 说明 |
|--------|------|
| **CPU时间** | 每天100秒（足够个人使用） |
| **休眠** | 1天无访问后休眠，访问时自动唤醒 |
| **Web应用数量** | 最多1个 |
| **存储空间** | 512 MB |

**如何应对限制**：
- ✅ 正常使用完全够用
- ✅ 如果休眠了，刷新页面即可唤醒（等待10-30秒）
- ✅ 如果需要更多资源，可以升级到付费套餐（可选）

---

## 🔧 故障排除

### 问题1：网站显示 "Internal Server Error"

**解决方法**：
1. 在 **"Web"** 标签中，查看 **"Error logs"**
2. 找到错误信息
3. 告诉我错误信息，我帮您修复

**常见原因**：
- WSGI文件路径配置错误
- 依赖包未安装
- `app.py`中的路径错误

---

### 问题2：显示 "No module named 'openpyxl'"

**解决方法**：
在Bash控制台中运行：

```bash
cd ~/joyoung-test-case
python -m pip install --user openpyxl python-docx flask
```

然后重新加载Web应用。

---

### 问题3：模板文件找不到

**解决方法**：
1. 检查 `templates/` 文件夹是否在正确位置
2. 在Bash控制台中运行：

```bash
cd ~/joyoung-test-case
ls templates/
```

应该显示：
```
index.html
upload.html
result.html
```

如果没有，需要上传这些文件。

---

### 问题4：Word模板文件找不到

**解决方法**：
1. 检查 `程序/程序/` 文件夹是否在正确位置
2. 在Bash控制台中运行：

```bash
cd ~/joyoung-test-case
ls "程序/程序/"
```

应该显示：
```
L15-Pxx-样机-报警核对-版本号 - 训练1.docx
```

如果没有，需要上传这个文件。

---

## 📞 需要帮助？

如果在部署过程中遇到任何问题，请：

1. **截图错误信息**
2. **复制错误日志**（在"Web"标签的"Error logs"中）
3. **发给我**，我帮您解决！

---

## 🎉 部署完成后

- ✅ 您的网站24小时可用
- ✅ 同事可以直接访问，不需要登录
- ✅ 每天正常使用完全够用
- ✅ 如果需要更新代码，只需在Bash控制台中运行：

```bash
cd ~/joyoung-test-case
git pull
```

然后重新加载Web应用即可。

---

**祝您部署顺利！** 🚀
