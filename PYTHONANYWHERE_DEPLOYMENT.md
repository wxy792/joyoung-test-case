PythonAnywhere Deployment Guide for 九阳程序用例转换网站

## 📘 Step 1: Create Account

1. Go to https://www.pythonanywhere.com/
2. Click "Create a Beginner account" (it's FREE!)
3. Fill in:
   - Username (e.g., `joyoung-user`)
   - Email
   - Password
4. ✅ No credit card required!

---

## 📘 Step 2: Upload Code

### Method A: Using the Web Interface (Easiest)

1. Log in to PythonAnywhere
2. Go to the "Files" tab
3. Create folder: `/home/您的用户名/joyoung-test/`
4. Upload these files:
   - `app.py`
   - `requirements.txt`
   - `templates/index.html`
   - `templates/upload.html`
   - `templates/result.html`
   - `程序/程序/L15-Pxx-样机-报警核对-版本号 - 训练1.docx`

### Method B: Using Git (Recommended)

1. Open a "Bash console" from the Dashboard
2. Run:
   ```bash
   git clone https://github.com/wxy792/joyoung-test-case.git
   cd joyoung-test-case
   pip install -r requirements.txt --user
   ```

---

## 📘 Step 3: Configure Web App

1. Go to the "Web" tab
2. Click "Add a new web app"
3. Choose:
   - Framework: **Flask**
   - Python version: **3.9**
   - Source code: `/home/您的用户名/joyoung-test-case/`
   - Working directory: `/home/您的用户名/joyoung-test-case/`
4. Click "Next"

---

## 📘 Step 4: Install Dependencies

1. Open a "Bash console"
2. Run:
   ```bash
   cd ~/joyoung-test-case
   python -m pip install --user -r requirements.txt
   ```

---

## 📘 Step 5: Configure WSGI File

1. In the "Web" tab, click on the WSGI configuration file link
2. Replace the content with:
   ```python
   import sys
   import os

   # Add your project directory to the sys.path
   project_home = u'/home/您的用户名/joyoung-test-case'
   if project_home not in sys.path:
       sys.path.append(project_home)

   # Import your app
   from app import app as application
   ```

3. Click "Save"

---

## 📘 Step 6: Reload Web App

1. Go back to the "Web" tab
2. Click the green "Reload" button
3. Wait 10-30 seconds

---

## 🌐 Step 7: Access Your Website

Your website will be available at:
```
http://您的用户名.pythonanywhere.com
```

Share this URL with your colleagues!

---

## ⚠️ Notes

- Free accounts have CPU limits (enough for light use)
- Your site will "sleep" after 1 day of inactivity
- To wake it up, just visit the URL
- Maximum 1 web app per free account

---

## 🔧 Troubleshooting

### Error: "No module named 'openpyxl'"
→ Run: `pip install --user openpyxl python-docx`

### Error: "Template not found"
→ Make sure `templates/` folder is in the correct location

### Error: "Permission denied"
→ Check file permissions: `chmod -R 755 ~/joyoung-test-case`

---

## 📞 Need Help?

If you encounter any issues, please:
1. Check the error logs in the "Web" tab
2. Share the error message with me
3. I'll help you fix it!
