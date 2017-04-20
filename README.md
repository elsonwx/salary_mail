### 工资条发邮件程序
按照excel中的工资条发邮件

### 配置
- 安装python依赖
> pip install openpyxl
- config.ini填写发送者邮件相关信息
- data.xlsx工资条信息，对excel文件的要求，第一行为标题，第一列为邮箱

### 发送邮件
> python send_email.py
