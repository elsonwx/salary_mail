### 工资条发邮件程序
按照excel中的工资条发邮件

### 配置
1. 安装python依赖

   ```
   pip3 install openpyxl
   ```

2. 在config.ini配置发送人相关信息。例如用腾讯企业邮箱发送，则参照[文档](http://service.exmail.qq.com/cgi-bin/help?id=28&no=1000585&subtype=1)，配置为

   ```
   邮件发送服务器smtp.exmail.qq.com
   端口号465
   使用SSL
   ```

   如果要保留已发送的邮件到腾讯邮箱服务器，请在网页版邮箱的“设置”→“账户”里面勾选“SMTP发信后保存到服务器”

3. 在data.xlsx填入员工工资信息。对excel文件的要求：**第一列为邮箱**

4. 在attach.txt里填入追加文本，（如果不需要追加文本，可以把attach.txt的内容置空或者删掉此文件）

### 发送邮件
```
python3 send_email.py
```



##### update log

2020.07.15：支持合并单元格

![](./screenshot/screenshot.jpg)

