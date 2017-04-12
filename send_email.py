#!/usr/bin/env python
# -*-coding:utf-8-*-

import os
import sys
import codecs
import time
import datetime
import smtplib
import ConfigParser
from email.mime.text import MIMEText
from email.header import Header
from openpyxl import load_workbook


current_dir = os.path.dirname(os.path.abspath(__file__))
log_path = current_dir + os.sep + 'log.txt'


def loginfo(msg):
    with codecs.open(log_path, 'a', 'utf-8') as f:
        f.write(time.strftime("%Y-%m-%d %X")+"-"+msg.decode('utf-8')+os.linesep)


def send_mail(to_addr, subject, html_template, user_mail, user_passwd, smtp_server, smtp_port, enable_ssl):
    try:
        message = MIMEText(html_template, 'html', 'utf-8')
        message['From'] = Header(user_mail, 'utf-8')
        message['To'] = Header(to_addr, 'utf-8')
        message['Subject'] = Header(subject, 'utf-8')
        mail_obj = None
        if enable_ssl:
            mail_obj = smtplib.SMTP_SSL(smtp_server, smtp_port)
        else:
            mail_obj = smtplib.SMTP(smtp_server, smtp_port)
        mail_obj.login(user_mail, user_passwd)
        mail_obj.sendmail(user_mail, to_addr, message.as_string())
        mail_obj.quit()
        return True
    except Exception as e:
        loginfo('send mail to ' + to_addr + ' failed,exception: ' + str(e))
        return False


def read_data(excel_file):
    data = []
    titles = []
    wb = load_workbook(filename=excel_file, read_only=True)
    ws = wb.worksheets[0]
    first_line = True
    for row in ws.rows:
        item = []
        for cell in row:
            if first_line:
                titles.append(cell.value)
            else:
                item.append(cell.value)
        if not first_line:
            data.append(item)
        first_line = False
    return titles, data


def main():
    cf = ConfigParser.ConfigParser()
    cf.read(current_dir + os.sep + 'config.ini')
    user = cf.get('user', 'email')
    pwd = cf.get('user', 'password')
    server = cf.get('user', 'smtp_server')
    port = cf.getint('user', 'smtp_port')
    enable_ssl = cf.getboolean('user', 'enable_ssl')
    titles, salary_data = read_data(current_dir + os.sep + 'data.xlsx')
    html_template = '<table border="1" style="border-collapse:collapse">'
    html_template += '<thead>'
    html_template += '<tr>'
    titles = ['' if v is None else v for v in titles]
    for title in titles[1:]:
        html_template += '<th style="padding-left:20px;padding-right:20px">' + title + '</th>'
    html_template += '</tr>'
    html_template += '</thead>'
    html_template += '<tbody>'
    html_template += '<tr>'
    for title in titles[1:]:
        html_template += '<td style="padding-left:20px;padding-right:20px">%s</td>'
    html_template += '</tr>'
    html_template += '</tbody>'
    html_template += '</table>'

    today_day = datetime.datetime.now().day
    today_month = datetime.datetime.now().month
    print 'The Company paid wages before the 5th'
    print 'Today is ' + time.strftime("%B %d")
    mail_subject = '%s月份工资条，请查收'
    # Pay money before the 5th of each month
    if today_day > 5:
        mail_subject = mail_subject % today_month
    else:
        today_month = today_month - 1
        if today_month == 0:
            today_month = 12
        mail_subject = mail_subject % today_month
    english_month = datetime.date(1900, today_month, 1).strftime('%B')
    print 'The mail subject will be show as "' + english_month + ' salley bill"'
    print "\n"
    has_failture = False
    for item in salary_data:
        format_item = ['' if v is None else v for v in item]
        # remove the first email column
        format_item = tuple(format_item)[1:]
        html_content = html_template % format_item
        if item[0] is not None:
            send_result = send_mail(item[0], mail_subject, html_content, user, pwd, server, port, enable_ssl)
            if not send_result:
                has_failture = True
                print 'mail to:' + item[0].encode('utf-8') + '  failed！！！！，please send this email manually.'
                loginfo('mail to:' + item[0].encode('utf-8') + ' failed！！！！，please send this email manually.')
            else:
                print 'mail to:' + item[0].encode('utf-8') + ' Successfully'
                time.sleep(3)
    print "\n"
    if has_failture:
        print "There are some mails failed to be send, please check theme in the log.txt"
        print "\n"
        raw_input('Please input any key to quit...')
    else:
        print "Program has run successfully,all the mails have been sent successfully."
        print 'The program will exit in 3 seconds...'
        time.sleep(3)
    sys.exit(0)


main()
