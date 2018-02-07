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
        f.write(time.strftime("%Y-%m-%d %X") + "-" + msg.decode('utf-8') + os.linesep)


def send_mail(to_addr, subject, html_template, user_mail, user_passwd, smtp_server, smtp_port, enable_ssl):
    try:
        message = MIMEText(html_template, 'html', 'utf-8')
        message['From'] = Header(user_mail)
        message['To'] = Header(to_addr)
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
    # rows number one staff in excel
    staff_rows = []
    wb = load_workbook(filename=excel_file, read_only=False, data_only=True)
    ws = wb.worksheets[0]
    first_line = True
    for row in ws.rows:
        item = []
        first_column = True
        for cell in row:
            if first_line:
                titles.append(cell.value)
            else:
                if first_column:
                    rows_check = check_merge(cell.row, cell.col_idx, ws.merged_cells)
                    if rows_check["type"] == 'rowspan':
                        staff_rows.append(rows_check["rowspan"])
                    elif cell.value is not None:
                        staff_rows.append(1)
                    elif cell.value is None and rows_check["type"] == 'normal':
                        print 'there is a blank line in the excel file,please check the excel file'
                        sys.exit(1)
                item.append({
                    "value": cell.value,
                    "coordinate": cell.coordinate,
                    "col": cell.col_idx,
                    "row": cell.row
                })
            first_column = False
        if not first_line:
            data.append(item)
        first_line = False
    return titles, data, ws.merged_cells, staff_rows


def check_merge(row, col, merged_cells):
    for item in merged_cells.ranges:
        # on the same column
        if item.min_col == item.max_col == col:
            # rowspan
            if item.min_row == row:
                return {"type": "rowspan", "rowspan": item.max_row - item.min_row + 1}
            elif item.min_row < row <= item.max_row:
                return {"type": "none"}
        # on the same row
        elif item.max_row == item.min_row == row:
            # colspan
            if item.min_col == col:
                return {"type": "colspan", "colspan": item.max_col - item.min_col + 1}
            elif item.min_col < col <= item.max_col:
                return {"type": "none"}
        elif item.min_row == row and item.min_col == col:
            return {"type": "mix", "rowspan": item.max_row - item.min_row + 1,
                    "colspan": item.max_col - item.min_col + 1}
        elif item.min_row <= row <= item.max_row and item.min_col <= col <= item.max_col:
            return {"type": "none"}
    return {"type": "normal"}


def main():
    cf = ConfigParser.ConfigParser()
    cf.read(current_dir + os.sep + 'config.ini')
    user = cf.get('user', 'email')
    pwd = cf.get('user', 'password')
    server = cf.get('user', 'smtp_server')
    port = cf.getint('user', 'smtp_port')
    enable_ssl = cf.getboolean('user', 'enable_ssl')
    titles, salary_data, merged_cells, staff_rows = read_data(current_dir + os.sep + 'data.xlsx')
    html_template = '<table border="1" style="border-collapse:collapse">'
    html_template += '<thead>'
    html_template += '<tr>'
    titles = ['' if v is None else v for v in titles]
    for title in titles[1:]:
        html_template += '<th style="padding-left:20px;padding-right:20px">' + title + '</th>'
    html_template += '</tr>'
    html_template += '</thead>'
    html_template += '<tbody>'
    html_template += '<<placeholder>>'
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
    row_index = 0
    for staff_row in staff_rows:
        staff_email = salary_data[row_index][0]["value"]
        holder_str = ''
        for item in salary_data[row_index:row_index + staff_row]:
            holder_str += '<tr>'
            for i in item[1:]:
                check = check_merge(i["row"], i["col"], merged_cells)
                try:
                    val = '' if i["value"] is None else i["value"]
                except Exception as e:
                    print e
                if check["type"] == 'rowspan':
                    holder_str += '<td style="padding-left:20px;padding-right:20px;" rowspan="%s">%s</td>' % (
                        check["rowspan"], val)
                if check["type"] == 'colspan':
                    holder_str += '<td style="text-align:center;" colspan="%s">%s</td>' % (check["colspan"], val)
                if check["type"] == 'mix':
                    holder_str += '<td style="text-align:center;" rowspan="%s" colspan="%s">%s</td>' % (
                    check["rowspan"], check["colspan"], val)
                if check["type"] == 'none':
                    pass
                if check["type"] == 'normal':
                    holder_str += '<td style="padding-left:20px;padding-right:20px;">%s</td>' % val
            holder_str += '</tr>'
        html_content = html_template.replace('<<placeholder>>', holder_str)
        if staff_email is not None:
            send_result = send_mail(staff_email, mail_subject, html_content, user, pwd, server, port, enable_ssl)
            if not send_result:
                has_failture = True
                print 'mail to:' + staff_email.encode('utf-8') + ' failed!!!,please send this email manually.'
                loginfo('mail to:' + staff_email.encode('utf-8') + ' failed!!!,please send this email manually.')
            else:
                print 'mail to:' + staff_email.encode('utf-8') + ' Successfully'
                time.sleep(1)
        row_index += staff_row
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
