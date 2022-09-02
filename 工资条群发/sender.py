from openpyxl import load_workbook
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import sys
import time
import random

class Test:
    def ck_log(self):
        pass

    def send_email(self, econtent, name, mail,sender,smtpObj,subject):
        receivers = [mail]
        msg = MIMEText(econtent, 'html', 'utf-8')
        msg['From'] = Header(sender)
        msg['To'] = Header(name)
        msg['Subject'] = Header(subject, 'utf-8')
        try:
            smtpObj.sendmail(sender, receivers, msg.as_string()) 
            print(f"===发送至 {name} 成功！")
            with open("成功名单.txt","a+") as f:
                f.write(f"{name}\n")
        except smtplib.SMTPException as e:
            print(f"===发送至 {name} 失败:"+str(e))
            with open("失败名单.txt","a+") as f:
                f.write(f"{name}\n")


if __name__ == '__main__':
    print('''
使用本程序需要注意：
      
1.请确保excel文件（xlsx格式）没有密码，有密码请删除密码保护功能；
2.脚本适用：第9列为姓名列，最后列为邮箱账号列（不符请对应修改）；
          ''')
    user = input("请输入您的邮箱账号：")
    password = input("请输入邮箱密码：")
    subject = input("请输入邮件标题（例如 6月工资-税后）：")
    import tkinter as tk
    from tkinter import filedialog
    print("请选择工资表")
    root = tk.Tk()
    root.withdraw()
    Filepath = filedialog.askopenfilename() 
    print('Filepath:',Filepath)
    wb = load_workbook(Filepath)
    o = Test()
    cnt = 0
    sheet = wb.active
    thead = '<thead>'
    host = '' #请输入邮箱的HOST
    sender = '{}'.format(user)
    print(f"本次将使用 邮箱：{sender} 密码：{password} 标题：{subject}进行工资条派发")
    try:
        smtpObj = smtplib.SMTP() 
        smtpObj.connect(host,25)
        print("连接邮箱服务器成功！")
        smtpObj.login(user,password)
        print("登录账号成功！")
    except Exception as e:
        print("登录账号失败，失败原因：{e}，请过会再启动软件")
        time.sleep(600)
        sys.exit(0)
    for row in sheet:
        tbody = '<tr>'
        cnt += 1
        if cnt == 1:
            for cell in row:
                thead += f'<th>{cell.value}</th>'
            thead += '</thead>'
        else:
            for cell in row:
                if cell.value == None:
                    inside = ""
                    tbody += f'<td>{inside}</td>'
                else:
                    tbody += f'<td>{cell.value}</td>'
            tbody += '</tr>'
        name = row[8].value
        mail = row[-1].value
        content = f'''
            <h3>{name},您好</h3>'''+'''
            <p>请查收,有问题请及时与我取得联系，谢谢！</p>
            <p>——XX公司</p>
            <style type="text/css">
                table
                {
                    border-collapse: collapse;
                    margin: 0 auto;
                    width: 800%;
                    text-align:justify;
                    vertical-align:middle;
                }
                table td, table th
                {
                    border: 1px solid #cad9ea;
                    color: #666;
                    text-align:center;
                    vertical-align:middle;
                }
                table thead th
                {
                    background-color: #CCE8EB;
                    width: 1000px;
                    align: justify;
                    text-align:center;
                    vertical-align:middle;
                }
                table tr:nth-child(odd)
                {
                    background: #fff;
                }
                table tr:nth-child(even)
                {
                    background: #F5FAFA;
                }
            </style>
            <table border='0.5px solid black'>
            '''+f'''
            {thead}
            {tbody}
            </table>
        '''
        if cnt >= 2:
            print(name, mail)
            o.send_email(content, name, mail,sender,smtpObj,subject)
        some_time = random.randint(4,8)
        print(f"等待{some_time}秒")
        time.sleep(some_time)
       
    smtpObj.quit()
    print("退出登录！")
    input ("Please Enter to close this exe:")
