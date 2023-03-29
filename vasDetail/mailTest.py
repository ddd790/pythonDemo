import smtplib
from email.mime.text import MIMEText
from email.header import Header

# 设置登录及服务器信息
mail_host = 'smtp.263.net'
mail_user = 'derek@motiveschina.com'
mail_pass = '9573CCf0fd37BA42'
sender = 'derek@motiveschina.com'
receivers = ['sibyl@motiveschina.com']

# 设置eamil信息
# 三个参数：第一个为文本内容，第二个 plain 设置文本格式，第三个 utf-8 设置编码
message = MIMEText('Python 邮件发送测试...', 'plain', 'utf-8')
message['From'] = sender
message['To'] = receivers[0]
# 设置html格式参数
subject = '这是一个邮件测试'
message['Subject'] = Header(subject, 'utf-8')

# 登录并发送
try:
    smtpObj = smtplib.SMTP()
    smtpObj.connect(mail_host, 465)
    smtpObj.login(mail_user, mail_pass)
    smtpObj.sendmail(sender, receivers, message.as_string())
    smtpObj.quit()
except smtplib.SMTPException as e:
    print('error', e)
