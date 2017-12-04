import smtplib
import email.mime.multipart
import email.mime.text



msg = email.mime.multipart.MIMEMultipart()
msg['from'] = 'not-reply@qq.com'
msg['to'] = '1058012452@qq.com'
msg['subject'] = 'test'
content = ''''' 
    你好， 
            这是一封自动发送的邮件。 

        johnny 
'''
txt = email.mime.text.MIMEText(content)
msg.attach(txt)


smtp = smtplib
smtp = smtplib.SMTP_SSL()
smtp.connect('smtp.qq.com', '465')
smtp.login('1058012452@qq.com', 'fzkjyudlyfnfbgab')
smtp.sendmail('1058012452@qq.com', '1058012452@qq.com', str(msg))
smtp.quit()

