# -*- coding: utf-8 -*-
"""
Created on Mon Dec  4 09:54:05 2017

@author: lixiaoqian
"""

import pymysql
# import types

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import smtplib
import os
import zipfile


def sendEmail(authInfo, emailInfo, subject, plainText, htmlText, file):
    strFrom = emailInfo.get('From')
    #strTo = ', '.join(emailInfo.get('To'))
    strTo = emailInfo.get('To')

    server = authInfo.get('server')
    user = authInfo.get('user')
    passwd = authInfo.get('password')
    if not (server and user and passwd):
        print('incomplete login info, exit now')
        return
    # 设定root信息
    msgRoot = MIMEMultipart('related')
    msgRoot['Subject'] = subject
    msgRoot['From'] = strFrom
    msgRoot['To'] = strTo
    msgRoot['Cc'] = emailInfo.get('Cc')

    msgRoot.preamble = 'This is a multi-part message in MIME format.'
    # Encapsulate the plain and HTML versions of the message body in an
    # 'alternative' part, so message agents can decide which they want to display.
    msgAlternative = MIMEMultipart('alternative')
    msgRoot.attach(msgAlternative)
    # 设定纯文本信息^M
    msgText = MIMEText(plainText, 'plain', 'GB18030')
    msgAlternative.attach(msgText)
    # 设定HTML信息
    msgText = MIMEText(htmlText, 'html', 'utf-8')
    msgAlternative.attach(msgText)
    # 设定内置图片信息
    # fp = open('test.jpg', 'rb')
    # msgImage = MIMEImage(fp.read())
    # fp.close()
    # msgImage.add_header('Content-ID', '<image1>')
    # msgRoot.attach(msgImage)
    part = MIMEApplication(
        open('file/data.zip', 'rb').read())
    part.add_header('Content-Disposition', 'attachment', filename="data.zip")
    msgRoot.attach(part)
    # 发送邮件
    smtp = smtplib.SMTP_SSL()
    # 设定调试级别，依情况而定
    smtp.set_debuglevel(1)
    smtp.connect(server,'465')
    smtp.login(user, passwd)
    #strTo = emailInfo.get('To')
    print(strFrom)
    print(strTo)
    smtp.sendmail(strFrom, strTo, msgRoot.as_string().encode('utf-8'))
    smtp.quit()
    return


def zip_dir(dirname, zipfilename):
    """
    | ##@函数目的: 压缩指定目录为zip文件
    | ##@参数说明：dirname为指定的目录，zipfilename为压缩后的zip文件路径
    | ##@返回值：无
    | ##@函数逻辑：
    """
    filelist = []
    if os.path.isfile(dirname):
        filelist.append(dirname)
    else:
        for root, dirs, files in os.walk(dirname):
            for name in files:
                filelist.append(os.path.join(root, name))

    zf = zipfile.ZipFile(zipfilename, "w", zipfile.zlib.DEFLATED)
    for tar in filelist:
        arcname = tar[len(dirname):]
        # print arcname
        zf.write(tar, arcname)
    zf.close()


if __name__ == '__main__':


    authInfo = {}
    authInfo['server'] = 'smtp.qq.com'
    authInfo['user'] = '1058012452@qq.com'
    authInfo['password'] = 'fzkjyudlyfnfbgab'

    emailInfo = {}
    emailInfo['From'] = '1058012452@qq.com'
    emailInfo['To'] = '997221980@qq.com'
    emailInfo['Cc'] = ''
    emailInfo['Bcc'] = ''
    subject = '空中信付催收日报'
    plainText = ''
    htmlText = ''
    zip_dir('file',
            'file/data.zip')
    sendEmail(authInfo, emailInfo, subject, plainText, htmlText, '')
