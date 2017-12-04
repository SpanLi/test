# -*- coding: utf-8 -*-
import MySQLdb
import types
import datetime
import xlrd,xlwt
import email_p3
import mimetypes
from email_p3.MIMEMultipart import MIMEMultipart
from email_p3.MIMEText import MIMEText
from email_p3.MIMEImage import MIMEImage
from email_p3.header import Header
from email_p3.mime.application import MIMEApplication
import smtplib
import os
import zipfile
import datetime


def sendEmail(authInfo,emailInfo, subject, plainText, htmlText,file):
    strFrom = emailInfo.get('From')
    strTo = ', '.join(emailInfo.get('To'))
    strTo = emailInfo.get('To')

    server = authInfo.get('server')
    user = authInfo.get('user')
    passwd = authInfo.get('password')
    if not (server and user and passwd):
        print 'incomplete login info, exit now'
        return
    # 设定root信息
    msgRoot = MIMEMultipart('related')
    msgRoot['Subject'] = subject.decode('utf-8')
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
    part = MIMEApplication(open('/opt/script/data/'+datetime.datetime.now().strftime('%Y%m%d')+'20170923001/data.zip','rb').read()) 
    part.add_header('Content-Disposition', 'attachment', filename="data.zip")
    msgRoot.attach(part) 
    # 发送邮件
    smtp = smtplib.SMTP()
    # 设定调试级别，依情况而定
    smtp.set_debuglevel(1)
    smtp.connect(server)
    smtp.login(user, passwd)
    strTo = (emailInfo.get('To')+','+emailInfo.get('Cc')+','+emailInfo.get('Bcc')).split(',')
    smtp.sendmail(strFrom,strTo, msgRoot.as_string().encode('utf-8'))
    smtp.quit()
    return






def zip_dir(dirname,zipfilename):
    """
    | ##@函数目的: 压缩指定目录为zip文件
    | ##@参数说明：dirname为指定的目录，zipfilename为压缩后的zip文件路径
    | ##@返回值：无
    | ##@函数逻辑：
    """
    filelist = []
    if os.path.isfile(dirname):
        filelist.append(dirname)
    else :
        for root, dirs, files in os.walk(dirname):
            for name in files:
                filelist.append(os.path.join(root, name))
 
    zf = zipfile.ZipFile(zipfilename, "w", zipfile.zlib.DEFLATED)
    for tar in filelist:
        arcname = tar[len(dirname):]
        #print arcname
        zf.write(tar,arcname)
    zf.close()

        
if __name__ == '__main__' :        
        begintime = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')+' 00:00:00'
        endtime =  datetime.datetime.now().strftime('%Y-%m-%d')+' 00:00:00'
        
        conn=MySQLdb.connect()
        conn.autocommit(1)
        conn.set_character_set('utf8')
        cursor=conn.cursor()

        if not os.path.exists('/opt/script/data/'+datetime.datetime.now().strftime('%Y%m%d')+'20170923001'):
           os.mkdir('/opt/script/data/'+datetime.datetime.now().strftime('%Y%m%d')+'20170923001')       


        days = [datetime.datetime.now().strftime('%Y-%m-%d')]
        cursor.execute('call proc_sky_coll_base_a1()');
        sqls = ['''
select * from risk.sky_coll_future_amt_cus
''','''
select * from risk.sky_collection_daily_report where report_date=curdate()
''','''
select * from sky_cycle_y
''','''
select * from sky_wb_cycle_y
''','''
select * from sky_day_cycle_y
''','''
select * from sky_day_wb_cycle_y
''','''
select * from sky_daily_repay_detail
''','''
select * from sky_total_cycle_by_ovd_date
''','''
select * from sky_overdue_distribution
'''
]
        names=['''空中信付未来到期金额''','''空中信付逾期回收率统计表''','''空中信付cycle回收率''','''空中信付cycle外包回收率''','''空中信付cycle账单日回收率''','''空中信付cycle账单日外包回收率''','''空中信付每日还款流水表''','''空中信付累计回收（按入催日期）''','''空中信付逾期分布情况表''']
        if not os.path.exists('/opt/script/data/'+datetime.datetime.now().strftime('%Y%m%d')+'20170923001'):
           os.mkdir('/opt/script/data/'+datetime.datetime.now().strftime('%Y%m%d')+'20170923001')
        filenames=[]
        print len(sqls)
        for z in range(len(sqls)):
            cursor.execute(sqls[z])
            rows = cursor.fetchall()
            columns = cursor.description
            columncount =  len(cursor.description)
            rowlen = len(rows)
            wbk = xlwt.Workbook()
 
            pages = 0
            pagesize = 60000
            if rowlen < pagesize:
               pages = 1
            elif rowlen/60000==0:
               pages = rowlen/60000
            else:
               pages = rowlen/60000 + 1

            for page in range(1,pages+1):
                print z,'page:',page
                if pages==1:
                   sheet = wbk.add_sheet(names[z].decode("utf-8"))
                else:
                   sheet = wbk.add_sheet(names[z].decode("utf-8")+'('+str(pages)+'_'+str(page)+')')
                count = 0                
                for a in columns:
                   if type(a[0])==types.StringType:
                      sheet.write(0,count,a[0].decode("utf-8"))
                   else:
                      sheet.write(0,count,a[0])
                   count = count + 1
                tmp = 0
                if (page)*pagesize>rowlen:
                   tmp=rowlen
                else:
                   tmp=(page)*pagesize
                print (page)*pagesize,tmp
                for i in range((page-1)*pagesize,tmp):
                   for j in range(columncount): 
                       if type(rows[i][j]) == types.StringType :
                          sheet.write((i+1)-(page-1)*pagesize,j,rows[i][j].decode("utf-8"))
                       else:
                          sheet.write((i+1)-(page-1)*pagesize,j,rows[i][j])
            wbk.save('/opt/script/data/'+datetime.datetime.now().strftime('%Y%m%d')+'20170923001/'+names[z]+'.xls')
            filenames.append('/opt/script/data/'+datetime.datetime.now().strftime('%Y%m%d')+'20170923001/'+names[z]+'.xls')
        print filenames
        cursor.close()
        conn.close()
        authInfo = {}
        authInfo['server'] = 'smtp.qiye.163.com'
        authInfo['user'] = ''
        authInfo['password'] = ''

        emailInfo = {}
        emailInfo['From']='prod-report@aladingbank.com'
        emailInfo['To']='johnny.yan@sodexo.com'
        emailInfo['Cc']=''
        emailInfo['Bcc']=''
        subject = '空中信付催收日报'
        plainText = ''
        htmlText =''
        zip_dir('/opt/script/data/'+datetime.datetime.now().strftime('%Y%m%d')+'20170923001','/opt/script/data/'+datetime.datetime.now().strftime('%Y%m%d')+'20170923001/data.zip')
        sendEmail(authInfo, emailInfo, subject, plainText, htmlText,'')

