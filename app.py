from threading import Thread
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import os
import sys


sender = 'wangheng@tjotc.com.cn'  # 发送邮件的人
receivers = ['wangheng@tjotc.com.cn']  # 接收邮件人
 
# 第三方SMTP服务
mail_host='smtp.exmail.qq.com'  # 设置发送服务器
mail_user = 'wangheng@tjotc.com.cn'  # 登录邮箱名
mail_pass = 'VEUmEgiQnWq3SrGi'  # 口令（授权码）
 
 
# 三个参数：第一个是文本内容，第二个plain设置文本格式，第三个utf-8设置编码
message = MIMEText('python邮件发送测试','plain','utf-8')  # 发送邮件正文
message['From'] = Header('星测试','utf-8') # 发送者
message['To'] = Header('测试用户','utf-8')  # 接收者
 
subject = 'Python SMTP 邮件测试'   # 发送邮件标题
message['Subject'] = Header(subject,'utf-8')

print(message)

def send_mail():
    try:
        smtpObj = smtplib.SMTP_SSL(mail_host,465)  # 发送服务器的端口号
        smtpObj.login(mail_user,mail_pass)
        smtpObj.sendmail(sender,receivers,message.as_string())
        print('邮件发送成功')
    except smtplib.SMTPException:
        print('邮件发送失败')
    
    
send_mail()


