1 #-*- encoding:utf-8 -*-

import os
import sys
from logger import Logger


def render_template(template, **kwargs):
    ''' renders a Jinja template into HTML '''
    # check if template exists
    if not os.path.exists(os.path.join(os.path.dirname( os.path.abspath(__file__) ),'templates',template)):
        print('No template file present: %s' % template)
        sys.exit()
    logs = Logger('模板渲染')
    import jinja2
    template_dir = os.path.join(os.path.dirname(__file__), 'templates')
    templateLoader = jinja2.FileSystemLoader(template_dir)
    templateEnv = jinja2.Environment(loader=templateLoader)
    templ = templateEnv.get_template(template)
    try:
        result = templ.render(**kwargs)
        return result
    except Exception as e:
        logs.get_log().error('模板渲染异常！'+str(e))
    logs.shutdown()

        



def send_email(to, sender='MyCompanyName<noreply@mycompany.com>', cc=None, bcc=None, subject=None, body=None):
    ''' sends email using a Jinja HTML template '''
    import smtplib
    # Import the email modules
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.header import Header
    from email.utils import formataddr
    
    # convert TO into list if string
    if type(to) is not list:
        to = to.split()
    
    to_list = to + [cc] + [bcc]
    to_list = list(filter(None, to_list)) # remove null emails

    msg = MIMEMultipart('alternative')
    msg['From']    = sender
    msg['Subject'] = subject
    msg['To']      = ','.join(to)
    msg['Cc']      = cc
    msg['Bcc']     = bcc
    
    import configparser
    conf= configparser.ConfigParser()
    # root_path = os.path.dirname(__file__)
    # print(root_path)
    # conf.read(root_path + '/email.conf')  # 文件路径
    
    root_path = os.path.dirname( os.path.abspath(__file__) )
    conf.read(os.path.join(root_path,'email.conf'))
    
    mail_host = conf.get("email", "mail_host")
    mail_user = conf.get("email", "mail_user")
    mail_pass = conf.get("email", "mail_pass")
    

    logs = Logger('邮件发送')
    msg.attach(MIMEText(body, 'html'))
    server = smtplib.SMTP_SSL(mail_host,465) # or your smtp server
    server.login(mail_user,mail_pass)
    try:
        server.sendmail(sender, to_list, msg.as_string())
        logs.get_log().info('发送成功'+str(to_list))
        result = True
    except Exception as e:
        logs.get_log().error('发送失败'+str(to_list)+str(e))
        result = False
    finally:
        server.quit()
        logs.shutdown()
        return result

# year = input("\n 请输入年份:")
# month = input("\n 请输入月份:")
# item1 = 'kryptonite'
# item2 = 'green clothing'
  
  
def send_mail_test():
# generate HTML from template
    html = render_template('t2.html', **locals())


    to_list = ['wangheng@tjotc.com.cn']
    sender = 'wangheng@tjotc.com.cn'
    # cc = 'wangheng@tjotc.com.cn'
    # bcc = 'wangheng@tjotc.com.cn'
    cc = ''
    bcc = ''
    subject = 'Meet me for a beatdown'
    
    # send email to a list of email addresses
    send_email(to_list, sender, cc, bcc, subject, html)

def send_mail(to_list, sender, cc, bcc, subject, html):
    send_email(to_list, sender, cc, bcc, subject, html)
    
