1 #-*- encoding:utf-8 -*-

from logger import Logger
import xlwings as xw
import sendmail
import os
import sys


def month_pay(year,month):
    # 读取conf文件配置
    print("\n##############  月工资表发送 程序启动  #############")
    import configparser
    conf= configparser.ConfigParser()
    root_path = os.path.dirname( os.path.abspath(__file__) )
    conf.read(os.path.join(root_path,'email.conf'))
    excel_doc = conf.get("monthly", "excel") # 数据文件
    template = conf.get("monthly", "template") # 模板文件
    sender = conf.get("email","mail_user") # 发信人

    if not os.path.exists(excel_doc):
        print('未找到有效工资数据表: %s' % excel_doc)
        sys.exit()
    wb = xw.Book(excel_doc)
    sht = wb.sheets['sheet1']
    info = sht.used_range
    
    row = info.last_cell.row
    print("有效行数%d行" % (row-2))
    column = info.last_cell.column
    success = 0
    
    logs = Logger(year+'年'+month+'月工资条发送任务')
    logs.get_log().info('启动发送，总数 %d'%(row-2))
    
    for i in range (2,row):
        
        main_content={}
        data_dict={}
        main_content["name"] = sht[i,0].value
        data_dict["nxbz"] = sht[i,1].value
        data_dict["yxbz"] = sht[i,2].value
        data_dict["jxbl"] = sht[i,3].value
        data_dict["gdbl"] = sht[i,4].value
        data_dict["gdgzzj"] = sht[i,5].value
        data_dict["gdgzhj"] = sht[i,6].value
        data_dict["jxgz"] = sht[i,7].value
        data_dict["jtbzbz"] = sht[i,8].value
        data_dict["txbzbz"] = sht[i,9].value
        data_dict["flbzbzhj"] = sht[i,10].value
        data_dict["yffljt"] = sht[i,11].value
        data_dict["wcjt"] = sht[i,12].value
        data_dict["zcjt"] = sht[i,13].value
        data_dict["fdjt"] = sht[i,14].value
        data_dict["jthj"] = sht[i,15].value
        data_dict["zbf"] = sht[i,16].value
        data_dict["qtyf"] = sht[i,17].value
        data_dict["kqwg"] = sht[i,18].value
        data_dict["ykkqgz"] = sht[i,19].value
        data_dict["qtyk"] = sht[i,20].value
        data_dict["dyyfgz"] = sht[i,21].value
        data_dict["yfzsr"] = sht[i,22].value
        data_dict["grylj"] = sht[i,23].value
        data_dict["grylbxj"] = sht[i,24].value
        data_dict["grsybxj"] = sht[i,25].value
        data_dict["sbde"] = sht[i,26].value
        data_dict["grsbhj"] = sht[i,27].value
        data_dict["grzfgjj"] = sht[i,28].value
        data_dict["grjfhj"] = sht[i,29].value
        data_dict["ykgs"] = sht[i,30].value
        data_dict["cbsfhj"] = sht[i,31].value
        data_dict["djgrghhf"] = sht[i,32].value
        data_dict["sfhj"] = sht[i,33].value
        main_content["email"] = sht[i,34].value
        
        main_content["year"] =year
        main_content["month"] =month
        
        import re
        reg = r"^(\d{1,3}((((,|，)\d{3})*)|\d+))(\.\d+)?$"
        for key in data_dict:
            m = re.search(reg,str(data_dict[key]))
            if  not m:
                data_dict[key] = ' '
        
        html = sendmail.render_template(template,main_content=main_content,data_dict=data_dict)
        to_list = main_content["email"]
        cc = ''
        bcc = ''
        subject = '%s %s年%s月工资条' % (main_content["name"],main_content["year"],main_content["month"])
        
        try:
            if sendmail.send_email(to_list, sender, cc, bcc, subject, html)== True:

                success += 1
        except Exception as e:
            

            logs.get_log().error(str(e))

            continue
            
    logs.get_log().info("任务完成.总计{0}条，成功{1}条，成功率{2:.2%}".format(row-2,success,success/(row-2)))
    logs.shutdown()
    wb.close()
    sys.exit()
            
