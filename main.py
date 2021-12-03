1 #-*- encoding:utf-8 -*-

import sys
import os


def welcome():
    print("==================================================="+
        "\n===========欢迎使用工资表批量邮件发送程序===========\n"+
        "===================================================\n"+
        "请选择发送的工资表类型\n"+
        "[1] 月度工资表\n"+
        "[e] 退出\n"
        )
    name = input("请选择:")
    if name == "1":
        import configparser
        conf= configparser.ConfigParser()
        root_path = os.path.dirname( os.path.abspath(__file__) )
        # print("root:",os.path.dirname( os.path.abspath(__file__) ))
        conf.read(os.path.join(root_path,'email.conf'))
        # print(os.path.join(root_path,'email.conf'))
        excel_doc = conf.get("monthly", "excel") # 数据文件
        template = conf.get("monthly", "template") # 模板文件
        print("请检查是否已准备好数据文件%s,并已正确放置在当前目录下\n"%excel_doc)
        ready = input("准完成后请按[y],其他任意键退出：")
        
        if ready == 'y':
            year_input = input("请输入年份，如2021：")
            month_input = input("请输入月份，如02：")
        else:
            
            sys.exit
        import data
        data.month_pay(year_input,month_input)
            
    elif name == "e":
        print("再见")
        sys.exit
    else:
        print("选择无效\n")
        welcome()
        
if __name__ == '__main__':
    welcome()
