import logging
import time
import os



class Logger:
    def __init__(self, logger_name):
        # 创建一个logger
        self.logger = logging.getLogger(logger_name)

        self.logger.setLevel(logging.INFO) # 需要记录的日志级别
        basedir = os.path.abspath(os.path.dirname(__file__))
        log_path = os.path.join(basedir, 'logs', time.strftime("%F"))  # 日志根目录 ../logs/yyyy-mm-dd/

        if not os.path.exists(log_path):
            os.mkdir(log_path)

        # # 创建一个handler，用于写入日志文件
        log_name = os.path.join(log_path, 'out.log')
        fh = logging.FileHandler(log_name, encoding='utf-8', mode='a')  # 指定utf-8格式编码，避免输出的日志文本乱码
        fh.setLevel(logging.INFO) # 需要写入的日志级别

        # 创建一个handler，用于将日志输出到控制台
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO) # 需要到控制台的日志级别

        # 定义handler的输出格式
        formatter = logging.Formatter('%(asctime)s-%(name)s-%(levelname)s-%(message)s')
        fh.setFormatter(formatter)
        ch.setFormatter(formatter)

        # 给logger添加handler
        self.logger.addHandler(fh)
        self.logger.addHandler(ch)

    def get_log(self):  # 定义一个函数，回调logger实例"""
        return self.logger

    def shutdown(self):
        for handler in self.logger.handlers[:]:
            self.logger.removeHandler(handler)