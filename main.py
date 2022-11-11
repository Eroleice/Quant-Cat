import argparse
import os.path
import time
import app
import logging

# 设置参数名
parser = argparse.ArgumentParser()
parser.add_argument("-dev", dest="dev", type=str)
parser.add_argument("-logger", dest="logger", type=str)
parser.add_argument("-test", dest="test", type=str)

# 获取参数
args = parser.parse_args()

# 设定document保存路径，默认按日期新建文件夹
if args.dev == 'True':
    path = './output/dev/'
else:
    path = './output/' + time.strftime('%Y-%m-%d', time.localtime()) + '/'

# 如果path目录不存在就创建
if not os.path.exists(path):
    os.mkdir(path)

# 获取logger对象
logger = app.logging.logger

# 设定logger等级
LEVEL = {
    'debug': logging.DEBUG,
    'info': logging.INFO,
    'warning': logging.WARNING,
    'error': logging.ERROR,
    'critical': logging.CRITICAL
}
if args.logger in LEVEL:
    logger.setLevel(LEVEL[args.logger])
else:
    logger.setLevel(logging.INFO)

# 获取新的Document对象
document = app.document.new()

# 向document内添加内容
app.document.addModule(document, path)

# 应用页面设置
app.style.setLayout(document)

# 将document保存到本地
document.save(path + '[%s] QC Daily.docx' % time.strftime('%Y%m%d', time.localtime()))

# release版本后通过邮件发送给订阅者
if not args.dev == "True":
    pass
