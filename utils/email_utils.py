import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

"""
    邮件发送工具
"""
def send_email(file, receiver):
    from_addr = '1394783493@qq.com'  # 邮件发送账号
    qqCode = 'astwvdyluucehgaa'  # 授权码（这个要填自己获取到的）
    smtp_server = 'smtp.qq.com'  # 固定写死
    smtp_port = 465  # 固定端口

    # 配置服务器
    stmp = smtplib.SMTP_SSL(smtp_server, smtp_port)
    stmp.login(from_addr, qqCode)

    # 组装发送内容
    message = MIMEMultipart()

    message['From'] = Header("自动数据采集程序", 'utf-8')
    message['To'] = Header("招标项目组", 'utf-8')
    subject = '项目招标数据汇总'
    message['Subject'] = Header(subject, 'utf-8')

    # 邮件正文内容
    message.attach(MIMEText('这是今日的项目招标信息，请注意查收', 'plain', 'utf-8'))

    # 添加附件
    att1 = MIMEText(open(file, 'rb').read(), 'base64', 'utf-8')
    att1["Content-Type"] = 'application/octet-stream'
    # 解析文件路径，获取文件名
    splits = file.split("\\")
    # 这里的filename可以任意写，写什么名字，邮件中显示什么名字
    # 下面这种方式，会造成附件名有中文时乱码
    # att1["Content-Disposition"] = 'attachment; filename="' + "2021-04-21_项目招标.xlsx"
    att1.add_header('Content-Disposition', 'attachment', filename=('utf-8', '', splits[len(splits) - 1]))
    message.attach(att1)

    try:
        stmp.sendmail(from_addr, receiver, message.as_string())
        print('邮件发送成功')
    except Exception as e:
        print('邮件发送失败：' + str(e))


if __name__ == '__main__':
    file = r"E:\pythonProject\Crawler\spiders\2021-04-21_招标项目.xlsx"
    receiver = ['1309961163@qq.com','1394783493@qq.com']
    send_email(file, receiver)
