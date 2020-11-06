import smtplib
from email.mime.text import MIMEText
from email.header import Header
from email.message import EmailMessage
import imghdr
from email.utils import make_msgid

import os
# For guessing MIME type based on file name extension
import mimetypes

from argparse import ArgumentParser
from email.policy import SMTP

class smtpLogin:
    sender = 'guorenliang2009@live.cn'
    receivers = ['leo_of_love@qq.com']  # 接收邮件，可设置为你的QQ邮箱或者其他邮箱
    password = "I!verson2009"
    output = './CtrlBox/2.txt'
    directory = "./CtrlBox"
    loadImage = "./CtrlBox/p.jpg"
    
    def login(self):
        try:
            smtpObj = smtplib.SMTP('smtp.outlook.com')
            smtpObj.connect('smtp.outlook.com')
            smtpObj.starttls()
            smtpObj.login(self.sender,self.password)
            return smtpObj
        except smtplib.SMTPException:
            print("Error: 邮箱登录发生错误")

class SendMail(smtpLogin):

    def Text(self):
        # 三个参数：第一个为文本内容，第二个 plain 设置文本格式，第三个 utf-8 设置编码
        message = MIMEText('Python 邮件发送测试...', 'plain', 'utf-8')
        message['from'] = 'guorenliang2009@live.cn'
        message['to'] =  'leo_of_love@qq.com'
        # 第一个参数为邮件标题
        message['Subject'] = Header('Python SMTP 邮件测试', 'utf-8')
        return message

    def sendText(self):
        try:
            smtpObj = self.login()
            smtpObj.sendmail(self.sender, self.receivers, self.Text().as_string())
            smtpObj.quit()
            print("邮件发送成功")
        except smtplib.SMTPException:
            print("Error: 无法发送邮件")

    def sendMsg(self,msg):
        try:
            smtpObj = self.login()
            smtpObj.send_message(msg)
            smtpObj.quit()
            print("邮件发送成功")
        except smtplib.SMTPException:
            print("Error: 无法发送邮件")

    ### 发送文件夹里的所有文件和一个插入image的html为内容的信件
    def sendDirectory(self):
        # Create the message
        msg = EmailMessage()
        msg['Subject'] = f'Contents of directory {os.path.abspath(self.directory)}'
        # msg['To'] = ', '.join(args.recipients)
        # msg['From'] = args.sender
        msg['To'] = ['leo_of_love@qq.com']
        msg['From'] = self.sender
        msg.set_content("""\
Salut!

Cela ressemble à un excellent recipie[1] déjeuner.

[1] http://www.yummly.com/recipe/Roasted-Asparagus-Epicurious-203718

--Pepé
""")
        msg.preamble = 'You will not see this in a MIME-aware mail reader.\n'

        img_cid = make_msgid()
        msg.add_alternative("""\
<html>
  <head></head>
  <body>
    <p>Salut!</p>
    <p>Cela ressemble à un excellent
        <a href="http://www.yummly.com/recipe/Roasted-Asparagus-Epicurious-203718">
            recipie
        </a> déjeuner.
    </p>
    <img src="cid:{img_cid}" />
  </body>
</html>
""".format(img_cid=img_cid[1:-1]), subtype='html')

        with open(self.loadImage, 'rb') as img:
            msg.get_payload()[1].add_related(img.read(), 'image', 'jpeg',
                                     cid=img_cid)
        
        for filename in os.listdir(self.directory):
            path = os.path.join(self.directory, filename)
            if not os.path.isfile(path):
                continue
        # Guess the content type based on the file's extension.  Encoding
        # will be ignored, although we should check for simple things like
        # gzip'd or compressed files.
            ctype, encoding = mimetypes.guess_type(path)
            if ctype is None or encoding is not None:
                # No guess could be made, or the file is encoded (compressed), so
                # use a generic bag-of-bits type.
                ctype = 'application/octet-stream'
            maintype, subtype = ctype.split('/', 1)
            with open(path, 'rb') as fp:
                msg.add_attachment(fp.read(),
                                   maintype=maintype,
                                   subtype=subtype,
                                   filename=filename)
                # Now send or store the message
        if self.output:
            with open(self.output, 'wb') as fp:
                fp.write(msg.as_bytes(policy=SMTP))
            self.sendMsg(msg)

    def commandLine():
        parser = ArgumentParser(description="""\
Send the contents of a directory as a MIME message.
Unless the -o option is given, the email is sent by forwarding to your local
SMTP server, which then does the normal delivery process.  Your local machine
must be running an SMTP server.
""")
        parser.add_argument('-d', '--directory',
                        help="""Mail the contents of the specified directory,
                        otherwise use the current directory.  Only the regular
                        files in the directory are sent, and we don't recurse to
                        subdirectories.""")
        parser.add_argument('-o', '--output',
                        metavar='FILE',
                        help="""Print the composed message to FILE instead of
                        sending the message to the SMTP server.""")
        parser.add_argument('-s', '--sender', required=True,
                        help='The value of the From: header (required)')
        parser.add_argument('-r', '--recipient', required=True,
                        action='append', metavar='RECIPIENT',
                        default=[], dest='recipients',
                        help='A To: header value (at least one required)')
        args = parser.parse_args()
        directory = args.directory
                
        if not directory:
            directory = '.'

#n = SendMail().sendText()
n = SendMail().sendDirectory()

