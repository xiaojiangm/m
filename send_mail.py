import smtplib
from get_file import Get_Excle
from email.mime.text import MIMEText
import os
import sys
import time
from datetime import datetime

class Mail:

	def __init__(self, mail_host, mail_user, mail_pass, mail_title):

		self.mail_host = mail_host
		self.mail_user = mail_user
		self.mail_pass = mail_pass
		self.mail_title = mail_title

		self.smtpObj = smtplib.SMTP()

	def login_mail(self):
		try:
			self.smtpObj.connect(self.mail_host, 25)
			self.smtpObj.login(self.mail_user, self.mail_pass)
		except smtplib.SMTPException as e:
			print('账号密码错误请重新输入或服务器主动断开连接,请退出并于5分钟后重试')
			main()		#失败后重新调用主函数

	def get_data(self,datas):	#登录成功后获取数据
		count = 0
		table = datas.columns	#获取列名
		for row in range(len(datas)):	#获取循环的行数
			row_text = "<tr>"   #每次循环都会重新获取下一行的字段
			cell_text = "<tr>"
			data = datas.loc[row].values	#获取每行的数据
			count += 1
			if datas.loc[row]["邮箱"] != " ":
				for i in range(len(table)):		#循环每行的数据，并邮件发送每行
					row_text += f'<td>{table[i]}</td>'	#获取列名
					cell_text += f'<td>{data[i]}</td>'	#获取对应列名的对应值
				row_text += "</tr>"
				cell_text += "</tr>"
				mail = datas.loc[row]["邮箱"]
				names = datas.loc[row]['姓名']
				self.send_mail(row_text, cell_text, mail,names)
			else:
				print("没有找到第{}行的邮箱信息,请添加邮箱信息".format(row))
				filename = datetime.now().strftime("%Y%m%d")
				os.system(f"echo 没有找到第{row}行的邮箱信息，邮件未发送成功 >> {filename}_fail.log ")
				continue
		if count >= len(datas):
			print("邮件发送完毕")
			self.smtpObj.quit()
			os.system("pause")
			sys.exit()

	def send_mail(self,row_text, cell_text, mail,names):		#获取数据后发送数据
		content = f'''
			<table border="1px solid black" cellspacing="0px" style="border-collapse:collapse;collapse" align="center" width="5400">
			{row_text}
			{cell_text}
			</table>
			'''
		message = MIMEText(content,'html','utf-8')
		message['Subject'] = self.mail_title
		message['From'] = self.mail_user
		message['To'] = mail

		try:

			self.smtpObj.sendmail(
				self.mail_user, mail, message.as_string())
			filename = datetime.now().strftime("%Y%m%d")
			os.system(f"echo 邮件已成功发送给{names} >> {filename}_success.log ")
			print(f'邮件已成功发送给{names}')
		except smtplib.SMTPException as e:				#每发送20封，服务器就会主动断掉连接，需要重新连接
			self.login_mail()
			self.send_mail(row_text, cell_text, mail,names)

def main():

	mail_host = 'mail.cbcbeijing.net'
	mail_user = input("请输入邮箱账号:") + "@cbcbeijing.net"
	mail_pass = input("请输入密码:")
	mail_title = input("请输入邮件主题:")
	xlsxs = input(r"请输入文件路径及名称:")

	sends_mail = Mail(mail_host, mail_user, mail_pass, mail_title)
	datas = Get_Excle(xlsxs)

	sends_mail.login_mail()
	sends_mail.get_data(datas.get_excel())

if __name__ == "__main__":

	now = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
	s = '2023-03-16 00:00:00'
	if now > s:
		sys.exit()
	else:
		main()