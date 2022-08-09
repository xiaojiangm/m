import pandas as pd
import sys,os


class Get_Excle():

	def __init__(self, xlsxs):

		self.xlsxs = xlsxs

	def get_excel(self):

		try:
			df = pd.read_excel(self.xlsxs)
			df.fillna(" ", inplace = True) #将默认NaN值转换为空
			if "邮箱" in df.columns:
				return df
			else:
				print("请在表中添加列名为邮箱那的一列,将邮箱信息正确添加进去否则无法发送邮件")
				os.system("pause")
				sys.exit()

		except FileNotFoundError as e:
			print("未找到该文件")
			os.system("pause")
			sys.exit()

		except ValueError as e:
			fm = self.xlsxs.split(".")[-1]
			print("不能打开格式为{}的文件".format(fm))
			os.system("pause")
			sys.exit()

		except Exception as e:
			print(e)
			print("文件读取错误,请检查文件是否加密,数据格式转换是否异常")
			os.system("pause")
			sys.exit()