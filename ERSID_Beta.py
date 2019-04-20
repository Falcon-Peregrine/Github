from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import datetime as dt


import openpyxl
from openpyxl.styles import Font


class App:
    
	def __init__(self, master):

		master.resizable(width=False, height=False)

		self.menubar = Menu(master)
		self.menu_help = Menu(self.menubar, tearoff=False)
		self.menubar.add_cascade(label='Help', menu=self.menu_help)
		self.menu_help.add_command(label='报销限额', command=self.popupmsg)
		self.menu_help.add_command(label='注意事项', command=self.caution)
		self.menu_help.add_command(label='联系作者', command=self.coder)
		self.menu_help.add_command(label='版本信息', command=self.version)
		master.config(menu=self.menubar)
		
		






		self.frame_header = ttk.Frame(master)
		self.frame_header.place(x=250, y=5)

		self.label_header = ttk.Label(self.frame_header, text='出入境报销系统', foreground='orange')
		self.label_header.grid(column=0, columnspan=4, sticky='E')
		self.year = dt.datetime.today().year
		self.month = dt.datetime.today().month
		self.day = dt.datetime.today().day
		self.weekday = dt.datetime.today().strftime('%A')
		self.today = f'今天是 {self.year}年{self.month}月{self.day}日，{self.weekday}!'
		ttk.Label(self.frame_header, text=self.today).grid()




		
		self.frame_input = ttk.Frame(master)
		self.frame_input.place(x=300, y=160, anchor='center')
		
		ttk.Label(self.frame_input, text='金额（非出差）：').grid(row=0, column=0, columnspan=2, sticky='w')
		self.amount = StringVar()
		self.entry_amount = ttk.Entry(self.frame_input, width=7, textvariable=self.amount)
		self.entry_amount.grid(row=1, column=0, sticky='w')
		
		
		
		
		ttk.Label(self.frame_input, text='案名：').grid(row=0, column=9, padx=20, sticky='w')

		self.case = StringVar()
		self.entry_case = ttk.Entry(self.frame_input, textvariable=self.case, width=18)
		self.entry_case.grid(row=1, column=9, padx=0, sticky='e')

		ttk.Label(self.frame_input, text='车牌：').grid(row=5, column=0, sticky='w')
		self.carPlate = StringVar()
		self.comboCar = ttk.Combobox(self.frame_input, textvariable=self.carPlate, width=10)
		self.comboCar.grid(row=6, column=0, sticky='w')
		self.comboCar.config(values=('粤E2735警','粤E2749警','粤E2310警','粤E2095警','粤E2080警','粤YA086U','粤Y21P75'))
		self.carPlate.set('粤E2080警')

		ttk.Label(self.frame_input, text='地点：').grid(row=10, column=0, sticky='nw')
		self.spot = StringVar()
		self.entry_spot = ttk.Entry(self.frame_input, textvariable=self.spot, width=10)
		self.entry_spot.grid(row=11, column=0, sticky='w')

		ttk.Label(self.frame_input, text='起始：').grid(row=5, column=8, sticky='n')
		ttk.Label(self.frame_input, text='返回：').grid(row=5, column=9, sticky='n')
		self.year =StringVar()
		self.month =StringVar()
		self.day =StringVar()



		self.year_from = IntVar()
		self.month_from = IntVar()
		self.day_from = IntVar()
		self.year_to = IntVar()
		self.month_to = IntVar()
		self.day_to = IntVar()
		Spinbox(self.frame_input, from_=2018, to=2040, textvariable=self.year_from, width=7).grid(row=6, column=8, padx=5)
		self.year_from.set(2019)
		Spinbox(self.frame_input, from_=1, to=12, textvariable=self.month_from, width=7).grid(row=7, column=8, padx=5)
		self.month_from.set(1)
		Spinbox(self.frame_input, from_=1, to=31, textvariable=self.day_from, width=7).grid(row=8, column=8, padx=5)
		self.day_from.set(1)

		Spinbox(self.frame_input, from_=2018, to=2040, textvariable=self.year_to, width=7).grid(row=6, column=9)
		self.year_to.set(2019)
		Spinbox(self.frame_input, from_=1, to=12, textvariable=self.month_to, width=7).grid(row=7, column=9)
		self.month_to.set(1)
		Spinbox(self.frame_input, from_=1, to=31, textvariable=self.day_to, width=7).grid(row=8, column=9)
		self.day_to.set(1)



		self.frame_staff = ttk.Frame(master)
		self.frame_staff.place(x=280, y=280)
		ttk.Label(self.frame_staff, text='人员（双击选择）：').pack(anchor='w')
			

		self.scrollbar = Scrollbar(self.frame_staff)
		self.staff=StringVar()
		self.listbox=Listbox(self.frame_staff, listvariable=self.staff, selectmode=MULTIPLE)
		self.staff.set(('张鉴良','蔡继洪','白啟华','陈耀鹏','何敬谦','何志洪','陈仲文','严辉','张伟峰','范宇靖',
				'黄泽鸿','莫坚锐', '谭哲','许勇斌','邓建锋','邓超劲','叶土生','黄婉君',
				'------', '朱啟基','陈剑龙','关务安','周成光','李乐敏','赵文','吴俊锋','欧驰俊','刘煜熹', '黄沛云','王若希','刘丽君',
				'梁叶莹','江林鑫','曹枫越','赵德蛟','钟欢然','邓志鹏','------','蔡建青','李伟根','唐悦茹','邹明利','高桂泉','何雨佳',
				'张惠梅','区重贤','龙钦','苏越峰','吴永生','钟惜葵','柳力方','张永平','黄敏珊','陈羡仪','蔡卉婷','张瑞婵',
				'黄颖斯','郭雷','杨小云'))
		# self.listbox.unbind('<Button-1>')
		self.listbox.bind('<Double-Button-1>', self.getstaff)
		self.scrollbar.config(command=self.listbox.yview)
		self.listbox.config(yscrollcommand=self.scrollbar.set)
		self.scrollbar.pack(side=RIGHT, fill=Y)
		self.listbox.pack()

		# self.checkbut_case = ttk.Checkbutton(self.frame_staff, text='自动设置案名', command=None)
		# self.checkbut_case.place(x=100, y=400)




		self.frame_function = ttk.Frame(master)
		self.frame_function.place(x=233, y=500)
		
		self.run = ttk.Button(self.frame_function, text='启动', command=self.askType)
		self.run.grid(row=0, column=0, ipadx=20, ipady=10, padx=20, pady=20)

		self.clear = ttk.Button(self.frame_function, text='清除', command=self.inputclear)
		self.clear.grid(row=0, column=1, ipadx=20, padx=20, pady=20, ipady=10)
		self.staff_select=[]




		self.frame_trip = ttk.Frame(master, borderwidth=10, relief=RIDGE)
		self.frame_trip.place(x=600, y=65)

		ttk.Label(self.frame_trip, text='市内交通费：').grid(row=0, column=0)
		self.commute = StringVar()
		self.entry_commute = ttk.Entry(self.frame_trip, width=7, textvariable=self.commute)
		self.entry_commute.grid(row=1, column=0)
		self.commute.set(0)

		ttk.Label(self.frame_trip, text='伙食补助费：').grid(row=3, column=0)
		self.late_meal = StringVar()
		self.entry_late_meal = ttk.Entry(self.frame_trip, width=7, textvariable=self.late_meal)
		self.entry_late_meal.grid(row=4, column=0)
		self.late_meal.set(0)

		ttk.Label(self.frame_trip, text='住宿费：').grid(row=6, column=0)
		self.accomodation = StringVar()
		self.entry_accomodation = ttk.Entry(self.frame_trip, width=7, textvariable=self.accomodation)
		self.entry_accomodation.grid(row=7, column=0)
		self.accomodation.set(0)

		ttk.Label(self.frame_trip, text='交通费：').grid(row=9, column=0)
		self.transport = StringVar()
		self.entry_transport = ttk.Entry(self.frame_trip, width=7, textvariable=self.transport)
		self.entry_transport.grid(row=10, column=0)
		self.transport.set(0)


		ttk.Label(self.frame_trip, text='协同办案费：').grid(row=12, column=0)
		self.cowork = StringVar()
		self.entry_cowork = ttk.Entry(self.frame_trip, width=7, textvariable=self.cowork)
		self.entry_cowork.grid(row=13, column=0)
		self.cowork.set(0)

		ttk.Label(self.frame_trip, text='其他：').grid(row=16, column=0)
		self.others = StringVar()
		self.entry_others = ttk.Entry(self.frame_trip, width=7, textvariable=self.others)
		self.entry_others.grid(row=17, column=0)
		self.others.set(0)
		ttk.Label(self.frame_trip, text='此栏仅出差填写', foreground='orange').grid(row=20, column=0)




		

		self.style = ttk.Style()
		self.style.theme_use('vista')
		self.style.configure('Header.TLabel', font=('宋体',20, 'bold'))
		self.style.configure('Alarm.TButton', foreground='orange', font=('Arial', 12, 'bold'))

		self.label_header.config(style='Header.TLabel')
		self.style.map('Alarm.TButton', foreground=[('pressed', '#183A54'), ('disabled','#00A3E0')], background=[('pressed', '#00A3E0'), ('disabled','#00A3E0')])
		self.run.config(style='Alarm.TButton')
		self.clear.config(style='Alarm.TButton')
		







	def getstaff(self, event):
	    if self.listbox.get(ACTIVE) not in self.staff_select:
	    
	        self.staff_select.append(self.listbox.get(ACTIVE))
	    
	    # print(self.staff_select)

	def askType(self):

		ques = messagebox.askyesno(title='报销类型', message='是：办案； 否：出差')
		if ques is True:
			# print(self.amount.get(), self.case.get(),
			#  self.carPlate.get(), self.checkbut_case.instate(['selected']),
			#   self.spot.get(), self.year_from.get(), self.day_to.get(), self.staff_select)
			
			self.operation()
			self.savefile()

			
		
		elif ques is False:
			

			
			self.entry_amount['state'] = 'disabled'
			self.operation()
			sheet_mission['G7'].value = self.late_meal.get() # 伙食补助费
			

			try:
				sheet_mission['L7'].value = int(self.commute.get()) + int(self.late_meal.get()) + int(self.accomodation.get()) + int(self.transport.get()) + int(self.cowork.get()) + int(self.others.get()) # 合计
				sheet_mission['L11'].value = sheet_mission['L7'].value # 最后合计
			except Exception as e:
				print(str(e))
			
			self.savefile()

			


			


	def operation(self):
		PathOpen = 'ExcelSheet.xlsx'
		global wb
		wb = openpyxl.load_workbook(PathOpen)
		sheet_case = wb['18年三非详表']

		sheet_expenditure = wb['经费支出表'] 
		sheet_expenditure['B3'].value = '办理' + str(self.case.get()) + '差旅费'# 项目名称
		sheet_expenditure['F3'].value = self.amount.get() # 预算金额
		sheet_expenditure['B4'].value = '办理' + str(self.case.get()) + '差旅费'# 支出原因

		sheet_a4= wb['出差审批表']
		#sheet_a4['E3'].value = self.staff_select # 出差人姓名
		sheet_a4['B4'].value = f'{self.year_from.get()}.{self.month_from.get()}.{self.day_from.get()}—{self.year_to.get()}.{self.month_to.get()}.{self.day_to.get()}' # 出差时间
		sheet_a4['E4'].value = self.spot.get() # 目的地
		sheet_a4['E3'].value = ' '.join(self.staff_select) # 出差人姓名
		sheet_a4['B6'].value = '赴'+self.spot.get()+ '办理' + str(self.case.get()) # 出差（出行）事由
		sheet_a4['C7'].value = '是' # 是否单位派车
		sheet_a4['E7'].value = self.carPlate.get() # 车牌号码
		
		global sheet_mission
		sheet_mission = wb['差旅费（含办案）报销表']
		self.dict_staff_rank = {'张鉴良':'(正股职)', '蔡继洪':'（正股职）', '白啟华':'（正科级）','陈耀鹏':'（正股职）','何敬谦':'（正股级）', '何志洪':'(正股级）','陈仲文':'（正股级）','严辉':'（正股级）',\
                                        '张伟峰':'(科员)','范宇靖':'（科员）','黄泽鸿':'（辅警）','莫坚锐':'（辅警）','谭哲':'（辅警）','许勇斌':'（辅员）','邓建锋':'（辅警）','邓超劲':'（辅警）','叶土生':'（辅警）','黄婉君':'（辅警）',\
                                        '朱啟基':'（副科职）','陈剑龙':'（副科职）','关务安':'（正股职）','周成光':'（正股级）','李乐敏':'（副股级）','赵文':'（科员）','吴俊锋':'（副股级）','欧驰俊':'（科员）','刘煜熹':'（科员）',\
                                        '黄沛云':'（正股级）','王若希':'（科员）','刘丽君':'（科员）','梁叶莹':'（辅警）','江林鑫':'（辅警）','曹枫越':'（辅警）','赵德蛟':'（辅警）','钟欢然':'（辅警）','邓志鹏':'（辅警）',\
                                        '蔡建青':'（正科职）','李伟根':'（正科职）','唐悦茹':'（正科职）','邹明利':'（正科职）','高桂泉':'（副股级）','何雨佳':'（副股级）','张惠梅':'（科员）','区重贤':'（副股级）',\
                                        '龙钦':'（副股级）','苏越峰':'（正股级）','吴永生':'（正股级）','钟惜葵':'（正股级）','柳力方':'（副股级）','张永平':'（副股级）','黄敏珊':'（副股级）',\
                                        '陈羡仪':'（副股职）','蔡卉婷':'（副股级）','张瑞婵':'（副股级）','黄颖斯':'（副股级）','郭雷':'（副股级）','杨小云':'（副股级）'}
		
		self.list_staff_rank = []

		for name in self.staff_select:
			self.list_staff_rank.append(name+self.dict_staff_rank[name])
		self.str_staff_rank = ' '.join(self.list_staff_rank)

		sheet_mission['A3'].value = '姓名（级别）:' + self.str_staff_rank # 姓名（级别）
		sheet_mission['A7'].value = self.spot.get() # 出差地点
		sheet_mission['B7'].value = f'{self.year_from.get()}.{self.month_from.get()}.{self.day_from.get()}' # 起始时间
		sheet_mission['C7'].value = f'{self.year_to.get()}.{self.month_to.get()}.{self.day_to.get()}' # 返回时间
		
		start_day = dt.datetime(self.year_from.get(), self.month_from.get(), self.day_from.get())
		return_day = dt.datetime(self.year_to.get(), self.month_to.get(), self.day_to.get())
		day_delta = (return_day - start_day).days + 1

		sheet_mission['D7'].value = day_delta# 天数
		sheet_mission['E7'].value = len(self.staff_select)# 人数
		sheet_mission['F7'].value = self.commute.get() # 市内交通费 80/人/日
		sheet_mission['G7'].value = self.amount.get() # 伙食补助费 100/人/日
		sheet_mission['H7'].value = self.accomodation.get() # 住宿费
		sheet_mission['I7'].value = self.transport.get() # 交通费
		sheet_mission['J7'].value = self.cowork.get() # 协同办案费
		sheet_mission['K7'].value = self.others.get() # 其他
		sheet_mission['L7'].value = self.amount.get() # 合计
		sheet_mission['L11'].value = self.amount.get() # 最后合计
		
		sheet_cash = wb['非公务卡报销'] 
		sheet_latemeal = wb['误餐审批表']



	def inputclear(self):
		
		self.case.set('')
		self.spot.set('')
		self.staff_select=[]
		self.commute.set(0)
		self.late_meal.set(0)
		self.accomodation.set(0)
		self.transport.set(0)
		self.cowork.set(0)
		self.others.set(0)
		self.amount.set('')
		self.entry_amount['state'] = 'active'
		self.year_from.set(2019)
		self.month_from.set(1)
		self.day_from.set(1)
		self.year_to.set(2019)
		self.month_to.set(1)
		self.day_to.set(1)
		self.listbox.selection_clear(0, END)


		



	def savefile(self):	
		wb.save('报销备份/报销备份_'+dt.datetime.today().strftime('%Y%m%d')+'.xlsx')




	def popupmsg(self):
	    messagebox.showinfo(title='报销限额', message='市内交通费:80元/人/天\n伙食补助费:100元/人/天')
	def caution(self):
	    messagebox.showinfo(title='注意事项', message="请保证同级文件夹内的ExcelSheet.xlsx不被修改。\n左侧面板为基本面板，平时办案报销用。\n右侧面板为出差面板，出差报销如无该款项请填上0，需返回Excel修改。\n按下启动按钮后并选择报销类型后，自动在“报销备份”文件夹生成新表。\n如Excel表格已打开，则无法正常运作。")
	def coder(self):
		messagebox.showinfo(title='联系作者', message='1395908181')
	def version(self):
		messagebox.showinfo(title='版本信息', message='Copyright© 2019 Weifeng Zhang\nVersion: Alpha')



def main():
	root = Tk()

	root.geometry("800x600")
	root.wm_title("Expenses Reimbersment System for Immigration Department")
	root.iconbitmap(default='coin.ico')


	
	app = App(root)

	root.mainloop()
	
	
	

if __name__ == '__main__': main()
