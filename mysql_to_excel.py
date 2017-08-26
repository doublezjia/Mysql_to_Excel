# 环境python3.x
#
# 脚本说明
# 备份资产表格数据库到Excel表中
# 
# 用到的第三方库有 pymysql和xlwt,使用前要先安装好  
# 
# 用之前请修改存放文件的根目录 rootdir ，默认存放在脚本存放的目录下
# 
# excel表会存放在以当天日期为名的文件夹中。 
# 
# 脚本运行比较慢，后期再想方法解决
# 
# 
# 
# ver 1.0 2017-8-25
# All right reserved by zealous
# 
# 




import os,sys,time,datetime,pymysql,xlwt


# 文件夹目录
rootdir = r'.'


# 日志目录
logdir=r'.\log'
excel_logfile = logdir+r'\excel_log.log'
# 脚本的日志目录

error_logfile = logdir+r'\error_log.log'

# mysql地址 端口 用户名 密码 数据库 数据库编码 备份文件名和目录
localhost = 'localhost'
port = 3306
mysqluser = 'root'
mysqlpwd = 'root'
mysqldatabase = 'device_manager'
charset = 'utf8'




# excel表格字段
asset_belong =[]
asset_number = []
asset_type = []
device_type = []
device_brand = []
device_name = []
cpu = []
computer_board = []
displya_card = []
hard_disk = []
memory = []
device_version = []
config = []
imei = []
mac =[]
other = []
use_user = []
depart = []
device_status = []
receive_time = []
return_time= []
buy_time = []
remark = []


# mysql数据库连接，用来判断数据库是否能连接上，避免备份出错
def mysql_con():
	try:
		# 连接数据库
		conn = pymysql.connect(host=localhost,port=port,user=mysqluser,passwd=mysqlpwd,db=mysqldatabase,charset=charset)
		return conn
	except:
		error_log = '[Error] MySQL connect Error.----------------- '+datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')+'\n'
		with open(error_logfile,'a') as f:
			f.write(error_log)
		sys.exit()


# 转换时间
def str_time(strtime):
	strtime = datetime.datetime.fromtimestamp(strtime).strftime('%Y-%m-%d')
	return strtime



# 资产所属
def assetBelong(asset_belong):
	asset_belong = asset_belong
	if asset_belong == 1 :
		asset_belong = '广州'
	elif asset_belong == 2 :
		asset_belong = '上海'
	elif asset_belong == 3 :
		asset_belong = '北京'
	elif asset_belong == 4 :
		asset_belong = '珠海'
	return asset_belong

# 资产类型
def assetType(asset_type):
	asset_type = asset_type
	if asset_type == 1 :
		asset_type = 'IT类'
	elif asset_type == 2 :
		asset_type = '行政类'
	return asset_type

# 设备类型
def deviceType(device_type):
	device_type = device_type
	if device_type == 1 :
		device_type = '台式主机'
	elif device_type == 2:
		device_type = '显示器'
	elif device_type == 3 :
		device_type = '手绘板'
	elif device_type == 4 :
		device_type = '手机'
	elif device_type == 5 :
		device_type = '平板'
	elif device_type == 6 :
		device_type = '笔记本'
	elif device_type == 7 :
		device_type = 'IMAC'
	elif device_type == 8 :
		device_type = '其他'
	return device_type

# 设备状态
def deviceStatus(device_status):
	device_status = device_status
	if device_status == 1 :
		device_status = '正常使用'
	elif device_status == 2 :
		device_status = '维修'
	elif device_status == 3 :
		device_status = '借用'
	elif device_status == 4 :
		device_status = '报废'
	elif device_status == 5 :
		device_status = '损坏'
	elif device_status == 6 :
		device_status = '备用'
	return device_status

# 使用部门
def Depart(depart,cur):
	depart = depart
	if depart == '' or depart == '0':
		depart = ''
		return depart
	else:
		# 查询depart数据表
		comm = 'select * from depart where cate_id="'+str(depart)+'"'
		cur.execute(comm)
		depart = cur.fetchall()
		return depart[0][2]


# 总表
def excel_total(filedir):

	conn = mysql_con()
	cur = conn.cursor()
	comm = 'select * from asset_device '
	cur.execute(comm)
	ret1 = cur.fetchall()
	filedir = filedir
	filename = filedir+r'\总表.xls'

	# 设置标题颜色
	styleBoldRed  =xlwt.easyxf('font: color-index black, bold on')
	# 新建表格
	desxls = xlwt.Workbook(encoding='utf-8')
	write_sheet = desxls.add_sheet("sheet1",cell_overwrite_ok=True)

	headlist = ["资产所属","资产编号","资产类型",
	"设备类别","设备品牌","设备名称","设备型号","CPU","主板",
	"显卡","硬盘","内存","配置","IMEI码","MAC地址",
	"其他","使用人","使用部门","设备状态","领用时间",
	"归还时间","购买时间","备注"]

	for i in range(len(headlist)):
		write_sheet.write(0,i,headlist[i],styleBoldRed)

	for i in range(len(ret1)):
		# 资产所属
		asset_belong = ret1[i][1]
		asset_belong = assetBelong(asset_belong)
		# 资产编号
		asset_number = ret1[i][2]
		# 资产类型
		asset_type = ret1[i][3]
		asset_type = assetType(asset_type)
		# 设备类别
		device_type = ret1[i][4]
		device_type = deviceType(device_type) 
		# 设备品牌
		device_brand = ret1[i][5]
		# 设备名称
		device_name = ret1[i][6]
		# 设备型号
		device_version = ret1[i][12]
		# CPU
		cpu = ret1[i][7]
		# 主板
		computer_board = ret1[i][8]
		# 显卡
		displya_card = ret1[i][9]
		# 硬盘
		hard_disk = ret1[i][10]
		# 内存
		memory = ret1[i][11]
		# 配置
		config = ret1[i][13]
		# IMEI码
		imei = ret1[i][14]
		# MAC地址
		mac = ret1[i][15]
		# 其他
		other = ret1[i][16]
		# 使用人
		use_user = ret1[i][17]
		# 使用部门
		depart = ret1[i][18]
		depart = Depart(depart,cur)
		# 设备状态
		device_status = ret1[i][19]
		device_status = deviceStatus(device_status) 
		# 领用时间
		receive_time = ret1[i][20]
		if receive_time == 0:
			receive_time = ''
		else:
			receive_time = str_time(receive_time)
		# 归还时间
		return_time = ret1[i][21]
		if return_time == 0:
			return_time = ''
		else:
			return_time = str_time(return_time)
		# 购买时间
		buy_time = ret1[i][22]
		if buy_time == 0:
			buy_time = ''
		else:
			buy_time = str_time(buy_time)
		# 备注
		remark = ret1[i][24]

		write_sheet.write(i+1,0,asset_belong)
		write_sheet.write(i+1,1,asset_number)
		write_sheet.write(i+1,2,asset_type)
		write_sheet.write(i+1,3,device_type)
		write_sheet.write(i+1,4,device_brand)
		write_sheet.write(i+1,5,device_name)
		write_sheet.write(i+1,6,device_version)
		write_sheet.write(i+1,7,cpu)
		write_sheet.write(i+1,8,computer_board)
		write_sheet.write(i+1,9,displya_card)
		write_sheet.write(i+1,10,hard_disk)
		write_sheet.write(i+1,11,memory)
		write_sheet.write(i+1,12,config)
		write_sheet.write(i+1,13,imei)
		write_sheet.write(i+1,14,mac)
		write_sheet.write(i+1,15,other)
		write_sheet.write(i+1,16,use_user)
		write_sheet.write(i+1,17,depart)
		write_sheet.write(i+1,18,device_status)
		write_sheet.write(i+1,19,receive_time)
		write_sheet.write(i+1,20,return_time)
		write_sheet.write(i+1,21,buy_time)
		write_sheet.write(i+1,22,remark)
	# 保存表格
	desxls.save(filename)
	cur.close()
	conn.close()
	return 0

# 主机
def excel_computer(filedir):

	conn = mysql_con()
	cur = conn.cursor()
	comm = 'select * from asset_device where device_type="1" '
	cur.execute(comm)
	comp_ret = cur.fetchall()

	filedir = filedir
	filename = filedir+r'\主机.xls'

	# 设置标题颜色
	styleBoldRed  =xlwt.easyxf('font: color-index black, bold on')
	# 新建表格
	desxls = xlwt.Workbook(encoding='utf-8')
	write_sheet = desxls.add_sheet("sheet1",cell_overwrite_ok=True)

	headlist = ["资产所属","资产编号","资产类型",
	"设备类别","设备品牌","CPU","主板",
	"显卡","硬盘","内存",
	"其他","使用人","使用部门","设备状态","领用时间",
	"归还时间","购买时间","备注"]

	for i in range(len(headlist)):
		write_sheet.write(0,i,headlist[i],styleBoldRed)

	for i in range(len(comp_ret)):
		# 资产所属
		asset_belong = comp_ret[i][1]
		asset_belong = assetBelong(asset_belong)
		# 资产编号
		asset_number = comp_ret[i][2]
		# 资产类型
		asset_type = comp_ret[i][3]
		asset_type = assetType(asset_type)
		# 设备类别
		device_type = comp_ret[i][4]
		device_type = deviceType(device_type) 
		# 设备品牌
		device_brand = comp_ret[i][5]
		# CPU
		cpu = comp_ret[i][7]
		# 主板
		computer_board = comp_ret[i][8]
		# 显卡
		displya_card = comp_ret[i][9]
		# 硬盘
		hard_disk = comp_ret[i][10]
		# 内存
		memory = comp_ret[i][11]
		# 其他
		other = comp_ret[i][16]
		# 使用人
		use_user = comp_ret[i][17]
		# 使用部门
		depart = comp_ret[i][18]
		depart = Depart(depart,cur)
		# 设备状态
		device_status = comp_ret[i][19]
		device_status = deviceStatus(device_status) 
		# 领用时间
		receive_time = comp_ret[i][20]
		if receive_time == 0:
			receive_time = ''
		else:
			receive_time = str_time(receive_time)
		# 归还时间
		return_time = comp_ret[i][21]
		if return_time == 0:
			return_time = ''
		else:
			return_time = str_time(return_time)
		# 购买时间
		buy_time = comp_ret[i][22]
		if buy_time == 0:
			buy_time = ''
		else:
			buy_time = str_time(buy_time)
		# 备注
		remark = comp_ret[i][24]

		write_sheet.write(i+1,0,asset_belong)
		write_sheet.write(i+1,1,asset_number)
		write_sheet.write(i+1,2,asset_type)
		write_sheet.write(i+1,3,device_type)
		write_sheet.write(i+1,4,device_brand)
		write_sheet.write(i+1,5,cpu)
		write_sheet.write(i+1,6,computer_board)
		write_sheet.write(i+1,7,displya_card)
		write_sheet.write(i+1,8,hard_disk)
		write_sheet.write(i+1,9,memory)
		write_sheet.write(i+1,10,other)
		write_sheet.write(i+1,11,use_user)
		write_sheet.write(i+1,12,depart)
		write_sheet.write(i+1,13,device_status)
		write_sheet.write(i+1,14,receive_time)
		write_sheet.write(i+1,15,return_time)
		write_sheet.write(i+1,16,buy_time)
		write_sheet.write(i+1,17,remark)
	# 保存表格
	desxls.save(filename)
	cur.close()
	conn.close()
	return 0

# 显示器
def excel_display(filedir):

	conn = mysql_con()
	cur = conn.cursor()
	comm = 'select * from asset_device where device_type="2"  '
	cur.execute(comm)
	ret1 = cur.fetchall()
	filedir = filedir
	filename = filedir+r'\显示器.xls'

	# 设置标题颜色
	styleBoldRed  =xlwt.easyxf('font: color-index black, bold on')
	# 新建表格
	desxls = xlwt.Workbook(encoding='utf-8')
	write_sheet = desxls.add_sheet("sheet1",cell_overwrite_ok=True)

	headlist = ["资产所属","资产编号","资产类型",
	"设备类别","设备品牌","设备型号",
	"配置","使用人","使用部门","设备状态","领用时间",
	"归还时间","购买时间","备注"]

	for i in range(len(headlist)):
		write_sheet.write(0,i,headlist[i],styleBoldRed)

	for i in range(len(ret1)):
		# 资产所属
		asset_belong = ret1[i][1]
		asset_belong = assetBelong(asset_belong)
		# 资产编号
		asset_number = ret1[i][2]
		# 资产类型
		asset_type = ret1[i][3]
		asset_type = assetType(asset_type)
		# 设备类别
		device_type = ret1[i][4]
		device_type = deviceType(device_type) 
		# 设备品牌
		device_brand = ret1[i][5]
		# 设备型号
		device_version = ret1[i][12]
		# 配置
		config = ret1[i][13]
		# 使用人
		use_user = ret1[i][17]
		# 使用部门
		depart = ret1[i][18]
		depart = Depart(depart,cur)
		# 设备状态
		device_status = ret1[i][19]
		device_status = deviceStatus(device_status) 
		# 领用时间
		receive_time = ret1[i][20]
		if receive_time == 0:
			receive_time = ''
		else:
			receive_time = str_time(receive_time)
		# 归还时间
		return_time = ret1[i][21]
		if return_time == 0:
			return_time = ''
		else:
			return_time = str_time(return_time)
		# 购买时间
		buy_time = ret1[i][22]
		if buy_time == 0:
			buy_time = ''
		else:
			buy_time = str_time(buy_time)
		# 备注
		remark = ret1[i][24]

		write_sheet.write(i+1,0,asset_belong)
		write_sheet.write(i+1,1,asset_number)
		write_sheet.write(i+1,2,asset_type)
		write_sheet.write(i+1,3,device_type)
		write_sheet.write(i+1,4,device_brand)
		write_sheet.write(i+1,5,device_version)
		write_sheet.write(i+1,6,config)
		write_sheet.write(i+1,7,use_user)
		write_sheet.write(i+1,8,depart)
		write_sheet.write(i+1,9,device_status)
		write_sheet.write(i+1,10,receive_time)
		write_sheet.write(i+1,11,return_time)
		write_sheet.write(i+1,12,buy_time)
		write_sheet.write(i+1,13,remark)
	# 保存表格
	desxls.save(filename)
	cur.close()
	conn.close()
	return 0


# 手绘板
def excel_wacom(filedir):

	conn = mysql_con()
	cur = conn.cursor()
	comm = 'select * from asset_device where device_type="3" '
	cur.execute(comm)
	ret1 = cur.fetchall()
	filedir = filedir
	filename = filedir+r'\手绘板.xls'

	# 设置标题颜色
	styleBoldRed  =xlwt.easyxf('font: color-index black, bold on')
	# 新建表格
	desxls = xlwt.Workbook(encoding='utf-8')
	write_sheet = desxls.add_sheet("sheet1",cell_overwrite_ok=True)

	headlist = ["资产所属","资产编号","资产类型",
	"设备类别","设备品牌","设备型号",
	"配置","使用人","使用部门","设备状态","领用时间",
	"归还时间","购买时间","备注"]

	for i in range(len(headlist)):
		write_sheet.write(0,i,headlist[i],styleBoldRed)

	for i in range(len(ret1)):
		# 资产所属
		asset_belong = ret1[i][1]
		asset_belong = assetBelong(asset_belong)
		# 资产编号
		asset_number = ret1[i][2]
		# 资产类型
		asset_type = ret1[i][3]
		asset_type = assetType(asset_type)
		# 设备类别
		device_type = ret1[i][4]
		device_type = deviceType(device_type) 
		# 设备品牌
		device_brand = ret1[i][5]
		# 设备型号
		device_version = ret1[i][12]
		# 配置
		config = ret1[i][13]
		# 使用人
		use_user = ret1[i][17]
		# 使用部门
		depart = ret1[i][18]
		depart = Depart(depart,cur)
		# 设备状态
		device_status = ret1[i][19]
		device_status = deviceStatus(device_status) 
		# 领用时间
		receive_time = ret1[i][20]
		if receive_time == 0:
			receive_time = ''
		else:
			receive_time = str_time(receive_time)
		# 归还时间
		return_time = ret1[i][21]
		if return_time == 0:
			return_time = ''
		else:
			return_time = str_time(return_time)
		# 购买时间
		buy_time = ret1[i][22]
		if buy_time == 0:
			buy_time = ''
		else:
			buy_time = str_time(buy_time)
		# 备注
		remark = ret1[i][24]

		write_sheet.write(i+1,0,asset_belong)
		write_sheet.write(i+1,1,asset_number)
		write_sheet.write(i+1,2,asset_type)
		write_sheet.write(i+1,3,device_type)
		write_sheet.write(i+1,4,device_brand)
		write_sheet.write(i+1,5,device_version)
		write_sheet.write(i+1,6,config)
		write_sheet.write(i+1,7,use_user)
		write_sheet.write(i+1,8,depart)
		write_sheet.write(i+1,9,device_status)
		write_sheet.write(i+1,10,receive_time)
		write_sheet.write(i+1,11,return_time)
		write_sheet.write(i+1,12,buy_time)
		write_sheet.write(i+1,13,remark)
	# 保存表格
	desxls.save(filename)
	cur.close()
	conn.close()
	return 0

# 手机
def excel_phone(filedir):

	conn = mysql_con()
	cur = conn.cursor()
	comm = 'select * from asset_device where device_type="4" '
	cur.execute(comm)
	ret1 = cur.fetchall()
	filedir = filedir
	filename = filedir+r'\手机.xls'

	# 设置标题颜色
	styleBoldRed  =xlwt.easyxf('font: color-index black, bold on')
	# 新建表格
	desxls = xlwt.Workbook(encoding='utf-8')
	write_sheet = desxls.add_sheet("sheet1",cell_overwrite_ok=True)

	headlist = ["资产所属","资产编号","资产类型",
	"设备类别","设备品牌","设备型号","配置","IMEI码","MAC地址",
	"使用人","使用部门","设备状态","领用时间",
	"归还时间","购买时间","备注"]

	for i in range(len(headlist)):
		write_sheet.write(0,i,headlist[i],styleBoldRed)

	for i in range(len(ret1)):
		# 资产所属
		asset_belong = ret1[i][1]
		asset_belong = assetBelong(asset_belong)
		# 资产编号
		asset_number = ret1[i][2]
		# 资产类型
		asset_type = ret1[i][3]
		asset_type = assetType(asset_type)
		# 设备类别
		device_type = ret1[i][4]
		device_type = deviceType(device_type) 
		# 设备品牌
		device_brand = ret1[i][5]
		# 设备型号
		device_version = ret1[i][12]
		# 配置
		config = ret1[i][13]
		# IMEI码
		imei = ret1[i][14]
		# MAC地址
		mac = ret1[i][15]
		# 使用人
		use_user = ret1[i][17]
		# 使用部门
		depart = ret1[i][18]
		depart = Depart(depart,cur)
		# 设备状态
		device_status = ret1[i][19]
		device_status = deviceStatus(device_status) 
		# 领用时间
		receive_time = ret1[i][20]
		if receive_time == 0:
			receive_time = ''
		else:
			receive_time = str_time(receive_time)
		# 归还时间
		return_time = ret1[i][21]
		if return_time == 0:
			return_time = ''
		else:
			return_time = str_time(return_time)
		# 购买时间
		buy_time = ret1[i][22]
		if buy_time == 0:
			buy_time = ''
		else:
			buy_time = str_time(buy_time)
		# 备注
		remark = ret1[i][24]

		write_sheet.write(i+1,0,asset_belong)
		write_sheet.write(i+1,1,asset_number)
		write_sheet.write(i+1,2,asset_type)
		write_sheet.write(i+1,3,device_type)
		write_sheet.write(i+1,4,device_brand)
		write_sheet.write(i+1,5,device_version)
		write_sheet.write(i+1,6,config)
		write_sheet.write(i+1,7,imei)
		write_sheet.write(i+1,8,mac)
		write_sheet.write(i+1,9,use_user)
		write_sheet.write(i+1,10,depart)
		write_sheet.write(i+1,11,device_status)
		write_sheet.write(i+1,12,receive_time)
		write_sheet.write(i+1,13,return_time)
		write_sheet.write(i+1,14,buy_time)
		write_sheet.write(i+1,15,remark)
	# 保存表格
	desxls.save(filename)
	cur.close()
	conn.close()
	return 0

# 平板
def excel_pad(filedir):

	conn = mysql_con()
	cur = conn.cursor()
	comm = 'select * from asset_device where device_type="5"  '
	cur.execute(comm)
	ret1 = cur.fetchall()
	filedir = filedir
	filename = filedir+r'\平板.xls'

	# 设置标题颜色
	styleBoldRed  =xlwt.easyxf('font: color-index black, bold on')
	# 新建表格
	desxls = xlwt.Workbook(encoding='utf-8')
	write_sheet = desxls.add_sheet("sheet1",cell_overwrite_ok=True)

	headlist = ["资产所属","资产编号","资产类型",
	"设备类别","设备品牌","设备型号","配置","MAC地址",
	"使用人","使用部门","设备状态","领用时间",
	"归还时间","购买时间","备注"]

	for i in range(len(headlist)):
		write_sheet.write(0,i,headlist[i],styleBoldRed)

	for i in range(len(ret1)):
		# 资产所属
		asset_belong = ret1[i][1]
		asset_belong = assetBelong(asset_belong)
		# 资产编号
		asset_number = ret1[i][2]
		# 资产类型
		asset_type = ret1[i][3]
		asset_type = assetType(asset_type)
		# 设备类别
		device_type = ret1[i][4]
		device_type = deviceType(device_type) 
		# 设备品牌
		device_brand = ret1[i][5]
		# 设备型号
		device_version = ret1[i][12]
		# 配置
		config = ret1[i][13]
		# MAC地址
		mac = ret1[i][15]
		# 使用人
		use_user = ret1[i][17]
		# 使用部门
		depart = ret1[i][18]
		depart = Depart(depart,cur)
		# 设备状态
		device_status = ret1[i][19]
		device_status = deviceStatus(device_status) 
		# 领用时间
		receive_time = ret1[i][20]
		if receive_time == 0:
			receive_time = ''
		else:
			receive_time = str_time(receive_time)
		# 归还时间
		return_time = ret1[i][21]
		if return_time == 0:
			return_time = ''
		else:
			return_time = str_time(return_time)
		# 购买时间
		buy_time = ret1[i][22]
		if buy_time == 0:
			buy_time = ''
		else:
			buy_time = str_time(buy_time)
		# 备注
		remark = ret1[i][24]

		write_sheet.write(i+1,0,asset_belong)
		write_sheet.write(i+1,1,asset_number)
		write_sheet.write(i+1,2,asset_type)
		write_sheet.write(i+1,3,device_type)
		write_sheet.write(i+1,4,device_brand)
		write_sheet.write(i+1,5,device_version)
		write_sheet.write(i+1,6,config)
		write_sheet.write(i+1,7,mac)
		write_sheet.write(i+1,8,use_user)
		write_sheet.write(i+1,9,depart)
		write_sheet.write(i+1,10,device_status)
		write_sheet.write(i+1,11,receive_time)
		write_sheet.write(i+1,12,return_time)
		write_sheet.write(i+1,13,buy_time)
		write_sheet.write(i+1,14,remark)
	# 保存表格
	desxls.save(filename)
	cur.close()
	conn.close()
	return 0

# 笔记本
def excel_notebook(filedir):

	conn = mysql_con()
	cur = conn.cursor()
	comm = 'select * from asset_device where device_type="6" '
	cur.execute(comm)
	ret1 = cur.fetchall()
	filedir = filedir
	filename = filedir+r'\笔记本.xls'

	# 设置标题颜色
	styleBoldRed  =xlwt.easyxf('font: color-index black, bold on')
	# 新建表格
	desxls = xlwt.Workbook(encoding='utf-8')
	write_sheet = desxls.add_sheet("sheet1",cell_overwrite_ok=True)

	headlist = ["资产所属","资产编号","资产类型",
	"设备类别","设备品牌","设备型号","CPU","主板",
	"显卡","硬盘","内存","配置","MAC地址",
	"其他","使用人","使用部门","设备状态","领用时间",
	"归还时间","购买时间","备注"]

	for i in range(len(headlist)):
		write_sheet.write(0,i,headlist[i],styleBoldRed)

	for i in range(len(ret1)):
		# 资产所属
		asset_belong = ret1[i][1]
		asset_belong = assetBelong(asset_belong)
		# 资产编号
		asset_number = ret1[i][2]
		# 资产类型
		asset_type = ret1[i][3]
		asset_type = assetType(asset_type)
		# 设备类别
		device_type = ret1[i][4]
		device_type = deviceType(device_type) 
		# 设备品牌
		device_brand = ret1[i][5]
		# 设备型号
		device_version = ret1[i][12]
		# CPU
		cpu = ret1[i][7]
		# 主板
		computer_board = ret1[i][8]
		# 显卡
		displya_card = ret1[i][9]
		# 硬盘
		hard_disk = ret1[i][10]
		# 内存
		memory = ret1[i][11]
		# 配置
		config = ret1[i][13]
		# MAC地址
		mac = ret1[i][15]
		# 其他
		other = ret1[i][16]
		# 使用人
		use_user = ret1[i][17]
		# 使用部门
		depart = ret1[i][18]
		depart = Depart(depart,cur)
		# 设备状态
		device_status = ret1[i][19]
		device_status = deviceStatus(device_status) 
		# 领用时间
		receive_time = ret1[i][20]
		if receive_time == 0:
			receive_time = ''
		else:
			receive_time = str_time(receive_time)
		# 归还时间
		return_time = ret1[i][21]
		if return_time == 0:
			return_time = ''
		else:
			return_time = str_time(return_time)
		# 购买时间
		buy_time = ret1[i][22]
		if buy_time == 0:
			buy_time = ''
		else:
			buy_time = str_time(buy_time)
		# 备注
		remark = ret1[i][24]

		write_sheet.write(i+1,0,asset_belong)
		write_sheet.write(i+1,1,asset_number)
		write_sheet.write(i+1,2,asset_type)
		write_sheet.write(i+1,3,device_type)
		write_sheet.write(i+1,4,device_brand)
		write_sheet.write(i+1,5,device_version)
		write_sheet.write(i+1,6,cpu)
		write_sheet.write(i+1,7,computer_board)
		write_sheet.write(i+1,8,displya_card)
		write_sheet.write(i+1,9,hard_disk)
		write_sheet.write(i+1,10,memory)
		write_sheet.write(i+1,11,config)
		write_sheet.write(i+1,12,mac)
		write_sheet.write(i+1,13,other)
		write_sheet.write(i+1,14,use_user)
		write_sheet.write(i+1,15,depart)
		write_sheet.write(i+1,16,device_status)
		write_sheet.write(i+1,17,receive_time)
		write_sheet.write(i+1,18,return_time)
		write_sheet.write(i+1,19,buy_time)
		write_sheet.write(i+1,20,remark)
	# 保存表格
	desxls.save(filename)
	cur.close()
	conn.close()
	return 0

# IMAC
def excel_iMac(filedir):

	conn = mysql_con()
	cur = conn.cursor()
	comm = 'select * from asset_device where device_type="7" '
	cur.execute(comm)
	ret1 = cur.fetchall()

	filedir = filedir
	filename = filedir+r'\iMac.xls'

	# 设置标题颜色
	styleBoldRed  =xlwt.easyxf('font: color-index black, bold on')
	# 新建表格
	desxls = xlwt.Workbook(encoding='utf-8')
	write_sheet = desxls.add_sheet("sheet1",cell_overwrite_ok=True)

	headlist = ["资产所属","资产编号","资产类型",
	"设备类别","设备品牌","设备型号","CPU","主板",
	"显卡","硬盘","内存","其他","使用人","使用部门","设备状态","领用时间",
	"归还时间","购买时间","备注"]

	for i in range(len(headlist)):
		write_sheet.write(0,i,headlist[i],styleBoldRed)

	for i in range(len(ret1)):
		# 资产所属
		asset_belong = ret1[i][1]
		asset_belong = assetBelong(asset_belong)
		# 资产编号
		asset_number = ret1[i][2]
		# 资产类型
		asset_type = ret1[i][3]
		asset_type = assetType(asset_type)
		# 设备类别
		device_type = ret1[i][4]
		device_type = deviceType(device_type) 
		# 设备品牌
		device_brand = ret1[i][5]
		# 设备型号
		device_version = ret1[i][12]
		# CPU
		cpu = ret1[i][7]
		# 主板
		computer_board = ret1[i][8]
		# 显卡
		displya_card = ret1[i][9]
		# 硬盘
		hard_disk = ret1[i][10]
		# 内存
		memory = ret1[i][11]
		# 其他
		other = ret1[i][16]
		# 使用人
		use_user = ret1[i][17]
		# 使用部门
		depart = ret1[i][18]
		depart = Depart(depart,cur)
		# 设备状态
		device_status = ret1[i][19]
		device_status = deviceStatus(device_status) 
		# 领用时间
		receive_time = ret1[i][20]
		if receive_time == 0:
			receive_time = ''
		else:
			receive_time = str_time(receive_time)
		# 归还时间
		return_time = ret1[i][21]
		if return_time == 0:
			return_time = ''
		else:
			return_time = str_time(return_time)
		# 购买时间
		buy_time = ret1[i][22]
		if buy_time == 0:
			buy_time = ''
		else:
			buy_time = str_time(buy_time)
		# 备注
		remark = ret1[i][24]

		write_sheet.write(i+1,0,asset_belong)
		write_sheet.write(i+1,1,asset_number)
		write_sheet.write(i+1,2,asset_type)
		write_sheet.write(i+1,3,device_type)
		write_sheet.write(i+1,4,device_brand)
		write_sheet.write(i+1,5,device_version)
		write_sheet.write(i+1,6,cpu)
		write_sheet.write(i+1,7,computer_board)
		write_sheet.write(i+1,8,displya_card)
		write_sheet.write(i+1,9,hard_disk)
		write_sheet.write(i+1,10,memory)
		write_sheet.write(i+1,11,other)
		write_sheet.write(i+1,12,use_user)
		write_sheet.write(i+1,13,depart)
		write_sheet.write(i+1,14,device_status)
		write_sheet.write(i+1,15,receive_time)
		write_sheet.write(i+1,16,return_time)
		write_sheet.write(i+1,17,buy_time)
		write_sheet.write(i+1,18,remark)
	# 保存表格
	desxls.save(filename)
	cur.close()
	conn.close()
	return 0


# 其他
def excel_other(filedir):

	conn = mysql_con()
	cur = conn.cursor()
	comm = 'select * from asset_device where device_type="8" '
	cur.execute(comm)
	ret1 = cur.fetchall()

	filedir = filedir
	filename = filedir+r'\其他.xls'

	# 设置标题颜色
	styleBoldRed  =xlwt.easyxf('font: color-index black, bold on')
	# 新建表格
	desxls = xlwt.Workbook(encoding='utf-8')
	write_sheet = desxls.add_sheet("sheet1",cell_overwrite_ok=True)

	headlist = ["资产所属","资产编号","资产类型",
	"设备类别","设备品牌","设备名称","设备型号",
	"配置","使用人","使用部门","设备状态","领用时间",
	"归还时间","购买时间","备注"]

	for i in range(len(headlist)):
		write_sheet.write(0,i,headlist[i],styleBoldRed)

	for i in range(len(ret1)):
		# 资产所属
		asset_belong = ret1[i][1]
		asset_belong = assetBelong(asset_belong)
		# 资产编号
		asset_number = ret1[i][2]
		# 资产类型
		asset_type = ret1[i][3]
		asset_type = assetType(asset_type)
		# 设备类别
		device_type = ret1[i][4]
		device_type = deviceType(device_type) 
		# 设备品牌
		device_brand = ret1[i][5]
		# 设备名称
		device_name = ret1[i][6]
		# 设备型号
		device_version = ret1[i][12]
		# 配置
		config = ret1[i][13]
		# 使用人
		use_user = ret1[i][17]
		# 使用部门
		depart = ret1[i][18]
		depart = Depart(depart,cur)
		# 设备状态
		device_status = ret1[i][19]
		device_status = deviceStatus(device_status) 
		# 领用时间
		receive_time = ret1[i][20]
		if receive_time == 0:
			receive_time = ''
		else:
			receive_time = str_time(receive_time)
		# 归还时间
		return_time = ret1[i][21]
		if return_time == 0:
			return_time = ''
		else:
			return_time = str_time(return_time)
		# 购买时间
		buy_time = ret1[i][22]
		if buy_time == 0:
			buy_time = ''
		else:
			buy_time = str_time(buy_time)
		# 备注
		remark = ret1[i][24]

		write_sheet.write(i+1,0,asset_belong)
		write_sheet.write(i+1,1,asset_number)
		write_sheet.write(i+1,2,asset_type)
		write_sheet.write(i+1,3,device_type)
		write_sheet.write(i+1,4,device_brand)
		write_sheet.write(i+1,5,device_name)
		write_sheet.write(i+1,6,device_version)
		write_sheet.write(i+1,7,config)
		write_sheet.write(i+1,8,use_user)
		write_sheet.write(i+1,9,depart)
		write_sheet.write(i+1,10,device_status)
		write_sheet.write(i+1,11,receive_time)
		write_sheet.write(i+1,12,return_time)
		write_sheet.write(i+1,13,buy_time)
		write_sheet.write(i+1,14,remark)
	# 保存表格
	desxls.save(filename)
	cur.close()
	conn.close()
	return 0



if __name__ == '__main__' :

	# 创建文件夹文件夹
	filedir = rootdir+'\\'+datetime.datetime.now().strftime('%Y-%m-%d')
	if not os.path.isdir(filedir):
		os.mkdir(filedir)

	if not os.path.isdir(logdir):
		os.mkdir(logdir)

	try:
		r_computer = excel_computer(filedir)
		r_display = excel_display(filedir)
		r_wacom = excel_wacom(filedir)
		r_phone = excel_phone(filedir)
		r_pad = excel_pad(filedir)
		r_notebook = excel_notebook(filedir)
		r_iMac = excel_iMac(filedir)
		r_other = excel_other(filedir)

		excel_log = '[Export] Export excel is Successfully. The file is In the directory of the ['+filedir+']. -------------- '+datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')+'\n'
		with open(excel_logfile,'a') as f:
			f.write(excel_log)
	except:
		error_log = '[Error] Export excel is faild. -------------------'+datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')+'\n'
		with open(error_logfile,'a') as f:
			f.write(error_log)

