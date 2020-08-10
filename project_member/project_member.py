import xlrd
import pymysql
import tkinter

#更新信息，可知道录入是否成功
waiting = 'wait a minute...'

url = ''
url2 = ''
hostname = ''
username = ''
password = ''
dbname = ''
portnum = ''
charset = ''

def start1():
    global waiting
    global url
    global url2
    global hostname
    global username
    global password
    global dbname
    global portnum
    global charset
    # 打开数据所在的路径表名
    book = xlrd.open_workbook(url)
    # 获取excel文件中所有sheet表的名称
    sheets = book.sheet_names()
    # 这个是表里的sheet名称（注意大小写）
    sheet = book.sheet_by_index(int(url2))

    # 获取项目名称(第二行第二列）
    if '：' in sheet.cell(2, 1).value:
        project_name = str(sheet.cell(2, 1).value).split('：')[1]
    else:
        project_name = ''
    if '：' in sheet.cell(1, 1).value:
        project_wbs = str(sheet.cell(1, 1).value).split('：')[1]
    else:
        project_wbs = ''



    # 建立一个 MySQL连接

    conn = pymysql.connect(
        host=hostname,
        user=username,
        passwd=password,
        db=dbname,
        port=int(portnum),
        charset=charset
    )

    # 获得游标
    cur = conn.cursor()

    # 检索出pid
    if project_wbs != '':
        query = 'select id from project where WBS=%s;'
        param = project_wbs
        cur.execute(query, param)
        result = cur.fetchall()
        if result != ():
            pid = result[0][0]
        else:
            pid = None
    elif project_name != '':
        query = 'select id from project where project_name = %s;'
        param = project_name
        cur.execute(query, param)
        result = cur.fetchall()
        if result != ():
            pid = result[0][0]
        else:
            pid = None
    else:
        pid = None



    # 创建插入sql语句
    #query = 'insert into employee(WBS_NUM,PRO_NAME,DEPT_NAME,DE_MAN,EMP_NAME)values(%s,%s,%s,%s,%s)'

    # 表信息比对字典
    list = {'姓名':'name','性别':'sex','出生年月':'birthday','工龄\n（年）':'work_age',
            '学历':'education_background','外包部门':'department','人员类别':'category',
            '人员级别':'level','人员单价':'price','成员来源':'source','所属外包公司':'outsource_company',
            '是否本地化':'localize','岗位名称':'post','人员状态':'status','岗位类型':'job_type',
            '考勤\n状态':'attendance_status','本项目工作职责':'responsibility','计划开\n始时间':'planned_start_time',
            '进入项目时间':'project_entry_time','计划结\n束时间':'planned_finish_time',
            '离开项目时间\n（在岗人员不填，离岗人员由项目经理填写）':'leave_project_time',
            '离开\n原因':'leave_reason','在岗时长约\n（月）':'on_duty_months','工作地':'address',
            '备注':'remark','说明':'p_explain'}
    # 数据类型比对字典
    listk = {'姓名':'str','性别':'str','出生年月':'date','工龄\n（年）':'str',
            '学历':'str','外包部门':'str','人员类别':'str',
            '人员级别':'str','人员单价':'str','成员来源':'str','所属外包公司':'str',
            '是否本地化':'str','岗位名称':'str','人员状态':'str','岗位类型':'str',
            '考勤\n状态':'str','本项目工作职责':'str','计划开\n始时间':'date',
            '进入项目时间':'date','计划结\n束时间':'date',
            '离开项目时间\n（在岗人员不填，离岗人员由项目经理填写）':'date',
            '离开\n原因':'str','在岗时长约\n（月）':'int','地点':'str',
            '备注':'str','说明':'str'}


    # 录入int类型数据
    # a用来存放需要输入的数据
    a_int = []
    a_str = []
    a_date = []
    b_int = []
    c_int = []
    b_str = []
    c_str = []
    b_date = []
    c_date = []
    # 写好insert语句
    for i in range(1,sheet.ncols):
        if sheet.cell(3,i).value in list :
            #print(listk[sheet.cell(3,i).value])
            if sheet.cell(3,i).value in listk and listk[sheet.cell(3,i).value] == 'int':
                c_int.append('%s')
                b_int.append(list[sheet.cell(3, i).value])
    # 添加pid
    if pid != None:
        c_int.append('%s')
        b_int.append('pid')
    int_join = ','.join(b_int)
    int_join2 = ','.join(c_int)


    # 录入str类型数据
    for i in range(1,sheet.ncols):
        if sheet.cell(3,i).value in list :
            #print(listk[sheet.cell(3,i).value])
            if sheet.cell(3,i).value in listk and listk[sheet.cell(3,i).value] == 'str':
                c_str.append('%s')
                b_str.append(list[sheet.cell(3, i).value])
    str_join = ','.join(b_str)
    str_join2 = ','.join(c_str)


    # 录入date类型数据
    for i in range(1, sheet.ncols):
        if sheet.cell(3, i).value in list:
            #print(listk[sheet.cell(3,i).value])
            if sheet.cell(3, i).value in listk and listk[sheet.cell(3,i).value] == 'date':
                c_date.append('%s')
                b_date.append(list[sheet.cell(3, i).value])
    date_join = ','.join(b_date)
    date_join2 = ','.join(c_date)

    query = 'insert into employee('+int_join+','+str_join+','+date_join+')values('+int_join2+','+str_join2+','+date_join2+');'
    print(query)
    print(b_int)
    print(b_str)
    print(b_date)

    # 创建一个for循环迭代读取xls文件每行数据的，
    for r in range(4, sheet.nrows):
        # 查找int类型数据
        for k in range(1,sheet.ncols):
            if sheet.cell(3,k).value in list and list[sheet.cell(3,k).value] in b_int:
                if sheet.cell(r,k).value == '':
                    a_int.append(None)
                else:
                    a_int.append(int(sheet.cell(r, k).value))
        # 添加pid
        if pid != None:
            a_int.append(pid)

        # 查找str类型数据
        for k in range(1, sheet.ncols):
            if sheet.cell(3, k).value in list and list[sheet.cell(3, k).value] in b_str:
                    a_str.append(str(sheet.cell(r, k).value))

        # 查找date类型数据
        for k in range(1, sheet.ncols):
            if sheet.cell(3, k).value in list and list[sheet.cell(3, k).value] in b_date:
                if(sheet.cell(r,k).value != ''):
                    x = xlrd.xldate.xldate_as_datetime(sheet.cell(r,k).value, 0)
                else:
                    x = None
                a_date.append(x)
        values = a_int+a_str+a_date
        print(values)
        #print(values)
        cur.execute(query, values)
        # 清空a
        a_int = []
        a_str = []
        a_date = []

    #print(values)
    # close关闭文档
    cur.close()
    # commit 提交
    conn.commit()
    # 关闭MySQL链接
    conn.close()
    # 显示导入多少列
    columns = str(sheet.ncols)
    # 显示导入多少行
    rows = str(sheet.nrows)
    print('导入'+columns+'列'+rows+'行数据到MySQL数据库!')
    waiting = 'success!'

def start2():
    global waiting
    global url
    global url2
    global hostname
    global username
    global password
    global dbname
    global portnum
    global charset
    # 打开数据所在的路径表名
    book = xlrd.open_workbook(url)
    # 获取excel文件中所有sheet表的名称
    sheets = book.sheet_names()
    sheet = book.sheet_by_name('外包人员信息表')


# 建立一个 MySQL连接
    conn = pymysql.connect(
        host=hostname,
        user=username,
        passwd=password,
        db=dbname,
        port=int(portnum),
        charset=charset
    )

    # 获得游标
    cur = conn.cursor()

    # 录入int类型数据
    # a用来存放需要输入的数据
    a_int = []
    a_str = []
    a_date = []
    a_name = []
    b_int = []
    c_int = []
    b_str = []
    c_str = []
    b_date = []
    c_date = []
    # 表信息比对字典
    list = {'姓名': 'name', '性别': 'sex', '出生年月': 'birthday', '工龄\n（年）': 'work_age',
            '学历': 'education_background', '外包部门': 'department', '人员类别': 'category',
            '人员级别': 'level', '人员单价': 'price', '成员来源': 'source', '所属外包公司': 'outsource_company',
            '是否本地化': 'localize', '岗位名称': 'post', '人员状态': 'status', '岗位类型': 'job_type',
            '考勤\n状态': 'attendance_status', '本项目工作职责': 'responsibility', '计划开\n始时间': 'planned_start_time',
            '进入项目时间': 'project_entry_time', '计划结\n束时间': 'planned_finish_time',
            '离开项目时间\n（在岗人员不填，离岗人员由项目经理填写）': 'leave_project_time',
            '离开\n原因': 'leave_reason', '在岗时长约\n（月）': 'on_duty_months', '工作地': 'address',
            '备注': 'remark', '说明': 'p_explain', '外包人员': 'name'}
    # 数据类型比对字典
    listk = {'姓名': 'str', '性别': 'str', '出生年月': 'date', '工龄\n（年）': 'str',
             '学历': 'str', '外包部门': 'str', '人员类别': 'str',
             '人员级别': 'str', '人员单价': 'str', '成员来源': 'str', '所属外包公司': 'str',
             '是否本地化': 'str', '岗位名称': 'str', '人员状态': 'str', '岗位类型': 'str',
             '考勤\n状态': 'str', '本项目工作职责': 'str', '计划开\n始时间': 'date',
             '进入项目时间': 'date', '计划结\n束时间': 'date',
             '离开项目时间\n（在岗人员不填，离岗人员由项目经理填写）': 'date',
             '离开\n原因': 'str', '在岗时长约\n（月）': 'int', '地点': 'str',
             '备注': 'str', '说明': 'str', '外包人员':'str'}
    # 写好insert语句
    for i in range(0,sheet.ncols):
        if sheet.cell(0,i).value in list :
            #print(listk[sheet.cell(3,i).value])
            if sheet.cell(0,i).value in listk and listk[sheet.cell(0,i).value] == 'int':
                c_int.append('%s')
                b_int.append(list[sheet.cell(0, i).value])
    int_join = ','.join(b_int)
    int_join2 = ','.join(c_int)

    # 录入str类型数据
    for i in range(0,sheet.ncols):
        if sheet.cell(0,i).value in list :
            #print(listk[sheet.cell(3,i).value])
            if sheet.cell(0,i).value in listk and listk[sheet.cell(0,i).value] == 'str' and list[sheet.cell(0,i).value] != 'name':
                c_str.append('%s')
                b_str.append(list[sheet.cell(0, i).value])
    str_join = ','.join(b_str)
    str_join2 = ','.join(c_str)

    # 录入date类型数据
    for i in range(0,sheet.ncols):
        if sheet.cell(0,i).value in list :
            #print(listk[sheet.cell(3,i).value])
            if sheet.cell(0,i).value in listk and listk[sheet.cell(0,i).value] == 'date':
                c_date.append('%s')
                b_date.append(list[sheet.cell(0, i).value])
    date_join = ','.join(b_date)
    date_join2 = ','.join(c_date)

    print(len(b_int))
    print(len(b_str))
    print(len(b_date))
    # if len(b_int) != 0 and len(b_str) != 0 and len(b_date) != 0 :
    #     query = 'if noexists (select * from employee where name = %s) insert into employee('+int_join+','+str_join+','+date_join+')values('+int_join2+','+str_join2+','+date_join2+')'
    # elif len(b_int) == 0 and len(b_str) != 0 and len(b_date) == 0 :
    #     query = 'insert into employee(' + str_join + ') select '+str_join2+ 'from employee where name = %s'
    #print(query)
    print(b_int)
    print(b_str)
    print(b_date)

    # 创建一个for循环迭代读取xls文件每行数据的，
    for r in range(1, sheet.nrows):
        # 查找int类型数据
        for k in range(0,sheet.ncols):
            if sheet.cell(0,k).value in list and list[sheet.cell(0,k).value] in b_int:
                if sheet.cell(r,k).value == '':
                    a_int.append(None)
                else:
                    a_int.append(int(sheet.cell(r, k).value))
        # 查找str类型数据
        for k in range(0, sheet.ncols):
            if sheet.cell(0, k).value in list and list[sheet.cell(0, k).value] == 'name':
                a_name.append(str(sheet.cell(r, k).value))
            elif sheet.cell(0, k).value in list and list[sheet.cell(0, k).value] in b_str:
                a_str.append(str(sheet.cell(r, k).value))

        # 查找date类型数据
        for k in range(0, sheet.ncols):
            if sheet.cell(0, k).value in list and list[sheet.cell(0, k).value] in b_date:
                if(sheet.cell(r,k).value != ''):
                    x = xlrd.xldate.xldate_as_datetime(sheet.cell(r,k).value, 0)
                else:
                    x = None
                a_date.append(x)
        values = a_int+a_str+a_date+a_name
        #cur.execute('insert into employee(department,category,level,price,localize)values(%s,%s,%s,%s,%s) where name =%s')
        #cur.execute(query, values)
        print(b_str)
        print(a_str)
        query = 'update employee set ' + "'"+ b_str[0] +"'"+ '= ' +"'"+ a_str[0] +"'"+ ' where name = ' +"'"+ a_name[0]+"'"
        print(query)
        for i in range(0,len(b_str)):
            cur.execute( 'update employee set ' + b_str[i] + '= ' +"'"+ a_str[i] +"'"+ ' where name = ' +"'"+ a_name[0]+"'")
        # 清空a
        a_int = []
        a_str = []
        a_date = []
        a_name = []
        print(values)
    #print(values)
    # close关闭文档
    cur.close()
    # commit 提交
    conn.commit()
    # 关闭MySQL链接
    conn.close()
    # 显示导入多少列
    columns = str(sheet.ncols)
    # 显示导入多少行
    rows = str(sheet.nrows)
    print('导入'+columns+'列'+rows+'行数据到MySQL数据库!')
    waiting = 'success!'

def a():
    global waiting
    global url
    global url2
    global hostname
    global username
    global password
    global dbname
    global portnum
    global charset
    #print(t1.get())
    url = t1.get()    #emp.xls
    url2 = t2.get()    #test2
    hostname = t3.get() #localhost
    username = t4.get() #root
    password = t5.get() #root
    dbname = t6.get() #test
    portnum = t7.get() #3306
    charset = t8.get() #utf8
    #print(waiting)
    #top.update()
    print(url)
    print(url2)
    print(hostname)
    print(username)
    print(password)
    print(dbname)
    print(portnum)
    print(charset)
    if url2 == '外包人员信息表':
        start2()
    else:
        start1()


#刷新函数
def refreshText():
    msg.delete(0.0, tkinter.END)
    msg.insert("insert", waiting)
    msg.update()
    top.after(500, refreshText)

top = tkinter.Tk()
#设置窗口的大小宽x高+偏移量
top.geometry('300x500+500+200')
#设置窗口标题
top.title('Excel To Mysql')

# url
w = tkinter.Label(top, text="请输入文件名：").pack()
t1 = tkinter.StringVar()
t1.set('pro_member.xlsx')
entry = tkinter.Entry(top, textvariable=t1).pack()

# url2
w2 = tkinter.Label(top, text="请输入页名：").pack()
t2 = tkinter.StringVar()
t2.set('test')
entry = tkinter.Entry(top, textvariable = t2).pack()

# hostname
w3 = tkinter.Label(top, text="请输入hostname：").pack()
t3 = tkinter.StringVar()
t3.set('localhost')
entry = tkinter.Entry(top, textvariable = t3).pack()

# username
w4 = tkinter.Label(top, text="请输入username：").pack()
t4 = tkinter.StringVar()
t4.set('root')
entry = tkinter.Entry(top, textvariable = t4).pack()

# password
w5 = tkinter.Label(top, text="请输入password：").pack()
t5 = tkinter.StringVar()
t5.set('root')
entry = tkinter.Entry(top, textvariable = t5).pack()

# dbname
w6 = tkinter.Label(top, text="请输入dbname：").pack()
t6 = tkinter.StringVar()
t6.set('project_control')
entry = tkinter.Entry(top, textvariable = t6).pack()

# portnum
w7 = tkinter.Label(top, text="请输入portnum：").pack()
t7 = tkinter.StringVar()
t7.set('3306')
entry = tkinter.Entry(top, textvariable = t7).pack()

# charset
w8 = tkinter.Label(top, text="请输入charset：").pack()
t8 = tkinter.StringVar()
t8.set('utf8')
entry = tkinter.Entry(top, textvariable = t8).pack()

btn = tkinter.Button(top, text='开始录入',command=a).pack()

msg = tkinter.Text(top)
msg.pack()
top.after(500,refreshText)
top.mainloop()
