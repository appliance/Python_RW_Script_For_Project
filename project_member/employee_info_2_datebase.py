import xlrd
from project_member.dbHelper import DBHelper
from project_member.db_config import db_config
from datetime import date
import re



class ImportEmployeeInfo:
    def __init__(self):
        self.excel_url = db_config['excel_url']

        # excel表属性映射MySQL字段
        self.excel_property_name_2_sql_dir = {
            '姓名': 'name', '性别': 'sex', '出生年月': 'birthday', '工龄\n（年）': 'work_age',
            '学历': 'education_background', '外包部门': 'department', '人员类别': 'category',
            '级别': 'level', '单价\n（元）': 'price', '成员来源': 'source', '所属外包公司': 'outsource_company',
            '是否本地化': 'localize', '岗位名称': 'post', '人员状态': 'status', '岗位类型': 'job_type',
            '考勤\n状态': 'attendance_status', '本项目工作职责': 'responsibility', '计划开\n始时间': 'planned_start_time',
            '进入项目时间': 'project_entry_time', '计划结\n束时间': 'planned_finish_time',
            '离开项目时间\n（在岗人员不填，离岗人员由项目经理填写）': 'leave_project_time',
            '离开\n原因': 'leave_reason', '在岗时长约\n（月）': 'on_duty_months', '地点': 'address',
            '备注': 'remark', '说明': 'p_explain', '打卡项目': 'attendance_id'
        }

        # excel表属性类型映射
        self.excel_property_type_dir = {
            '姓名': 'str', '性别': 'str', '出生年月': 'date', '工龄\n（年）': 'str',
            '学历': 'str', '外包部门': 'str', '人员类别': 'str',
            '级别': 'str', '单价\n（元）': 'int', '成员来源': 'str', '所属外包公司': 'str',
            '是否本地化': 'str', '岗位名称': 'str', '人员状态': 'str', '岗位类型': 'str',
            '考勤\n状态': 'str', '本项目工作职责': 'str', '计划开\n始时间': 'date',
            '进入项目时间': 'date', '计划结\n束时间': 'date',
            '离开项目时间\n（在岗人员不填，离岗人员由项目经理填写）': 'date',
            '离开\n原因': 'str', '在岗时长约\n（月）': 'int', '地点': 'str',
            '备注': 'str', '说明': 'str', '打卡项目': 'int'
        }


    """
        获取excel数据表对象
        :parameter excel_url excel总表地址
        :parameter sheet_name 总表中子表名称
        :return sheet 子表处理对象
    """
    def get_sheet_by_name(self, sheet_name):
        book = xlrd.open_workbook(self.excel_url)
        sheets = book.sheet_names()
        # 表名 和 序列映射
        sheet_name_2_index_dir = dict(zip(sheets, [i for i in range(len(sheets))]))
        if sheet_name in sheet_name_2_index_dir.keys():
            index = sheet_name_2_index_dir[sheet_name]
            sheet = book.sheet_by_index(index)
            return sheet
        else:
            print('警告: 该表中不存在' + sheet_name + '表，请核实表信息！！！')
            return


    '''
        获取pid 
        :param wbs
        :return pid   
    '''
    def get_pid_by_wbs(self, wbs):
        dbHelper = DBHelper()
        dbHelper.connect_database()
        sql = 'select id from project where WBS = %s'
        param = wbs
        res = dbHelper.select(sql, param)
        dbHelper.close_database()
        # project中
        #   不存在WBS返回 ()
        #   存在返回 ((num,))
        if res == ():
            return None
        else:
            return res[0][0]


    '''
        判断待插入记录是否存在
        :param pid
        :param name
        :return True / False
    '''
    def is_exist(self, pid, name):
        dbHelper = DBHelper()
        dbHelper.connect_database()
        sql = 'select count(*) from employee where pid = %s and name = %s'
        param = [pid, name]
        res = dbHelper.select(sql, param)
        dbHelper.close_database()
        if res[0][0] == 0:
            return False
        else:
            return True

    '''
        按行，构造人员信息dir
        :param sheet excle表对象
        :param row 行
        :return {} dir
    '''
    def create_employee_info_dir(self, sheet, row_position):
        employee_info_dir = {}
        for col_index in range(sheet.ncols):
            # 人员信息表结构 从0开始 第3行为属性名
            if sheet.cell(3, col_index).value in self.excel_property_name_2_sql_dir.keys():
                property_set_in_database = self.excel_property_name_2_sql_dir.get(sheet.cell(3, col_index).value)
                # 判断是否为date数据类型,需要进行格式转化
                if self.excel_property_type_dir.get(sheet.cell(3, col_index).value) == 'date':
                    # date 分有数值和空数值两种情况处理
                    if sheet.cell(row_position, col_index).value != '':
                        if sheet.cell(row_position, col_index).value == '/':
                            employee_info_dir[property_set_in_database] = None
                        else:
                            date_cell = xlrd.xldate_as_tuple(sheet.cell(row_position, col_index).value, 0)
                            employee_info_dir[property_set_in_database] = date(*date_cell[0:3]).strftime('%Y-%m-%d')
                    else:
                        employee_info_dir[property_set_in_database] = None
                elif self.excel_property_type_dir.get(sheet.cell(3, col_index).value) == 'int':
                    # int 分空值数值， 非数值
                    if self.is_number(sheet.cell(row_position, col_index).value):
                        employee_info_dir[property_set_in_database] = int(sheet.cell(row_position, col_index).value)
                    else:
                        employee_info_dir[property_set_in_database] = None
                    # str 类型直接处理
                elif self.excel_property_type_dir.get(sheet.cell(3, col_index).value) == 'str':
                    employee_info_dir[property_set_in_database] = str(sheet.cell(row_position, col_index).value)

        return employee_info_dir

    """
        依据dir,构建sql查询语句
        :parameter employee_info_dir 提取的项目信息字典
        :return sql（查询语句）  params（值） 
    """
    def create_insert_sql_by_dir(self, employee_info_dir):
        # 构造insert语句
        sql = 'insert into employee('
        for i in range(len(employee_info_dir.keys())):
            sql += list(employee_info_dir.keys())[i]
            if i == len(employee_info_dir.keys()) - 1:
                break
            else:
                sql += ','
        sql += ')values('
        for i in range(len(employee_info_dir.keys())):
            sql += '%s'
            if i == len(employee_info_dir.keys()) - 1:
                break
            else:
                sql += ','
        sql += ');'

        # 构造parmas
        params = []
        for key in employee_info_dir.keys():
            params.append(employee_info_dir.get(key))
        return sql, list(params)

    """
        判断是否为数值
        :parameter num 
        :return ture/false
    """
    def is_number(self, num):
        pattern = re.compile(r'^[-+]?[-0-9]\d*\.\d*|[-+]?\.?[0-9]\d*$')
        result = pattern.match(str(num))
        if result:
            return True
        else:
            return False


    '''
        插入新项目到project表中,返回pid
        :param wbs
        :return pid
    '''
    def insert_get_pid(self, wbs):
        dbHelper = DBHelper()
        dbHelper.connect_database()
        sql = "insert into project (WBS) values (%s);"
        param = wbs
        dbHelper.excute(sql, param)
        sql = "select id from project where WBS = %s;"
        res = dbHelper.select(sql, param)
        dbHelper.close_database()
        return res[0][0]

