from project_member.employee_info_2_datebase import ImportEmployeeInfo
from project_member.dbHelper import DBHelper
from project_member.db_config import db_config
import xlrd

def main():
    insert_total_num = 0
    importEmployeeInfo = ImportEmployeeInfo()
    dbHelper = DBHelper()
    dbHelper.connect_database()

    # 获取 excel工作表 对象集合
    # sheets : ['智慧云二期-张旭坤', '智慧云运维项目-张旭坤', '国网客服中2019年营配贯通项目-张旭坤', '2020数据中台标签库-张旭坤', '客户域数据中台-程飞飞']
    book = xlrd.open_workbook(db_config['excel_url'])
    sheets = book.sheet_names()

    # 遍历所有表 进行人员信息读写
    for i in range(len(sheets)):
        try:
            sheet = importEmployeeInfo.get_sheet_by_name(sheets[i])

            # 依据表框架 判断是否为项目成员信息表
            if sheet.cell(0, 1).value == '项目成员信息表':
                # 获取wbs号   (WBS：B24623190128)
                if '：' in sheet.cell(1, 1).value:
                    if len(str(sheet.cell(1, 1).value).split('：')) > 1:
                        # WBS：B24623190128
                        wbs = str(sheet.cell(1, 1).value).split('：')[1]
                    else:
                        # WBS：
                        wbs = None;

                # 获取pid
                if wbs != None:
                    pid = importEmployeeInfo.get_pid_by_wbs(wbs)
                    # 若pid 为 None 则将该项目新添加到project中
                    if pid == None:
                        pid = importEmployeeInfo.insert_get_pid(wbs)
                else:
                    pid = None

                # 人员信息表结构 从第4行开始 为人员信息
                for row_index in range(4, sheet.nrows):
                    employee_info_dir = importEmployeeInfo.create_employee_info_dir(sheet=sheet, row_position=row_index)
                    # 补充pid字段信息
                    employee_info_dir['pid'] = pid
                    # 补充打卡字段信息
                    if 'attendance_status' in employee_info_dir.keys():
                        if employee_info_dir['attendance_status'] == '本项目':
                            employee_info_dir['attendance_id'] = pid
                        else:
                            employee_info_dir['attendance_id'] = None
                    else:
                        employee_info_dir['attendance_id'] = None

                    sql, params = importEmployeeInfo.create_insert_sql_by_dir(employee_info_dir=employee_info_dir)
                    # 判断是否插入
                    if importEmployeeInfo.is_exist(pid=pid, name=employee_info_dir['name']) == False:
                        print(employee_info_dir)
                        dbHelper.excute(sql=sql, params=params)
                        # 提示信息
                        insert_total_num += 1
                        print(employee_info_dir)

                print(sheets[i] + '，该表一共录入' + str(sheet.nrows-4) + '条数据。')
                print("****************************************************************************************************")
            else:
                continue
        except IndexError:
            print('人员信息表已插入完毕！')
    print('一共插入 '+ str(insert_total_num) + ' 条记录')


    # 通过
    dbHelper.close_database()





main()