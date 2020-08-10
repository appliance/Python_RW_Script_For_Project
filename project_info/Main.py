from dbHelper import DBHelper
from project_info_2_database import ImportProjectInfo
from db_config import db_config
import tkinter

def main():
    dbHepler = DBHelper()
    # 连接数据库
    dbHepler.connect_database()

    importProInfo = ImportProjectInfo('D:/Desktop/7.31_最新_项目综合信息表.xlsx')
    sheet_info = importProInfo.get_sheet_by_name('项目综合信息表')
    wbs_list_from_project_info = importProInfo.get_wbs_list(sheet_info, '项目WBS号')
    # print(wbs_list_from_project_info)

    # 获取prioject_dir
    for i in range(len(wbs_list_from_project_info)):
        project_dir = importProInfo.get_project_info_by_wbs_info(sheet_info, wbs_list_from_project_info[i])
        print(project_dir)

        sql, params = importProInfo.create_update_sql_by_dir(project_dir)
        print(sql)
        print(params)

        dbHepler.excute(sql, list(params))

    dbHepler.close_database()

    print('success')


main()