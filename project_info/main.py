from project_info.dbHelper import DBHelper
from project_info.project_info_2_database import ImportProjectInfo
from project_info.db_config import db_config

def main():
    dbHepler = DBHelper()
    # 连接数据库
    dbHepler.connect_database()

    importProInfo = ImportProjectInfo()
    # 获取 ”合同表“ 对象
    sheet_contract = importProInfo.get_sheet_by_name('合同表')

    # 获取 "合同表"中的WBS集合
    wbs_in_contract_list = importProInfo.get_wbs_list(sheet_contract, '项目WBS号')
    print(wbs_in_contract_list)
    # 依据wbs 构造信息
    for i in range(len(wbs_in_contract_list)):
        wbs = wbs_in_contract_list[i]
        project_info_dir = importProInfo.get_project_info_by_wbs_info(sheet_contract, wbs)
        if importProInfo.is_exist(wbs) != True:
            # 插入数据
            sql, params = importProInfo.create_insert_sql_by_dir(project_info_dir)
            dbHepler.excute(sql, params)
            print("插入数据：")
            print(project_info_dir)
            print('\n')
        else:
            # update 更新部分数据
            sql, params = importProInfo.create_update_sql_by_dir(project_info_dir)
            dbHepler.excute(sql, params)
            print("补充数据：")
            print(project_info_dir)
            print('\n')

    print('合同表，共修改' + str(len(wbs_in_contract_list)) + '条数据')
    print('************************************************************************************************')

    # 获取 基本信息表 对象
    sheet_basic = importProInfo.get_sheet_by_name('基本信息表')
    # 获取 "合同表"中的WBS集合
    wbs_in_basic_list = importProInfo.get_wbs_list(sheet_basic, '项目WBS号')
    print(wbs_in_basic_list)
    # 依据wbs 构造信息
    for i in range(len(wbs_in_basic_list)):
        wbs = wbs_in_basic_list[i]
        project_info_dir = importProInfo.get_project_info_by_wbs_info(sheet_basic, wbs)
        if importProInfo.is_exist(wbs) != True:
            # 插入数据
            sql, params = importProInfo.create_insert_sql_by_dir(project_info_dir)
            # dbHepler.excute(sql, params)
            print("插入数据：")
            print(project_info_dir)
            print('\n')
        else:
            # update 更新部分数据
            sql, params = importProInfo.create_update_sql_by_dir(project_info_dir)
            # dbHepler.excute(sql, params)
            print("补充数据：")
            print(project_info_dir)
            print('\n')

    print('基本信息表，共修改' + str(len(wbs_in_basic_list)) + '条数据')
    print('************************************************************************************************')





main()