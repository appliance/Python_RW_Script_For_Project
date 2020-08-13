from project_purchase.purchase_info_2_datebase import  ImportPurchaseInfo
from project_purchase.dbHelper import DBHelper
from project_purchase.db_config import db_config

def main():
    total_insert_record_count = 0
    # 连接数据库
    dbHelper = DBHelper()
    dbHelper.connect_database()

    # 获取 采购清单表对象
    importPurchaseInfo = ImportPurchaseInfo()

    # 采购表名称 默认为 ‘采购表(非人力)’  注意表明的改动
    sheet_purchase = importPurchaseInfo.get_sheet_by_name(db_config['purchase_sheet_name'])


    # 按行获 构建 记录dir
    # 注： 默认第一行为属性名
    for row_index in range(1, sheet_purchase.nrows):
        # 获取该行项目对应pid
        wbs = sheet_purchase.cell(row_index, 0).value
        # 获取该行项目信息dir
        purchase_info_dir = importPurchaseInfo.get_purchase_info_dir_by_row(row_index, sheet_purchase)

        # 将dir转化为 sql params 插入数据库
        sql, params = importPurchaseInfo.create_sql_by_dir(purchase_info_dir)

        # 判断记录是否存在
        if importPurchaseInfo.is_exist(wbs, purchase_info_dir['purchase_order']) == True:
            continue
        else:
            # 插入数据库
            dbHelper.excute(sql, params)
            total_insert_record_count += 1
            print(purchase_info_dir)


    print('表 ' + db_config['purchase_sheet_name'] +' 导入完成，一共导入' + str(total_insert_record_count) + '条项目数据。')

    dbHelper.close_database()


main()