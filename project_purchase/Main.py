from purchase_info_2_datebase import ImportPurchaseInfo
from dbHelper import DBHelper
from db_config import db_config
import tkinter


def main():
    global waiting
    total_insert_record_count = 1
    # 连接数据库
    dbHelper = DBHelper()
    dbHelper.connect_database()

    # 获取 采购清单表对象
    excel_url = excel_url_input.get()
    importPurchaseInfo = ImportPurchaseInfo(excel_url=excel_url)

    # ‘采购清单’表对象
    sheet_purchase = importPurchaseInfo.get_sheet_by_name('采购清单')


    # 按行获 构建 记录dir
    # 注： 默认第一行为属性名
    for row_index in range(1, sheet_purchase.nrows):
        # 获取该行项目对应pid
        wbs = sheet_purchase.cell(row_index, 0).value
        if wbs == '':
            pid = None
        else:
            pid = importPurchaseInfo.get_project_id_by_wbs(wbs)
        # 获取该行项目信息dir
        purchase_info_dir = importPurchaseInfo.get_purchase_info_dir_by_row(row_index, sheet_purchase)
        purchase_info_dir['pid'] = pid

        # 将dir转化为 sql params 插入数据库
        sql, params = importPurchaseInfo.create_sql_by_dir(purchase_info_dir)
        print(sql)
        print(params)
        # 插入数据库
        dbHelper.excute(sql, params)
        total_insert_record_count += 1
        waiting = "导入完成，一共导入" + str(total_insert_record_count - 1) + '条项目数据。'

    dbHelper.close_database()







#刷新函数
def refreshText():
    msg.delete(0.0, tkinter.END)
    msg.insert("insert", waiting)
    msg.update()
    top.after(50, refreshText)



waiting = 'wait a minute....\n'
top = tkinter.Tk()
#设置窗口的大小宽x高+偏移量
top.geometry('300x500+500+200')
#设置窗口标题
top.title('Excel To Mysql')
#url
w = tkinter.Label(top, text="请输入文件名：").pack()
excel_url_input = tkinter.StringVar()
excel_url_input.set(db_config['excel_url'])
entry = tkinter.Entry(top, textvariable=excel_url_input).pack()
btn = tkinter.Button(top, text='开始录入', command=main).pack()
msg = tkinter.Text(top)
msg.pack()
top.after(50, refreshText)
top.mainloop()