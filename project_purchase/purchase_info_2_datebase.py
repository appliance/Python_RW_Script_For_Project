from dbHelper import DBHelper
import  xlrd

class ImportPurchaseInfo:
    def __init__(self, excel_url):
        self.excel_url = excel_url
        self.excel_property_name_2_sql_dir = {
            '采购订单': 'purchase_order', '合同名称': 'contract_name ',
            '供应商名称': 'supplier ', '合同类型': 'contract_type', '付款条件': 'payment_terms',
            '付款率': 'payment_rate', '合同总额': 'contract_money', '项目回款比例': 'refunds_rate',
            '已收票': 'receive_money', '背靠背可付金额': 'backtoback_money', '实际已付款': 'actual_payment',
            '未付款': 'unpaid_money', '背靠背欠款金额': 'backtoback_debt', '确认成本金额': 'cost_confirmed',
            '入成本比例': 'cost_rate', '未入成本金额': 'not_cost_monet'

        }
        self.excel_property_type_dir = {
            '采购订单': 'str', '合同名称': 'str',
            '供应商名称': 'str', '合同类型': 'str', '付款条件': 'str',
            '付款率': 'float', '合同总额': 'float', '项目回款比例': 'float',
            '已收票': 'float', '背靠背可付金额': 'float', '实际已付款': 'float',
            '未付款': 'float', '背靠背欠款金额': 'float', '确认成本金额': 'float',
            '入成本比例': 'float', '未入成本金额': 'float'
        }


    """
        获取关联项目id
        :parameter wbs wbs号
        :return None 或 pid
    """
    def get_project_id_by_wbs(self, wbs):
        dbHelper =DBHelper()
        dbHelper.connect_database()
        sql = 'select id from project where WBS=%s;'
        params = wbs
        result = dbHelper.select(sql, params)
        dbHelper.close_database()
        if result == ():
            return None
        else:
            return result[0][0]

    """
        获取excel数据表对象
        :parameter excel_url excel总表地址
        :parameter sheet_name 总表中子表名称
        :return sheet 子表处理对象
    """
    def get_sheet_by_name(self, sheet_name):
        book = xlrd.open_workbook(self.excel_url)
        sheets = book.sheet_names()
        sheet_name_2_index_dir = dict(zip(sheets, [i for i in range(len(sheets))]))
        if sheet_name in sheet_name_2_index_dir.keys():
            index = sheet_name_2_index_dir[sheet_name]
            sheet = book.sheet_by_index(index)
            return sheet
        else:
            print('警告: 该表中不存在' + sheet_name + '表，请核实表信息！！！')
            return

    """
        按行构建交易记录信息字典
        :parameter row  某一行
        :parameter sheet 表对象
        :return purchase_info_dir
    """
    def get_purchase_info_dir_by_row(self, row, sheet):
        purchase_info_dir = {}
        for col_index in range(sheet.ncols):
            # 注: 此处的0 根据表结构而来 默认表属性放置在第一行
            if str(sheet.cell(0, col_index).value).replace('\n', '') in self.excel_property_name_2_sql_dir.keys():
                property_type = self.excel_property_type_dir[sheet.cell(0, col_index).value]
                if property_type == 'str':
                    purchase_info_dir[self.excel_property_name_2_sql_dir[sheet.cell(0, col_index).value]] = str(sheet.cell(row, col_index).value)
                    # 订单号单独处理
                    if sheet.cell(0, col_index).value == '采购订单':
                        if sheet.cell(row, col_index).value != '':
                            purchase_info_dir[self.excel_property_name_2_sql_dir[sheet.cell(0, col_index).value]] = str(int(sheet.cell(row, col_index).value)).replace('\n', '')
                        else:
                            purchase_info_dir[self.excel_property_name_2_sql_dir[sheet.cell(0, col_index).value]] = None
                elif property_type == 'float':
                    if sheet.cell(row, col_index).value == '':
                        purchase_info_dir[self.excel_property_name_2_sql_dir[sheet.cell(0, col_index).value]] = None;
                    else:
                        purchase_info_dir[self.excel_property_name_2_sql_dir[sheet.cell(0, col_index).value]] = round(sheet.cell(row, col_index).value, 12)
        return purchase_info_dir

    """
        依据dir,构建sql查询语句
        :parameter purchase_info_dir 提取的购买信息字典
        :return sql（查询语句）  params（值） 
    """
    def create_sql_by_dir(self, purchase_info_dir):
        # 构造insert语句
        sql = 'insert into purchase('
        for i in range(len(purchase_info_dir.keys())):
            sql += list(purchase_info_dir.keys())[i]
            if i == len(purchase_info_dir.keys()) - 1:
                break
            else:
                sql += ','
        sql += ')values('
        for i in range(len(purchase_info_dir.keys())):
            sql += '%s'
            if i == len(purchase_info_dir.keys()) - 1:
                break
            else:
                sql += ','
        sql += ');'

        # 构造parmas
        params = []
        for key in purchase_info_dir.keys():
            params.append(purchase_info_dir.get(key))
        return sql, list(params)

