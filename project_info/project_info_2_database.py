import xlrd
import re
from datetime import date
from project_info.dbHelper import DBHelper
from project_info.db_config import db_config


class ImportProjectInfo:
    def __init__(self):
        self.excel_url = db_config['excel_url']
        # excel表属性映射mysql字段字典
        self.excel_property_name_2_sql_dir = {
            '项目WBS号': 'WBS', '销售订单': 'order_number', '合同编号': 'contract_number', '签约方': 'contract_party',
            '转签方式': 'transfer_type', '合同名称': 'contract_name', '合同总金额': 'contract_money',
            '已开票比例': 'invoiced_rate', '已投运比例': 'use_rate', '已回款比例': 'pay_received_rate',
            '付款方式': 'pay_type', '合同签约时间': 'contract_sign_date', '质保到期日': 'warranty_expiration_date',
            '合同收入确认总金额': 'contract_confirmed_money',
            '项目名称': 'project_name', '立项时间': 'project_approval_time', '办公地点': 'location',
            '项目类型': 'project_type', '项目类别': 'project_category', '合同性质': 'contract_nature',
            '所属产品线': 'product_line', '售前经理': 'presales_manager', '项目经理': 'project_manager',
            '项目简介': 'introduction', '备注': 'remark',
            '项目状态': 'project_status', '项目人数': 'people_number', '工期（月）': 'project_period',
            '计划启动完成日期': 'planned_launching_date', '实际启动完成日期': 'actual_launching_date',
            '计划编码开发完成日期': 'planned_coding_date', '实际编码开发完成日期': 'actual_coding_date',
            '计划第三方测试完成日期': 'planned_test_date', '实际第三方测试完成日期': 'actual_test_date',
            '计划实施入场日期': 'planned_implement_date', '实际实施入场日期': 'actual_implement_date',
            '计划上线试运行完成日期': 'planned_online_date', '实际上线试运行完成日期': 'actual_online_date',
            '计划验收完成时间': 'planned_acceptance_date', '实际验收完成时间': 'actual_acceptance_date',
            '计划收入（万元）': 'planned_income', '计划毛利率': 'margin_rate', '计划人工成本（万元）': 'planned_labor_cost',
            '计划外包服务费（万元）': 'outsource_service_cost', '计划差旅费（万元）': 'planned_travel_cost',
            '计划软硬件费用（万元）': 'hardware_software_cost', '计划技术分包费（万元）': 'technical_subcontract_cost',
            '计划其他费用（万元）': 'planned_other_cost', '预算合计（万元）': 'total_budget', '技术关闭状态': 'shutdown_status'
        }
        self.excel_property_type_dir = {
            '项目WBS号': 'str', '销售订单': 'str', '合同编号': 'str', '签约方': 'str',
            '转签方式': 'str', '合同名称': 'str', '合同总金额': 'float',
            '已开票比例': 'str', '已投运比例': 'str', '已回款比例': 'str',
            '付款方式': 'str', '合同签约时间': 'date', '质保到期日': 'date',
            '合同收入确认总金额': 'float',
            '项目名称': 'str', '立项时间': 'date', '办公地点': 'str',
            '项目类型': 'str', '项目类别': 'str', '合同性质': 'str',
            '所属产品线': 'str', '售前经理': 'str', '项目经理': 'str',
            '项目简介': 'str', '备注': 'str',
            '项目状态': 'str', '项目人数': 'int', '工期（月）': 'int',
            '计划启动完成日期': 'date', '实际启动完成日期': 'date',
            '计划编码开发完成日期': 'date', '实际编码开发完成日期': 'date',
            '计划第三方测试完成日期': 'date', '实际第三方测试完成日期': 'date',
            '计划实施入场日期': 'date', '实际实施入场日期': 'date',
            '计划上线试运行完成日期': 'date', '实际上线试运行完成日期': 'date',
            '计划验收完成时间': 'date', '实际验收完成时间': 'date',
            '计划收入（万元）': 'float', '计划毛利率': 'str', '计划人工成本（万元）': 'float',
            '计划外包服务费（万元）': 'float', '计划差旅费（万元）': 'float',
            '计划软硬件费用（万元）': 'float', '计划技术分包费（万元）': 'float',
            '计划其他费用（万元）': 'float', '预算合计（万元）': 'float', '技术关闭状态': 'str'
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
        sheet_name_2_index_dir = dict(zip(sheets, [i for i in range(len(sheets))]))
        if sheet_name in sheet_name_2_index_dir.keys():
            index = sheet_name_2_index_dir[sheet_name]
            sheet = book.sheet_by_index(index)
            return sheet
        else:
            print('警告: 该表中不存在' + sheet_name + '表，请核实表信息！！！')
            return


    """
        获取excel子表中 wbs 集合
        :parameter sheet 表对象
        :parameter property_wbs_name excel 表中wbs属性名称（WBS元素|项目WBS号）
        :return  wbs_list
    """
    def get_wbs_list(self, sheet, property_wbs_name):
        wbs_list = []
        wbs_row_position = ''
        wbs_col_position = ''
        for row_index in range(sheet.nrows):
            for col_index in range(sheet.ncols):
                if sheet.cell(row_index, col_index).value == property_wbs_name:
                    wbs_row_position = row_index
                    wbs_col_position = col_index
                    break
        if wbs_col_position == '':
            print('该表中' + property_wbs_name + '不存在！')
            return wbs_list

        for row_index in range(wbs_row_position + 1, sheet.nrows):
            wbs_list.append(str(sheet.cell(row_index, wbs_col_position).value))

        return wbs_list



    """
        依据wbs 项目综合信息表，单表查询获取项目信息
        :parameter sheet_info 项目综合信息表对象
        :parameter wbs wbs号
        :return dir
    """
    def get_project_info_by_wbs_info(self, sheet_info, wbs):
        # 在项目执行表中找到此wbs记录
        row_position = ''
        col_position = ''
        project_info_dir = {}

        for row_index in range(sheet_info.nrows):
            for col_index in range(sheet_info.ncols):
                if sheet_info.cell(row_index, col_index).value == wbs:
                    row_position = row_index
                    col_position = col_index
        # 构造数据
        for col_index in range(col_position, sheet_info.ncols):
            # 注： 此处1是根据excel项目综合信息表确定的， 若表结构变化则出错
            property_name = sheet_info.cell(0, col_index).value.strip().replace('\n', '')
            if property_name in self.excel_property_name_2_sql_dir.keys():
                property_set_in_database = self.excel_property_name_2_sql_dir.get(property_name)
                # 判断是否为date数据类型,需要进行格式转化
                if self.excel_property_type_dir.get(property_name) == 'date':
                    # date 分有数值和空数值两种情况处理
                    if sheet_info.cell(row_position, col_index).value != '':
                        if sheet_info.cell(row_position, col_index).value == '/':
                            project_info_dir[property_set_in_database] = date(1970, 1, 1)
                        else:
                            date_cell = xlrd.xldate_as_tuple(sheet_info.cell(row_position, col_index).value, 0)
                            project_info_dir[property_set_in_database] = date(*date_cell[0:3]).strftime('%Y-%m-%d')
                    else:
                        project_info_dir[property_set_in_database] = None
                elif self.excel_property_type_dir.get(property_name) == 'int':
                    # int 分空值数值， 非数值
                    if self.is_number(sheet_info.cell(row_position, col_index).value):
                        project_info_dir[property_set_in_database] = int(sheet_info.cell(row_position, col_index).value)
                    else:
                        project_info_dir[property_set_in_database] = None
                    # str 类型直接处理
                elif self.excel_property_type_dir.get(property_name) == 'str':
                    project_info_dir[property_set_in_database] = str(sheet_info.cell(row_position, col_index).value)
                elif self.excel_property_type_dir.get(property_name) == 'float':
                    project_info_dir[property_set_in_database] = float(sheet_info.cell(row_position, col_index).value)

        return project_info_dir



    """
        依据dir,构建sql查询语句
        :parameter project_info_dir 提取的项目信息字典
        :return sql（查询语句）  params（值） 
    """
    def create_insert_sql_by_dir(self, project_info_dir):
        # 构造insert语句
        sql = 'insert into project('
        for i in range(len(project_info_dir.keys())):
            sql += list(project_info_dir.keys())[i]
            if i == len(project_info_dir.keys())-1:
                break
            else:
                sql += ','
        sql += ')values('
        for i in range(len(project_info_dir.keys())):
            sql += '%s'
            if i == len(project_info_dir.keys())-1:
                break
            else:
                sql += ','
        sql += ');'

        # 构造parmas
        params = []
        for key in project_info_dir.keys():
            params.append(project_info_dir.get(key))
        return sql, list(params)

    """
        依据dir,构建sql update语句
        :parameter project_info_dir 提取的项目信息字典
        :return sql（查询语句）  params（值） 
    """
    def create_update_sql_by_dir(self, project_info_dir):
        params = []
        sql = 'update project set '
        for i in range(len(project_info_dir.keys())):
            if list(project_info_dir.keys())[i] == 'WBS':
                continue
            params.append(project_info_dir.get(list(project_info_dir.keys())[i]))
            sql += (list(project_info_dir.keys())[i] + ' = %s')
            if i == len(project_info_dir.keys()) - 1:
                break
            sql += ','
        sql += ' where WBS = %s;'
        params.append(project_info_dir['WBS'])

        return sql, params

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
        依据 wbs号 检索记录是否已插入数据库
        :param wbs
        :return True/False
    '''
    def is_exist(self, wbs):
        dbHelper = DBHelper()
        dbHelper.connect_database()
        sql = 'select count(*) from project where WBS = %s'
        param = [wbs]
        res = dbHelper.select(sql, param)
        dbHelper.close_database()
        if res[0][0] == 0:
            return False
        else:
            return True


