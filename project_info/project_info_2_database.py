import xlrd
import re
from datetime import date
from dbHelper import DBHelper
from db_config import db_config


class ImportProjectInfo:
    def __init__(self,  excel_url):
        self.excel_url = excel_url
        # excel表属性映射mysql字段字典
        self.excel_property_name_2_sql_dir = {
            '项目WBS号': 'WBS',
            '采购订单': 'purchase_order', '采购\n性质': 'purchase_nature', '采购\n订单\n数量': 'purchase_order_number',
            '办公地点': 'location', '项目\n类型': 'project_type', '项目\n类别': 'project_category',
            '所属\n产品线': 'product_line', '售前\n经理': 'presales_manager', '项目人数': 'people_number',
            '工期\n(月)': 'project_period', '计划\n启动会\n完成日期': 'planned_launching_date',
            '实际\n启动会\n完成日期': 'actual_launching_date', '计划\n编码开发\n完成日期': 'planned_coding_date',
            '实际\n编码开发\n完成日期': 'actual_coding_date', '计划\n第三方测试\n完成日期': 'planned_test_date',
            '实际\n第三方测试\n完成日期': 'actual_test_date', '计划\n实施入场\n日期': 'planned_implement_date',
            '实际\n实施入场\n日期': 'actual_implement_date', '计划\n上线试运行\n完成日期': 'planned_online_date',
            '实际\n上线试运行\n完成日期': 'actual_online_date', '计划\n验收\n完成时间': 'planned_acceptance_date',
            '实际\n验收\n完成时间': 'actual_acceptance_date', '当前执行阶段简述': 'brief_description',
            '是否\n验收': 'acceptance_status', '技术关闭': 'shutdown_status'
        }
        self.excel_property_type_dir = {
            '项目WBS号': 'str',
            '采购订单': 'str', '采购\n性质': 'str', '采购\n订单\n数量': 'int',
            '办公地点': 'str', '项目\n类型': 'str', '项目\n类别': 'str',
            '所属\n产品线': 'str', '售前\n经理': 'str', '项目人数': 'int',
            '工期\n(月)': 'int', '计划\n启动会\n完成日期': 'date',
            '实际\n启动会\n完成日期': 'date', '计划\n编码开发\n完成日期': 'date',
            '实际\n编码开发\n完成日期': 'date', '计划\n第三方测试\n完成日期': 'date',
            '实际\n第三方测试\n完成日期': 'date', '计划\n实施入场\n日期': 'date',
            '实际\n实施入场\n日期': 'date', '计划\n上线试运行\n完成日期': 'date',
            '实际\n上线试运行\n完成日期': 'date', '计划\n验收\n完成时间': 'date',
            '实际\n验收\n完成时间': 'date', '当前执行阶段简述': 'str',
            '是否\n验收': 'str', '技术关闭': 'str'
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
            if sheet_info.cell(1, col_index).value in self.excel_property_name_2_sql_dir.keys():
                property_set_in_database = self.excel_property_name_2_sql_dir.get(sheet_info.cell(1, col_index).value)
                # 判断是否为date数据类型,需要进行格式转化
                if self.excel_property_type_dir.get(sheet_info.cell(1, col_index).value) == 'date':
                    # date 分有数值和空数值两种情况处理
                    if sheet_info.cell(row_position, col_index).value != '':
                        if sheet_info.cell(row_position, col_index).value == '/':
                            project_info_dir[property_set_in_database] = None
                        else:
                            date_cell = xlrd.xldate_as_tuple(sheet_info.cell(row_position, col_index).value, 0)
                            project_info_dir[property_set_in_database] = date(*date_cell[0:3]).strftime('%Y-%m-%d')
                    else:
                        project_info_dir[property_set_in_database] = None
                elif self.excel_property_type_dir.get(sheet_info.cell(1, col_index).value) == 'int':
                    # int 分空值数值， 非数值
                    if self.is_number(sheet_info.cell(row_position, col_index).value):
                        project_info_dir[property_set_in_database] = int(sheet_info.cell(row_position, col_index).value)
                    else:
                        project_info_dir[property_set_in_database] = None
                    # str 类型直接处理
                elif self.excel_property_type_dir.get(sheet_info.cell(1, col_index).value) == 'str':
                        project_info_dir[property_set_in_database] = str(sheet_info.cell(row_position, col_index).value)

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
    break


