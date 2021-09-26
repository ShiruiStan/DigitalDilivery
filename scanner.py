# -- coding: utf-8 --
# @Time : 2021/9/23 16:10
# @Author : Shi Rui

import sqlite3
import xml.dom.minidom
from openpyxl import Workbook
import json
from tkinter import filedialog
import os


class Scanner:
    def __init__(self, db, schema, code):
        self.schema = {}
        self.analyse_schema(schema)
        self.type_code = {}
        self.material_code = {}
        self.analyse_code(code)
        self.components = {}
        self.component_tree = []
        self.analyse_components(db)

    def start(self):
        result_path = filedialog.askdirectory()
        if os.path.exists(result_path):
            # 写入 xls 文件，包括构件和构件树
            wb = Workbook()
            ws = wb.active
            ws.append(['Id', 'Name', 'Pid', 'NodeType', 'DisplayLabel', 'TypeCode', 'Guid', 'ElementId', 'Property', 'JD', 'WD', 'GD'])
            for i in self.component_tree:
                row = [i['id'], i['name'], i['pid'], i['type']]
                if i['guid'] in self.components.keys():
                    self.components[i['guid']]['property'].sort(key=lambda x: x['label'])
                    component = self.components[i['guid']]
                    row.append(component['displayLabel'])
                    row.append(component['typeCode'])
                    row.append(component['guid'])
                    row.append(component['elementId'])
                    row.append(json.dumps(component['property'], ensure_ascii=False))
                ws.append(row)
            wb.save(result_path + '/components.xlsx')
            # 导出json文件，只包含每个构件不含树结构
            with open(result_path + '/components.json', 'w', encoding='utf-8') as json_file:
                json.dump(list(self.components.values()), json_file, ensure_ascii=False)

    # 解析 ecSchema文件， 解析结果为schema——class——property——handler
    # handler 分为普通、材料和列表，其中嵌套字段struct在property层处理，归为普通handler
    # 对于每个属性，即数据库表中的某个字段，
    def analyse_schema(self, schema):
        for file in os.listdir(schema):
            file_type = file.split('.')
            if file_type[-1] != 'xml' and file_type[-2] != 'ecschema':
                continue
            else:
                ec_schema = xml.dom.minidom.parse(schema + '/' + file).documentElement
                ec_schema_name = ec_schema.getAttribute('schemaName')
                self.schema[ec_schema_name] = {}
                struct = {}
                for ec_class in ec_schema.getElementsByTagName('ECClass'):
                    if ec_class.getAttribute('isStruct') == 'True':
                        struct[ec_class.getAttribute('typeName')] = [prop.getAttribute('propertyName') for prop in
                                                                     ec_class.getElementsByTagName('ECProperty')]
                    ec_class_name = ec_class.getAttribute('typeName')
                    self.schema[ec_schema_name][ec_class_name] = {}
                    for ec_property in ec_class.getElementsByTagName('ECProperty'):
                        ec_property_name = ec_property.getAttribute('propertyName')
                        if ec_property_name == 'atcdi_materials':
                            self.schema[ec_schema_name][ec_class_name][ec_property_name] = 'MaterialHandler'
                        else:
                            self.schema[ec_schema_name][ec_class_name][ec_property_name] = 'NormalHandler'
                    for ec_struct_property in ec_class.getElementsByTagName('ECStructProperty'):
                        # 暂时弃用struct结构体
                        # 目前仅有养护schema文件中还存在结构体
                        pass
                    for ec_array_property in ec_class.getElementsByTagName('ECArrayProperty'):
                        # TODO 地层管桩中需要此属性
                        pass

    # 解析 code 的sqlite文件， 获取每个构建编码对应的名称， 获取每个材料编码对应的名称和单位
    def analyse_code(self, code):
        con = sqlite3.connect(code)
        cur = con.cursor()
        cur.execute('SELECT code,name_cn,name_en FROM table_elementTypeCode')
        for type_row in cur.fetchall():
            self.type_code[type_row[0]] = {'name_cn': type_row[1], 'name_en': type_row[2]}
        cur.execute('SELECT code,name,unit FROM table_materials')
        for material_row in cur.fetchall():
            self.material_code[material_row[0]] = {'name': material_row[1], 'unit': material_row[2]}
        cur.close()
        con.close()

    # 1-通过GuidElementTable获取所有构件id
    # 2-获取所有表名，对每个表找到其对应的schema和class，查看其中有哪些属性和对应的handler类型
    # 3-通过handler处理后将属性填充至 property 字段
    def analyse_components(self, db):
        con = sqlite3.connect(db)
        cur = con.cursor()
        cur.execute('SELECT * FROM GuidElementTable')
        # 初次遍历获取项目所有guid
        for i in cur.fetchall():
            # 每个构件
            self.components[i[0]] = {
                'guid': i[0],
                'elementId': i[1],
                'typeCode': '',
                'displayLabel': i[2],
                'property': []
            }
        cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
        table_list = cur.fetchall()
        # 获取所有属性相关的表，将每个表中对应的属性组装到对应的构件里去
        for i in table_list:
            table = i[0]
            if table in ['UnitTable', 'GuidElementTable', 'Tree_IDPID_temp']:
                continue
            cur.execute("SELECT * FROM UnitTable WHERE PropertyName = '%s'" % table)
            class_info = table.split('_____')  # 表名拆分后包括schema和class
            class_label = cur.fetchall()[0][2]
            property_info = self.schema[class_info[0]][class_info[1]]
            cur.execute('SELECT ATCDI_guid FROM ' + table)
            els = cur.fetchall()
            # 获取指定表中所有行的guid
            for j in els:
                guid = j[0]
                # 第一层property， 一般可能00_基本属性、01_结构属性、02_材料及零件属性
                class_value = {
                    'field': class_info[1],
                    'label': class_label,
                    'value': [],
                    'unit': ''
                }
                # 针对该guid，从self.schema中获取所有字段和对应的handler
                for field in property_info:
                    cur.execute("SELECT * FROM UnitTable WHERE PropertyName = '%s'" % (table + '_____' + field))
                    field_info = cur.fetchall()[0]
                    # 第二层property， 一般是各字段属性、值和单位
                    # 材料属性去除了无用的层级
                    # 数组属性会继续嵌套一层
                    field_value = {
                        'field': field,
                        'label': field_info[2],
                        'value': [],  # 列表对应数组，一般情况直接为值
                        'unit': field_info[1]
                    }
                    cur.execute("SELECT %s FROM %s WHERE ATCDI_GUID = '%s'" % (field, table, guid))
                    val = cur.fetchall()[0][0]
                    if property_info[field] == 'NormalHandler':
                        if field == "Typecode" and class_info[1] == "ATCDI_BaseProperty":
                            # TODO 此处未考虑一个构件存在多个typeCode的可能性，后期可能要修正
                            self.components[guid]['typeCode'] = val
                        field_value['value'] = val
                        class_value['value'].append(field_value)
                    elif property_info[field] == 'MaterialHandler':
                        field_value = self.handle_material(val)
                        if len(field_value) != 0:
                            class_value['value'] = field_value
                        else:
                            class_value = None
                    elif property_info[field] == 'ArrayHandler':
                        # TODO 数组处理
                        pass
                if class_value is not None:
                    self.components[guid]['property'].append(class_value)
        # 获取构件树
        cur.execute("SELECT * FROM Tree_IDPID_temp")
        for node in cur.fetchall():
            self.component_tree.append({
                'id': node[3],
                'pid': node[4],
                'name': node[5],
                'guid': node[0],
                'type': node[7] - 1,
            })
        cur.close()
        con.close()

    # 处理材料属性字段，直接返回处理完的list
    def handle_material(self, mat_str):
        mat_info = []
        if mat_str != '':
            for mat in json.loads(mat_str):
                mat_info.append({
                    'field': mat['Code'],
                    'label': self.material_code[mat['Code']]['name'],
                    'value': mat['Value'],
                    'unit': self.material_code[mat['Code']]['unit']
                })
        return mat_info
