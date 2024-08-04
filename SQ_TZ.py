# -*- coding:gbk -*-
"""
充分就业社区台账
"""
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.exceptions import InvalidFileException
from datetime import datetime
from Reset_width import Reset
from pprint import pprint
import os, re


class JCTZ:

    def __init__(self):
        self.header = []
        self.value_lis = []
        self.sy_values = []
        self.border_style = Border(left=Side(style='thin'), 
                            right=Side(style='thin'), 
                            top=Side(style='thin'), 
                            bottom=Side(style='thin'))  # 设置单元格的边框样式，全部为细线
        self.alignment_style = Alignment(horizontal='center', vertical='center') # 设置单元格的对齐方式，居中对齐
        self.font = Font(name='宋体', size=10)
        self.path = os.path.join(os.path.dirname(__file__), 'template_excel')
        self.out_path = os.path.join(os.path.dirname(__file__), 'template_excel_bak')
        # self.file_tag = ''
        self.jntc = ''
        self.run_smz_status = False
        self.Rest = Reset()

    def insert_row(self, ws, data, row_index):
        ws.insert_rows(row_index)
        for col_index, value in enumerate(data, start=1):
            cell = ws.cell(row=row_index, column=col_index)
            if isinstance(value, datetime):
                cell.number_format = 'yyyy/mm/dd'
            cell.border = self.border_style
            cell.alignment = self.alignment_style
            cell.font = self.font
            cell.value = value
            
    def get_headers(self, ws, min_num, max_num):
        # 获取excel表头
        if min_num != max_num:
            # 复合表头的情况
            mer_row = [mer_cell.coord for mer_cell in ws.merged_cells.ranges if (min_num in (mer_cell.min_row, mer_cell.max_row) or max_num in (mer_cell.min_row, mer_cell.max_row))]
            # 拿到复合表头中含有合并单元格的区域
            rows_list = []
            # 对多种列名(A2,AA2,ABC2)处理
            for long in range(5, 20, 2):
                rows = [i for i in mer_row if len(i)==int(long)]
                if not rows:
                    break
                # 排序，保证表头顺序，方便直接写入一行数据
                rows.sort(key=lambda x: x.split(':')[0])
                rows_list.extend(rows)
            header = []
            for i in rows_list:
                numbers = re.findall('[0-9]+', i)
                if numbers[0] == numbers[1]:
                    # 单元格同行合并
                    start, end = re.findall('[A-Z]+', i)
                    # 列名从字母转换为数字
                    col_num1 = column_index_from_string(start[0])
                    col_num2 = column_index_from_string(end[0])
                    for n in range(col_num1, col_num2 + 1):
                        coord = get_column_letter(n) + str(int(i[1]) + 1)
                        if coord in [i[:2] for i in mer_row]:
                            continue
                        header.append(ws[coord].value.replace('\n', '').replace(' ', '').replace('（必填）', ''))

                else:
                    value = ws[i.split(':')[0]].value
                    if value:
                        header.append(value.replace('\n', '').replace(' ', '').replace('（必填）', ''))
            h2 = [i.value for i in ws[max_num]]
            if len(header) != len(h2):
                l = h2[len(header) - len(h2):]
                if None not in l:
                    header.extend(l)
        else:
            # 单行表头直接取
            header = [i.value.replace('\n', '').replace(' ', '').replace('（必填）', '') for i in ws[max_num] if i.value]
        # print(header)
        return header
    
    def syry(self, ws_sy):
        header_sy = self.get_headers(ws_sy, 2, 2)
        syry_values = []
        for row in [i for i in ws_sy.iter_rows()][2:]:
            if not row[0].value:
                break
            lis_sy = []
            for i in row:
                lis_sy.append(i.value)
            syry_values.append(dict(zip(header_sy, lis_sy)))
        keys = ['序号', '姓名', '性别', '年龄', '文化程度', '身份证号', '户籍性质', '技能特长', '就业创业证号', '是否登记失业人员', '失业人员类型', '失业时间', '领取失业保险金起止时间', '求职意向', '培训意向', '就业服务需求', '联系电话', '类型', '等级']
        for i in syry_values:
            values = []
            for key in keys:
                v = i.get(key)
                if key == '序号':
                    v = i.get('总序号')
                if key == '年龄':
                    v = datetime.now().year - int(i.get('身份证号')[6:10])
                if key == '文化程度':
                    v = i.get('学历')
                if key == '户籍性质':
                    v = '城镇'
                if key == '技能特长':
                    v = i.get('特殊技能')
                if key == '是否登记失业人员':
                    v = '是'
                if key == '失业人员类型':
                    v = '(9)'
                if key == '领取失业保险金起止时间':
                    v = '无'
                if key == '求职意向':
                    v = '服务'
                if key == '培训意向':
                    v = '无'
                if key == '就业服务需求':
                    v = '（1）'
                if key == '联系电话':
                    v = i.get('电话')
                if key == '类型':
                    v = '中短期'
                if key == '等级':
                    v = '缺乏技能' if i.get('特殊技能') == '无' else '具备技能'
                if key == '失业时间':
                    v_time = i.get('失业时间')
                    if v_time:
                        month = datetime.now().month - v.month
                        v = v_time.replace(month=datetime.now().month - 2) if month > 2 else v_time
                    else:
                        v = datetime.now().replace(month=datetime.now().month - 2, day=1, hour=0, minute=0, second=0, microsecond=0)
                values.append(v)
            self.sy_values.append(values)

    def read_file(self, file, min_num, max_num):
        # 从excel文件读取内容，将表头和内容组成字典保存在self.value_lis列表中
        # path = os.path.join(os.path.dirname(__file__), '充分就业社区基础台账', file)
        path = file
        # print(f'path:{path}')
        self.value_lis = []
        try:
            wb = load_workbook(path, data_only=True)
        except InvalidFileException as E:
            print(f'不能处理{path.split(".")[-1]}类型的Excel文件,推荐另存为xlsx类型')
            return 'error'
        else:
            if "实名制" in path.split('/')[-1] and '失业人员情况' in wb.sheetnames:
                self.syry(wb['失业人员情况'])
            try:
                ws = wb['新就业人员']
            except KeyError:
                ws = wb.worksheets[0]
            finally:
                header = self.get_headers(ws, min_num, max_num)
                for row in [i for i in ws.iter_rows()][max_num:]:
                    if not row[0].value:
                        break
                    lis = []
                    for i in row:
                        lis.append(i.value)
                    self.value_lis.append(dict(zip(header, lis)))
                # self.file_tag = file
                # pprint(self.value_lis)

    def write_excel(self, file, min_num, max_num):
        # 将self.value_lis列表中的内容根据相同表头写入到excel，其中的有些列特殊处理
        path = os.path.join(self.path, file)
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        mer_lis = [mer_cell.coord for mer_cell in ws.merged_cells.ranges if max_num < mer_cell.min_row or max_num < mer_cell.max_row]
        # print(f'mer_lis:{mer_lis}')
        num = max([cell.value if isinstance(cell.value, int) else 0 for cell in ws['A'][max_num:] if cell.value is not None], default=0)
        header = self.get_headers(ws, min_num, max_num)
        value_lis = []
        values_lis = self.value_lis
        if "3就业困难人员管理台账" in file:
            values_lis = [i for i in self.value_lis if i.get('就业困难人员') == '是']
        for n, i in enumerate(values_lis, start=1):
            # 根据key刚从关联表获得数据
            lis = []
            for k in header:
                # 对特征列的特定处理
                v = i.get(k)
                if k == '序号':
                    v = n + num
                
                # 台账12
                # 文化程度
                if '12求职人员登记台帐' in file:
                    if k == '所学专业':
                        v = '无'
                    if k == '求职地区':
                        v = '本地'
                    if k == '用工形式':
                        v = '长期'
                    if k == '期望薪资':
                        v = '面议'
                    if k == '是否应届高校毕业生' or k == '是否农村转移劳动者':
                        v = '否'
                    if k == '职业指导次数' or k == '职业介绍次数':
                        v = '1'
                    if k == '登记时间':
                        v = i.get('失业时间')
                # 台账15
                elif '15退休人员基本情况及相关信息台帐' in file:
                    if k == '健康状况':
                        v = '⑵'
                    if k == '特殊群体类别':
                        pass
                    if k == '原工作单位':
                        v = i.get('所在单位')
                    if k == '退休时间':
                        v = i.get('退休/伤亡时间')
                    if k == '与退休人员关系':
                        v = '本人'
                    if k == '退休人员联系电话':
                        v =  i.get('联系电话')
                    if k == '身份证号码':
                        v = i.get('身份证号')
                # 台账6
                elif '6新就业人员信息台账' in file:
                    if k == '地市名称（市、县（区））':
                        v = '鸡西市滴道区'
                    if k == '就业单位' and not v:
                        v = i.get('就业单位(灵活就业填具体工作内容）')
                    if k == '单位就业':
                        v = '是' if i.get('就业方式') == '单位就业' else '否'
                    if k == '个体工商户':
                        v = '是' if i.get('就业方式') == '个体工商户' else '否'
                    if k == '公益性岗位':
                        v = '是' if i.get('就业方式') == '公益性岗位安置' else '否'
                    if k == '灵活就业':
                        v = '是' if i.get('就业方式') == '灵活就业' else '否'
                    if k == '失业人员再就业':
                        # v = '是' if i.get('就业类型') == '失业再就业' else '否'
                        v = '是'
                    if k == '就业困难再就业':
                        v = '是' if i.get('就业困难人员') == '是' else '否'
                # 台账5
                elif '5失业人员再就业信息明细台账' in file:
                    if k == '文化程度' and not v:
                        v = i.get('学历')
                    if k == '户籍性质':
                        v = '城镇'
                    if k == '技能特长' and not v:
                        v = i.get('技能等级证书', '无')
                    if k == '失业人员类型':
                        v = '(9)'
                    if k == '失业时间':
                        d = i.get('登记就业时间（年/月/日）')
                        if not d:
                            d = i.get('就业时间')
                        if d:
                            v = d.replace(month=d.month-1)
                    if k == '再就业时间' and not v:
                        v = i.get('登记就业时间（年/月/日）')
                    if k == '就业渠道':
                        jyqd = {"灵活就业": "(5)", "单位就业": "(3)", "个体工商户": "(4)", "公益性岗位安置": "(6)"}
                        v = jyqd.get(i.get('就业方式'))
                    if k == '现就业单位':
                        v = str(i.get("就业单位(灵活就业填具体工作内容）", "")).split("/")[0]
                    if k == '就业岗位':
                        if i.get('就业方式') == '灵活就业':
                            v = str(i.get('就业单位(灵活就业填具体工作内容）', '')).split('/')[-1]
                        else:
                            v = '员工'
                    if k == '所属产业':
                        v = i.get('从事产业类型')
                # 台账4
                elif '4失业人员管理台账' in file:
                    if k == '序号':
                        v += len(self.sy_values)
                    if k == '年龄':
                        v = str(i.get('年龄', '')).split('-')[0]
                    if k == '是否登记失业人员':
                        v = '否'
                    if k == '技能特长':
                        self.jnct = v = i.get('技能等级证书')
                    if k == '失业时间':
                        d = i.get('登记就业时间（年/月/日）')
                        if d:
                            v = d.replace(month=d.month-1)
                    if k == '文化程度' and not v:
                        v = i.get('学历')
                    if k == '失业人员类型':
                        v = '(9)'
                    if k == '户籍性质':
                        v = '城镇'
                    if k == '领取失业保险金起止时间':
                        v = '无'
                    if k == '求职意向':
                        if i.get('就业方式') == '灵活就业':
                            v = str(i.get('就业单位(灵活就业填具体工作内容）', '')).split('/')[-1]
                        else:
                            v = '员工'
                    if k == '培训意向':
                        v = '无'
                    if k == '就业服务需求':
                        v = '（1）'
                    if k == '类型':
                        v = '中短期'
                    if k == '等级':
                        if '级' in self.jnct:
                            v = '具备技能'
                        else:
                            v = '缺乏技能'
                elif "3就业困难人员管理台账" in file:
                    if k == "文化程度":
                        v = i.get('学历')
                    if k == "技能特长":
                        v = i.get('技能等级证书')
                    if k == "就业困难人员类型":
                        v = '①'
                    if k == "家庭住址":
                        v = '鸡西市滴道区'
                    if k == "再就业时间":
                        v = i.get('登记就业时间（年/月/日）')
                    if k == '现就业单位':
                        v = str(i.get("就业单位(灵活就业填具体工作内容）", "")).split("/")[0]
                    if k == '就业岗位':
                        if i.get('就业方式') == '灵活就业':
                            v = str(i.get('就业单位(灵活就业填具体工作内容）', '')).split('/')[-1]
                        else:
                            v = '员工'
                    if k == '就业去向':
                        jyqx = {"灵活就业": "⑤", "单位就业": "③", "个体工商户": "④", "公益性岗位安置": "⑥"}
                        v = jyqx.get(i.get('就业方式'))
                    if k == '是否签定劳动合同':
                        v = '否'
                    if k == '合同期限':
                        v = '无'
                    if k == '是否就业援助对象':
                        v = ' 是'
                    if k == '就业援助形式':
                        v = '⑤'
                lis.append(v)
            value_lis.append(lis)
        if "4失业人员管理台账" in file:
            value_lis = self.sy_values + value_lis
        # pprint(value_lis)
        # 计算插入行号
        insert_num = num + max_num
        if mer_lis:
            self.write_before(ws, mer_lis)
        for i, value in enumerate(value_lis, 1):
            self.insert_row(ws, value, insert_num + i)
        if mer_lis:
            self.write_tail(ws, mer_lis, len(value_lis))
        out_path = os.path.join(self.out_path, file)
        reset_rows = [1, 4] if "12求职人员登记台帐" in file else [1, 4, 5]
        self.Rest.reset(ws, rows=reset_rows, value=True)
        wb.save(out_path)
        print(f'文件{out_path}，写入完成')

    def write_before(self, ws, mer_lis):
        for i in mer_lis:
            # num1, num2 = re.findall('[0-9]+', i)
            # 取消单元格合并
            ws.unmerge_cells(i)
        
    def write_tail(self, ws, mer_lis, rows):
        for i in mer_lis:
            num1, num2 = re.findall('[0-9]+', i)
            # 单元格合并
            start, end = re.findall('[A-Z]+', i)
            ws.merge_cells(f'{start}{int(num1) + rows}:{end}{int(num2) + rows}')

    def run_smz(self, smz_path, out_path):
        # 根据实名制信息，导出台账5、6
        res = self.read_file(smz_path, 2, 3)
        if res and res == 'error':
            return
        if out_path:
            self.out_path = out_path
        self.write_excel('3就业困难人员管理台账.xlsx', 4, 5)
        self.write_excel('4失业人员管理台账.xlsx', 4, 5)
        self.write_excel('5失业人员再就业信息明细台账.xlsx', 4, 5)
        self.write_excel('6新就业人员信息台账.xlsx', 4, 5)
        self.run_smz_status = True

    def run_4to12(self, tz4_path, out_path):
        # 根据台账4，导出台账12
        if os.path.isfile(tz4_path):
            res = self.read_file(tz4_path, 4, 5)
        else:
            res = self.read_file(os.path.join(tz4_path, '4失业人员管理台账.xlsx'), 4, 5)
        if res and res == 'error':
            return
        if out_path:
            self.out_path = out_path
        self.write_excel('12求职人员登记台帐.xlsx', 4, 4)
    
    def run_7to15(self, tz7_path, out_path):
        # 将台账7导入台账15
        res = self.read_file(tz7_path, 4, 4)
        if res and res == 'error':
            return
        if out_path:
            self.out_path = out_path
        self.write_excel('15退休人员基本情况及相关信息台帐.xlsx', 4, 5)

    def main(self):
        # 根据实名制信息，导出台账3、4、5、6
        self.run_smz()

        if self.run_smz_status:
            # 根据台账4，导出台账12
            self.run_4to12()

        # 将台账7导入台账15
        # self.run_7to15()


if __name__ == '__main__':
    jctz = JCTZ()
    # jctz.run_smz(smz_path='C:/Users/Administrator/Desktop/立井2024年4月实名制台账20240426版.xlsx', out_path='')
    # jctz.run_smz(smz_path='C:/Users/XBD/Desktop/实名制.xlsx', out_path='C:/Users/XBD/Desktop/台账结果')
    # jctz.run_4to12(tz4_path='C:/Users/XBD/Desktop/台账结果/4失业人员管理台账.xlsx', out_path='C:/Users/XBD/Desktop/台账结果')
