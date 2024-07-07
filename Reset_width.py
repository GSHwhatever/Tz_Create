# -*- coding:gbk -*-
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.cell import MergedCell
from icecream import ic
import datetime


class Reset:

    def __init__(self):
        self.border_style = Border(left=Side(style='thin'), 
                                    right=Side(style='thin'), 
                                    top=Side(style='thin'), 
                                    bottom=Side(style='thin'))  # 设置单元格的边框样式，全部为细线
        self.alignment_style = Alignment(horizontal='center', vertical='center') # 设置单元格的对齐方式，居中对齐
        self.font = Font(name='宋体', size=10)
        # excel格式设置

    def reset(self, ws, rows=None, value=False):
        mer_lis = [mer_cell.coord.split(':')[0] for mer_cell in ws.merged_cells.ranges]
        # 修改单元格样式
        for row in ws:
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                coordinate = f'{cell.column_letter}{cell.row}'
                if coordinate in mer_lis:
                    continue
                if rows and cell.row in rows:
                    continue
                if value:
                    if cell.value:
                        cell.border = self.border_style
                        cell.alignment = self.alignment_style
                        cell.font = self.font
                else:
                    cell.border = self.border_style
                    cell.alignment = self.alignment_style
                    cell.font = self.font

        lks = [] # 第一步：计算每列最大宽度，并存储在列表lks中。
        for i in range(1,ws.max_column+1): #每列循环
            lk = 1 #定义初始列宽，并在每个行循环完成后重置
            for j in range(1, ws.max_row + 1): #每行循环
                cell = ws.cell(row=j,column=i)
                if isinstance(cell, MergedCell):
                    continue
                coordinate = f'{cell.column_letter}{cell.row}'
                if coordinate in mer_lis:
                    continue
                if rows and cell.row in rows:
                    continue
                sz = cell.value #每个单元格内容
                if isinstance(sz,str): #
                    try:
                        lk1 = len(sz.encode('gbk')) #gbk解码一个中文两字节，utf-8一个中文三字节，gbk合适
                    except UnicodeEncodeError:
                        lk1 = len(sz.encode('utf-8')) - 1
                elif isinstance(sz, datetime.datetime):
                    lk1 = len(sz.strftime("%Y/%m/%d"))
                else:
                    lk1 = len(str(sz))
                if lk < lk1:
                    lk = lk1 #借助每行循环将最大值存入lk中
                # ic(lk)
            lks.append(lk) # 将每列最大宽度加入列表。（犯了一个错，用lks = lks.append(lk)报错，append会修改列表变量，返回值none，而none不能继续用append方法）
        
        # 第二步：设置列宽
        for i in range(1, ws.max_column +1):
            k = get_column_letter(i) #将数字转化为列名,26个字母以内也可以用[chr(i).upper() for i in range(97, 123)]，不用导入模块\
            ws.column_dimensions[k].width = lks[i-1] + 2 #设置列宽，一般加两个字节宽度，可以根据实际情况灵活调整


if __name__ == "__main__":
    R = Reset()
    path = r'C:\Users\GSH\Desktop\4失业人员管理台账 - 副本.xlsx'
    wb = load_workbook(path)
    ws = wb.active
    R.reset(ws)
    wb.save(path)
