# -*- coding:gbk -*-
"""
��־�ҵ����̨��
"""
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter, column_index_from_string
from datetime import datetime
from pprint import pprint
import os, re


class JCTZ:

    def __init__(self):
        self.header = []
        self.value_lis = []
        self.hy_code_dic = {
            "ũ���֡�������ҵ": "A0000",
            "�ɿ�ҵ": "B0000",
            "����ҵ": "C0000",
            "������������ȼ����ˮ�����͹�Ӧҵ": "D0000",
            "����ҵ": "E0000",
            "����������ҵ": "F0000",
            "��ͨ���䡢�ִ�������ҵ": "G0000",
            "ס�޺Ͳ���ҵ": "H0000",
            "��Ϣ���䡢�������Ϣ��������ҵ": "I0000",
            "����ҵ": "J0000",
            "���ز�ҵ": "K0000",
            "���޺��������ҵ": "L0000",
            "��ѧ�о��ͼ�������ҵ": "M0000",
            "ˮ���������͹�����ʩ����ҵ": "N0000",
            "��������������������ҵ": "O0000",
            "����": "P0000",
            "��������Ṥ��": "Q0000",
            "�Ļ�������������ҵ": "R0000",
            "����������ᱣ�Ϻ������֯": "S0000",
            "������֯": "T0000"
        }
        self.border_style = Border(left=Side(style='thin'), 
                            right=Side(style='thin'), 
                            top=Side(style='thin'), 
                            bottom=Side(style='thin'))  # ���õ�Ԫ��ı߿���ʽ��ȫ��Ϊϸ��
        self.alignment_style = Alignment(horizontal='center', vertical='center') # ���õ�Ԫ��Ķ��뷽ʽ�����ж���
        self.font = Font(name='����', size=10)
        self.path = os.path.join(os.path.dirname(__file__), 'template_excel')
        self.out_path = os.path.join(os.path.dirname(__file__), 'template_excel_bak')
        self.file_tag = ''
        self.jntc = ''

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
        # ��ȡexcel��ͷ
        if min_num != max_num:
            # ���ϱ�ͷ�����
            mer_row = [mer_cell.coord for mer_cell in ws.merged_cells.ranges if (min_num in (mer_cell.min_row, mer_cell.max_row) or max_num in (mer_cell.min_row, mer_cell.max_row))]
            # �õ����ϱ�ͷ�к��кϲ���Ԫ�������
            rows_list = []
            # �Զ�������(A2,AA2,ABC2)����
            for long in range(5, 20, 2):
                rows = [i for i in mer_row if len(i)==int(long)]
                if not rows:
                    break
                # ���򣬱�֤��ͷ˳�򣬷���ֱ��д��һ������
                rows.sort(key=lambda x: x.split(':')[0])
                rows_list.extend(rows)
            header = []
            for i in rows_list:
                numbers = re.findall('[0-9]+', i)
                if numbers[0] == numbers[1]:
                    # ��Ԫ��ͬ�кϲ�
                    start, end = re.findall('[A-Z]+', i)
                    # ��������ĸת��Ϊ����
                    col_num1 = column_index_from_string(start[0])
                    col_num2 = column_index_from_string(end[0])
                    for n in range(col_num1, col_num2 + 1):
                        coord = get_column_letter(n) + str(int(i[1]) + 1)
                        if coord in [i[:2] for i in mer_row]:
                            continue
                        header.append(ws[coord].value.replace('\n', '').replace(' ', '').replace('�����', ''))
                else:
                    value = ws[i.split(':')[0]].value
                    if value:
                        header.append(value.replace('\n', '').replace(' ', '').replace('�����', ''))
            h2 = [i.value for i in ws[max_num]]
            if len(header) != len(h2):
                l = h2[len(header) - len(h2):]
                if None not in l:
                    header.extend(l)
        else:
            # ���б�ͷֱ��ȡ
            header = [i.value.replace('\n', '').replace(' ', '').replace('�����', '') for i in ws[max_num] if i.value]
        print(header)
        return header

    def read_file(self, file, min_num, max_num):
        # ��excel�ļ���ȡ���ݣ�����ͷ����������ֵ䱣����self.value_lis�б���
        # path = os.path.join(os.path.dirname(__file__), '��־�ҵ��������̨��', file)
        path = file
        # print(f'path:{path}')
        self.value_lis = []
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        header = self.get_headers(ws, min_num, max_num)
        for row in [i for i in ws.iter_rows()][max_num:]:
            if not row[0].value:
                break
            lis = []
            for i in row:
                lis.append(i.value)
            self.value_lis.append(dict(zip(header, lis)))
        self.file_tag = file
        # pprint(self.value_lis)

    def write_excel(self, file, min_num, max_num):
        # ��self.value_lis�б��е����ݸ�����ͬ��ͷд�뵽excel�����е���Щ�����⴦��
        path = os.path.join(self.path, file)
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        mer_lis = [mer_cell.coord for mer_cell in ws.merged_cells.ranges if max_num < mer_cell.min_row or max_num < mer_cell.max_row]
        # print(f'mer_lis:{mer_lis}')
        num = max([cell.value if isinstance(cell.value, int) else 0 for cell in ws['A'][max_num:] if cell.value is not None], default=0)
        header = self.get_headers(ws, min_num, max_num)
        value_lis = []
        values_lis = self.value_lis
        if "3��ҵ������Ա����̨��" in file:
            values_lis = [i for i in self.value_lis if i.get('��ҵ�����پ�ҵ') == '��']
        for n, i in enumerate(values_lis, start=1):
            # ����key�մӹ�����������
            lis = []
            for k in header:
                # �������е��ض�����
                v = i.get(k)
                if k == '���':
                    v = n + num
                
                # ̨��12
                # �Ļ��̶�
                if '12��ְ��Ա�Ǽ�̨��' in file:
                    if k == '��ѧרҵ':
                        v = '��'
                    if k == '��ְ����':
                        v = '����'
                    if k == '�ù���ʽ':
                        v = '����'
                    if k == '����н��':
                        v = '����'
                    if k == '�Ƿ�Ӧ���У��ҵ��' or k == '�Ƿ�ũ��ת���Ͷ���':
                        v = '��'
                    if k == 'ְҵָ������' or k == 'ְҵ���ܴ���':
                        v = '1'
                    if k == '�Ǽ�ʱ��':
                        v = i.get('ʧҵʱ��')
                # ̨��15
                elif '15������Ա��������������Ϣ̨��' in file:
                    if k == '����״��':
                        v = '��'
                    if k == '����Ⱥ�����':
                        pass
                    if k == 'ԭ������λ':
                        v = i.get('���ڵ�λ')
                    if k == '����ʱ��':
                        v = i.get('����/����ʱ��')
                    if k == '��������Ա��ϵ':
                        v = '����'
                    if k == '������Ա��ϵ�绰':
                        v =  i.get('��ϵ�绰')
                    if k == '���֤����':
                        v = i.get('���֤��')
                # ̨��6
                elif '6�¾�ҵ��Ա��Ϣ̨��' in file:
                    if k == '�������ƣ��С��أ�������':
                        v = '�����еε���'
                    if k == '��ҵ��λ' and not v:
                        v = i.get('��ҵ��λ(����ҵ����幤�����ݣ�')
                    if k == '��λ��ҵ':
                        v = '��' if i.get('��ҵ��ʽ') == '��λ��ҵ' else '��'
                    if k == '���幤�̻�':
                        v = '��'
                    if k == '�����Ը�λ':
                        v = '��'
                    if k == '����ҵ':
                        v = '��' if i.get('��ҵ��ʽ') == '����ҵ' else '��'
                    if k == 'ʧҵ��Ա�پ�ҵ':
                        v = '��' if i.get('��ҵ����') == 'ʧҵ�پ�ҵ' else '��'
                    if k == '��ҵ�����پ�ҵ':
                        v = '��' if i.get('��ҵ������Ա') == '��' else '��'
                # ̨��5
                elif '5ʧҵ��Ա�پ�ҵ��Ϣ��ϸ̨��' in file:
                    if k == '�Ļ��̶�' and not v:
                        v = i.get('ѧ��')
                    if k == '��������':
                        v = '����'
                    if k == '�����س�' and not v:
                        v = i.get('���ܵȼ�֤��', '��')
                    if k == 'ʧҵ��Ա����':
                        v = '(9)'
                    if k == 'ʧҵʱ��':
                        d = i.get('�ǼǾ�ҵʱ�䣨��/��/�գ�')
                        if not d:
                            d = i.get('��ҵʱ��')
                        if d:
                            s = d.strftime('%Y/%m/%d')
                            l = s.split('/')
                            l[1] = str(d.month - 1)
                            l2 = '/'.join(l)
                            v = datetime.strptime(l2, "%Y/%m/%d")
                    if k == '�پ�ҵʱ��' and not v:
                        v = i.get('�ǼǾ�ҵʱ�䣨��/��/�գ�')
                    if k == '��ҵ����':
                        if i.get('��ҵ��ʽ') == '��λ��ҵ':
                            v = '(3)'
                        elif i.get('��ҵ��ʽ') == '����ҵ':
                            v = '(5)'
                    if k == '�־�ҵ��λ':
                        v = i.get('��ҵ��λ(����ҵ����幤�����ݣ�')
                    if k == '������ҵ':
                        v = i.get('���²�ҵ����')
                # ̨��4
                elif '4ʧҵ��Ա����̨��' in file:
                    if k == '�Ƿ�Ǽ�ʧҵ��Ա':
                        v = '��'
                    if k == '�����س�':
                        self.jnct = v = i.get('���ܵȼ�֤��')
                    if k == 'ʧҵʱ��':
                        d = i.get('�ǼǾ�ҵʱ�䣨��/��/�գ�')
                        if d:
                            s = d.strftime('%Y/%m/%d')
                            l = s.split('/')
                            l[1] = str(d.month - 1)
                            l2 = '/'.join(l)
                            v = datetime.strptime(l2, "%Y/%m/%d")
                    if k == '�Ļ��̶�' and not v:
                        v = i.get('ѧ��')
                    if k == 'ʧҵ��Ա����':
                        v = '(9)'
                    if k == '��������':
                        v = '����'
                    if k == '��ȡʧҵ���ս���ֹʱ��':
                        v = '��'
                    if k == '��ְ����':
                        v = i.get('��ҵ��λ')
                    if k == '��ѵ����':
                        v = '��'
                    if k == '��ҵ��������':
                        v = '��1��'
                    if k == '����':
                        v = '�ж���'
                    if k == '�ȼ�':
                        if '��' in self.jnct:
                            v = '�߱�����'
                        else:
                            v = 'ȱ������'
                elif "3��ҵ������Ա����̨��" in file:
                    if k == "�Ļ��̶�":
                        v = i.get('ѧ��')
                    if k == "�����س�":
                        v = i.get('���ܵȼ�֤��')
                    if k == "��ҵ������Ա����":
                        v = '��'
                    if k == "��ͥסַ":
                        v = 'С���'
                    if k == "�پ�ҵʱ��":
                        v = i.get('�ǼǾ�ҵʱ�䣨��/��/�գ�')
                    if k == '��ҵȥ��':
                        if i.get('����ҵ') == '��':
                            v = '��'
                        elif i.get('��λ��ҵ') == '��':
                            v = '��'
                        elif i.get('���幤�̻�') == '��':
                            v = '��'
                        elif i.get('�����Ը�λ') == '��':
                            v = '��'
                    if k == '�Ƿ�ǩ���Ͷ���ͬ':
                        v = '��'
                    if k == '��ͬ����':
                        v = '��'
                    if k == '�Ƿ��ҵԮ������':
                        v = ' ��'
                    if k == '��ҵԮ����ʽ':
                        v = '��'
                elif "ȫʡ����ģ��" in file:
                    if k == '��':
                        v = '�����о�ҵ��������'
                    if k == '�������ƣ��أ�������':
                        v = '�ε���������Դ����ᱣ�Ϸ�������,23030432'
                    if k == '��ҵ��ʽ':
                        if i.get('��λ��ҵ') == '��':
                            v = '��λ��ҵ'
                        elif i.get('���幤�̻�') == '��':
                            v = '���幤�̻�'
                        elif i.get('����ҵ') == '��':
                            v = '����ҵ'
                        elif i.get('�����Ը�λ') == '��':
                            v = '�����Ը�λ����'
                    if k == '��ҵ����':
                        if i.get('ʧҵ��Ա�پ�ҵ') == '��':
                            v = 'ʧҵ�پ�ҵ'
                        else:
                            v = '���ξ�ҵ'
                    if k == '������ҵ(��ҵ��ʽΪ��λ��ҵʱ����)':
                        v = i.get('��ҵ����')
                    if k == '���²�ҵ����':
                        v = i.get('��ҵ����')
                    if k == '�Ǽ�ʧҵ��Ա':
                        if i.get('ʧҵ��Ա�پ�ҵ') == '��':
                            v = '��'
                        else:
                            v = '��'
                    if k == '��ҵ������Ա':
                        v = i.get('��ҵ�����پ�ҵ')
                    if k == '�Ƿ�¼��𱣣���/��':
                        v = '��'
                    if k == '�Ƿ��ǲм��ˣ���/��':
                        v = '��'
                    if k == '�Ƿ������۾��ˣ���/��':
                        v = '��'
                lis.append(v)
            value_lis.append(lis)
        # pprint(value_lis)
        # ��������к�
        insert_num = num + max_num
        if mer_lis:
            self.write_before(ws, mer_lis)
        for i, value in enumerate(value_lis, 1):
            self.insert_row(ws, value, insert_num + i)
        if mer_lis:
            self.write_tail(ws, mer_lis, len(value_lis))
        out_path = os.path.join(self.out_path, file)
        wb.save(out_path)
        print(f'�ļ�{out_path}��д�����')

    def write_before(self, ws, mer_lis):
        for i in mer_lis:
            # num1, num2 = re.findall('[0-9]+', i)
            # ȡ����Ԫ��ϲ�
            ws.unmerge_cells(i)
        
    def write_tail(self, ws, mer_lis, rows):
        for i in mer_lis:
            num1, num2 = re.findall('[0-9]+', i)
            # ��Ԫ��ϲ�
            start, end = re.findall('[A-Z]+', i)
            ws.merge_cells(f'{start}{int(num1) + rows}:{end}{int(num2) + rows}')

    def run_smz(self, smz_path, out_path):
        # ����ʵ������Ϣ������̨��5��6
        self.read_file(smz_path, 2, 3)
        if out_path:
            self.out_path = out_path
        self.write_excel('3��ҵ������Ա����̨��.xlsx', 4, 5)
        self.write_excel('4ʧҵ��Ա����̨��.xlsx', 4, 5)
        self.write_excel('5ʧҵ��Ա�پ�ҵ��Ϣ��ϸ̨��.xlsx', 4, 5)
        self.write_excel('6�¾�ҵ��Ա��Ϣ̨��.xlsx', 4, 5)

    def run_4to12(self, tz4_path, out_path):
        # ����̨��4������̨��12
        if os.path.isfile(tz4_path):
            self.read_file(tz4_path, 4, 5)
        else:
            self.read_file(os.path.join(tz4_path, '4ʧҵ��Ա����̨��.xlsx'), 4, 5)
        if out_path:
            self.out_path = out_path
        self.write_excel('12��ְ��Ա�Ǽ�̨��.xlsx', 4, 4)
    
    def run_7to15(self, tz7_path, out_path):
        # ��̨��7����̨��15
        self.read_file(tz7_path, 4, 4)
        if out_path:
            self.out_path = out_path
        self.write_excel('15������Ա��������������Ϣ̨��.xlsx', 4, 5)
    
    def run_gb(self, smz_path, out_path):
        # ����ʵ������Ϣ������̨�˹���ʵ����
        self.read_file(smz_path, 2, 3)
        if out_path:
            self.out_path = out_path
        self.write_excel('С�������ȫʡ����ģ��.xlsx', 2, 3)

    def to_smz(self):
        # ����������Ϣ����ʵ����
        pass

    def main(self):
        # ����ʵ������Ϣ������̨��3��4��5��6
        self.run_smz()

        # ����̨��4������̨��12
        self.run_4to12()

        # ��̨��7����̨��15
        # self.run_7to15()


if __name__ == '__main__':
    jctz = JCTZ()
    # jctz.run_smz(smz_path='C:/Users/Administrator/Desktop/����2024��4��ʵ����̨��20240426��.xlsx', out_path='')
    # jctz.run_smz(smz_path='C:/Users/XBD/Desktop/ʵ����.xlsx', out_path='C:/Users/XBD/Desktop/̨�˽��')
    jctz.run_4to12(tz4_path='C:/Users/XBD/Desktop/̨�˽��/4ʧҵ��Ա����̨��.xlsx', out_path='C:/Users/XBD/Desktop/̨�˽��')
