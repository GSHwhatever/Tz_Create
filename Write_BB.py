# -*- coding:gbk -*-
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from collections import Counter
from datetime import datetime
import os


class Write:

    def __init__(self) -> None:
        self.smz_path = r'C:/Users/GSH\Desktop/С�������2024ʵ����.xlsx'
        out_path = os.path.join(os.environ['USERPROFILE'], 'Desktop', '̨�˽��', '��ҵͳ�Ʊ���.xlsx')
        self.bb_path = out_path if os.path.exists(out_path) else os.path.join(os.path.dirname(__file__), 'template_excel', '��ҵͳ�Ʊ���.xlsx')
        self.out_path = os.path.join(os.environ['USERPROFILE'], 'Desktop', '��ҵͳ�Ʊ���.xlsx')
        self.industries = [
            "ũ���֡�������ҵ",
            "�ɿ�ҵ",
            "����ҵ",
            "������ȼ����ˮ�������͹�Ӧҵ",
            "����ҵ",
            "��ͨ���䡢�ִ�������ҵ",
            "��Ϣ���䡢�������������ҵ",
            "����������ҵ",
            "ס�޺Ͳ���ҵ",
            "����ҵ",
            "���ز�ҵ",
            "���޺��������ҵ",
            "��ѧ�о�����������͵��ʿ���ҵ",
            "ˮ���������͹�����ʩ����ҵ",
            "����������������ҵ",
            "����",
            "��������ᱣ�Ϻ���ḣ��ҵ",
            "�Ļ�������������ҵ",
            "��ѩ����ҵ",
            "��������������֯"
        ]
        self.EP1 = []
        self.sjy01= []
        self.sy02 = []
        self.hyhf = []
    
    def re_init(self):
        self.EP1 = []
        self.sjy01 = []
        self.sy02 = []
        self.hyhf = []
    
    def read(self, month : int):
        smz_wb = load_workbook(self.smz_path)
        all_ws = smz_wb['�¾�ҵ��Ա']
        xzjy_ws = smz_wb['������ҵ��Ա']
        zrjy_ws = smz_wb['��Ȼ��Ա����ҵ��']
        syry_ws = smz_wb['ʧҵ��Ա���']

        # print([v.value for i, v in zip(zrjy_ws['A'], zrjy_ws['I']) if (i.value and isinstance(v.value, datetime)) and v.value.month <= datetime.now().month])
        xzjy_num = len([i for i, v in zip(xzjy_ws['A'], xzjy_ws['J'])  if (i.value and isinstance(v.value, datetime)) and v.value.month <= month])        # ������ҵ����
        zrjy_num = len([i for i, v in zip(zrjy_ws['A'], zrjy_ws['I']) if (i.value and isinstance(v.value, datetime)) and v.value.month <= month])        # ��Ȼ��Ա����
        
        syzjy_num = len([i for i, v in zip(all_ws['K'], all_ws['H']) if i.value and i.value == 'ʧҵ�پ�ҵ' and isinstance(v.value, datetime) and v.value.month <= month])        # ʧҵ�پ�ҵ����
        jykn_num = len([i for i, v in zip(all_ws['R'], all_ws['H']) if i.value and i.value == '��' and isinstance(v.value, datetime) and v.value.month <= month])     # ��ҵ��������

        xl_num = len([i for i, v in zip(all_ws['G'], all_ws['H']) if i.value and i.value in ('��ѧר��', '��ѧ����', '˶ʿ�о���') and isinstance(v.value, datetime) and v.value.month <= month])     # ��ר����ѧ������
        nv_num = len([i for i, v in zip(all_ws['X'], all_ws['H']) if i.value and i.value == 'Ů' and isinstance(v.value, datetime) and v.value.month <= month])       # Ů������
        cylb_dic = Counter([i.value for i, v in zip(all_ws['P'], all_ws['H']) if i.value and i.value in ('��һ��ҵ', '�ڶ���ҵ', '������ҵ') and isinstance(v.value, datetime) and v.value.month <= month])     # ��ҵ������
        jyqd_dic = Counter([i.value for i, v in zip(all_ws['J'], all_ws['H']) if i.value and i.value in ('��λ��ҵ', '����ҵ', '���幤�̻�', '�����Ը�λ����') and isinstance(v.value, datetime) and v.value.month <= month])     # ��ҵ��������
        xydm_lis = [i.value for i, j, v in zip(all_ws['L'], all_ws['J'], all_ws['H']) if j == '��λ��ҵ' and i.value and i.value.isalnum() and isinstance(v.value, datetime) and v.value.month <= month]      # ��λ��ҵ��Ա��λͳһ��ҵ����
        sydw_num = len([i for i in xydm_lis if i[0] == '1'])      # ����ͳһ���ô���ɸѡ��һλΪ'1'��Ϊ������ҵ��λ
        nlhf_dic = Counter([i.value for i, v in zip(all_ws['AF'], all_ws['H']) if i.value and i.value in ('16-24', '25-45', '46-60') and isinstance(v.value, datetime) and v.value.month <= month])     # ���仮�ּ���

        syry_num = len([i for i in syry_ws['A'] if i.value]) - 2        # ʧҵ��Ա����
        synx_num = len([i for i in syry_ws['C'] if i.value and i.value == 'Ů'])        # ʧҵ��Ա����Ů������
        sykn_num = len([i for i in syry_ws['K'] if i.value and i.value == '��ҵ����'])        # ʧҵ��Ա���о�ҵ��������

        hyhf_dic = Counter([i.value for i, v in zip(all_ws['N'], all_ws['H']) if i.value and i.value in self.industries and isinstance(v.value, datetime) and v.value.month <= month])
        
        self.EP1.extend([xzjy_num, xzjy_num + zrjy_num, zrjy_num, syzjy_num, jykn_num])
        self.sjy01.extend([xzjy_num + zrjy_num, syzjy_num, jykn_num, xl_num, nv_num, cylb_dic.get('��һ��ҵ', 0), cylb_dic.get('�ڶ���ҵ', 0), cylb_dic.get('������ҵ', 0),
                           0, 0, jyqd_dic.get('��λ��ҵ', 0) - sydw_num, sydw_num, jyqd_dic.get('���幤�̻�', 0), jyqd_dic.get('����ҵ', 0), jyqd_dic.get('�����Ը�λ����', 0), nlhf_dic.get('16-24', 0),
                           nlhf_dic.get('25-45', 0), nlhf_dic.get('46-60', 0)])
        self.sy02.extend([syry_num, 0, 0, 0, 0, 0, syry_num, 0, 0, 0, 0, 0, synx_num, sykn_num])
        self.hyhf.extend([xzjy_num + zrjy_num])
        self.hyhf.extend([hyhf_dic.get(i, 0) for i in self.industries])
        # print(self.EP1)
        # print(self.sjy01)
        # print(self.sy02)
        # print(self.hyhf)
    
    def write(self, month : int, bb_wb):
        row = month + 5
        for sheet, datas in [('����ͳEP1', self.EP1), ('ʡ��ҵ01', self.sjy01), ('02�����Ǽ�ʧҵ��Ա���', self.sy02), ('��ҵ����', self.hyhf)]:
            ws = bb_wb[sheet]
            for i, v in enumerate(datas, start=3):
                ws.cell(row=row, column=i, value=v)

    def run(self, path, out_path):
        self.smz_path = path
        self.out_path = os.path.join(out_path, '��ҵͳ�Ʊ���.xlsx')
        smz_wb = load_workbook(self.smz_path)
        all_ws = {i.value.month: '' for i in smz_wb['�¾�ҵ��Ա']['H'] if i.value and isinstance(i.value, datetime)}
        bb_wb = load_workbook(self.bb_path)
        for month in list(all_ws.keys()):
            self.re_init()
            self.read(month)
            self.write(month, bb_wb)
        bb_wb.save(self.out_path)
        print(f'�ļ�{self.out_path}��д�����')
    
    def main(self):
        smz_wb = load_workbook(self.smz_path)
        all_ws = {i.value.month: '' for i in smz_wb['�¾�ҵ��Ա']['H'] if i.value and isinstance(i.value, datetime)}
        bb_wb = load_workbook(self.bb_path)
        for month in list(all_ws.keys()):
            self.re_init()
            self.read(month)
            self.write(month, bb_wb)
        bb_wb.save(self.out_path)
        print(f'�ļ�{self.out_path}��д�����')


if __name__ == '__main__':
    W = Write()
    W.main()
    


        

