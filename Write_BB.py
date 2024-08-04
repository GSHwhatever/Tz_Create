# -*- coding:gbk -*-
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from collections import Counter
from datetime import datetime
import os


class Write:

    def __init__(self) -> None:
        self.smz_path = r'D:/weixin/Other/С�������7���¾�ҵ��Աʵ����.xlsx'
        out_path = os.path.join(os.environ['USERPROFILE'], 'Desktop', '̨�˽��', '��ҵͳ�Ʊ���.xlsx')
        self.bb_path = out_path if os.path.exists(out_path) else os.path.join(os.path.dirname(__file__), 'template_excel', '��ҵͳ�Ʊ���.xlsx')
        self.out_path = os.path.join(os.environ['USERPROFILE'], 'Desktop', '��ҵͳ�Ʊ���.xlsx')
        self.row = datetime.now().month + 5
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
    
    def read(self):
        smz_wb = load_workbook(self.smz_path)
        all_ws = smz_wb['�¾�ҵ��Ա']
        xzjy_ws = smz_wb['������ҵ��Ա']
        zrjy_ws = smz_wb['��Ȼ��Ա����ҵ��']
        syry_ws = smz_wb['ʧҵ��Ա���']

        xzjy_num = len([i for i in xzjy_ws['A'] if i.value]) - 2        # ������ҵ����
        zrjy_num = len([i for i in zrjy_ws['A'] if i.value]) - 2        # ��Ȼ��Ա����
        
        syzjy_num = len([i for i in all_ws['K'] if i.value and i.value == 'ʧҵ�پ�ҵ'])        # ʧҵ�پ�ҵ����
        jykn_num = len([i for i in all_ws['R'] if i.value and i.value == '��'])     # ��ҵ��������

        xl_num = len([i for i in all_ws['G'] if i.value and i.value in ('��ѧר��', '��ѧ����', '˶ʿ�о���')])     # ��ר����ѧ������
        nv_num = len([i for i in all_ws['X'] if i.value and i.value == 'Ů'])       # Ů������
        cylb_dic = Counter([i.value for i in all_ws['P'] if i.value and i.value in ('��һ��ҵ', '�ڶ���ҵ', '������ҵ')])     # ��ҵ������
        jyqd_dic = Counter([i.value for i in all_ws['J'] if i.value and i.value in ('��λ��ҵ', '����ҵ', '���幤�̻�', '�����Ը�λ����')])     # ��ҵ��������
        xydm_lis = [i.value for i, j in zip(all_ws['L'], all_ws['J']) if j == '��λ��ҵ' and i.value and i.value.isalnum()]      # ��λ��ҵ��Ա��λͳһ��ҵ����
        sydw_num = len([i for i in xydm_lis if i[0] == '1'])      # ����ͳһ���ô���ɸѡ��һλΪ'1'��Ϊ������ҵ��λ
        nlhf_dic = Counter([i.value for i in all_ws['AF'] if i.value and i.value in ('16-24', '25-45', '46-60')])     # ���仮�ּ���

        syry_num = len([i for i in syry_ws['A'] if i.value]) - 2        # ʧҵ��Ա����
        synx_num = len([i for i in syry_ws['C'] if i.value and i.value == 'Ů'])        # ʧҵ��Ա����Ů������
        sykn_num = len([i for i in syry_ws['M'] if i.value and i.value == '��'])        # ʧҵ��Ա���о�ҵ��������

        hyhf_dic = Counter([i.value for i in all_ws['N'] if i.value and i.value in self.industries])
        
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
    
    def write(self):
        bb_wb = load_workbook(self.bb_path)
        for sheet, datas in [('����ͳEP1', self.EP1), ('ʡ��ҵ01', self.sjy01), ('02�����Ǽ�ʧҵ��Ա���', self.sy02), ('��ҵ����', self.hyhf)]:
            ws = bb_wb[sheet]
            for i, v in enumerate(datas, start=3):
                ws.cell(row=self.row, column=i, value=v)
        bb_wb.save(self.out_path)
        print(f'�ļ�{self.out_path}��д�����')

    def run(self, path, out_path):
        self.smz_path = path
        self.out_path = os.path.join(out_path, '��ҵͳ�Ʊ���.xlsx')
        self.read()
        self.write()
        

if __name__ == '__main__':
    W = Write()
    W.read()
    W.write()


        

