# coding=gbk
import xlrd
from xlwt import *
import pandas as pd
import openpyxl as op
import datetime

add_list=[]
add_list_new=[]

dic1 ={'1': '����ͷ��Ŀ�����嵥��ϸ.xlsx',
       '2': '�����ջ���Ŀ�����嵥��ϸ.xlsx',
       '3': '������Ŀ�����嵥��ϸ.xlsx',
       '4': '�ֱ�һ����Ŀ�����嵥��ϸ.xlsx',
       '5': '�ֱ������Ŀ�����嵥��ϸ.xlsx',
       '6': 'RFID���ܻ��������Ŀ�����嵥��ϸ.xlsx',
       '7': 'PHAI��Ŀ�����嵥��ϸ.xlsx',
       '8': 'ELO�к�����Ŀ�����嵥��ϸ.xlsx',
       '9': 'IOT��Ŀ�����嵥��ϸ.xlsx',
       '10': 'IOT-485��Ŀ�����嵥��ϸ.xlsx',
       }

fileName='���̶���������ϸ.xlsx'
bk=xlrd.open_workbook(fileName)
shxrange=range(bk.nsheets)

try:
  sh=bk.sheet_by_name("Sheet1")
except:
  print("�������")
nrows=sh.nrows #��ȡ����
book = Workbook(encoding='utf-8')
sheet = book.add_sheet('Sheet1') #����һ��sheet
for i in range(0,nrows):
  row_data=sh.row_values(i)
  #��ȡ��i�е�3������
  #sh.cell_value(i,3)
  #---------д���ļ���excel--------

  # print ("д�� "+str(i)+" ��")
  sheet.write(i,0, label = sh.cell_value(i,4))
  sheet.write(i,1, label = sh.cell_value(i,23)) #���1�е�1��д���ȡ����ֵ
  sheet.write(i,2, label = sh.cell_value(i,24))  #���1�е�2��д���ȡ����ֵ
  sheet.write(i,3, label = sh.cell_value(i,15))
  sheet.write(i,4, label = sh.cell_value(i,34))
  sheet.write(i,5, label = sh.cell_value(i,35))
  sheet.write(i,6, label = sh.cell_value(i,5))
  sheet.write(i,7, label = sh.cell_value(i,1))
  sheet.write(i,8, label=sh.cell_value(i,12))
  sheet.write(i,9, label=sh.cell_value(i,36))
  sheet.write(i,10, label=sh.cell_value(i,37))
  sheet.write(i,11, label=sh.cell_value(i,39))
book.save('Bridge.xls')


shijian = datetime.date.today()
print(shijian)
print('*************************')
print("1:����ͷ")
print("2:�����ջ�")
print('3:����')
print('4:�ֱ�һ��')
print('5:�ֱ����')
print('6:RFID���ܻ������')
print('7:PHAI')
print('8:ELO�к���')
print('9:IOT')
print('10:IOT-485')
print('*************************')
num = input('������Ҫ�Ǽ�̨�����ƣ�')

add_Name_list5 = []
add_Market = []
add_Name_list5 = []
add_Market = []
add_Name_list = []
add_Name_list2 = []
add_Name_list4 = []
add_Product_id_list = []
add_Product_name_num = []
add_Product_name_list = []
add_Name_list6 = []
add_Name_list3 = []
add_Order_num = []
add_Price = []

add_Name_list5_new = []
add_Market_new = []
add_Name_list5_new = []
add_Market_new = []
add_Name_list_new = []
add_Name_list2_new = []
add_Name_list4_new = []
add_Product_id_list_new = []
add_Product_name_num_new = []
add_Product_name_list_new = []
add_Name_list6_new = []
add_Name_list3_new = []
add_Order_num_new = []
add_Price_new = []



if dic1[num] == '����ͷ��Ŀ�����嵥��ϸ.xlsx':
    df = pd.read_excel('����ͷ��Ŀ�����嵥��ϸ.xlsx', sheet_name = 'Sheet1')
    JDE_list = df['�г�/�ִ�'].values
    df = pd.read_excel('Bridge.xls', sheet_name = 'Sheet1')
    Name_list = df['�跽���˹�˾����'].values
    Name_list2 = df['�跽��Ʊ��λ'].values
    Name_list3 = df['����'].values
    Name_list4 = df['�ϱ�����'].values
    Name_list5 = df['��������'].values
    Name_list6 = df['JDE�ʲ���'].values
    Product_id_list = df['JDE������'].values
    Product_name_num = df['�豸JDE����'].values
    Product_name_list = df['��ʩ/�豸����'].values
    Order_num=df['�ɹ�����'].values
    Price = df['���ۣ�˰��'].values
    Market = df['�г�'].values
    today = datetime.date.today()


    def write():
        Apend_len = len(JDE_list) + 2
        bg = op.load_workbook('����ͷ��Ŀ�����嵥��ϸ.xlsx')
        sheet = bg['Sheet1']
        a = 0
        while True:
            if Product_name_list[a] == '����AI����ͷDS2XA2346FMLS-1̨' :
                add_Name_list5.append(Name_list5[a])
                add_Market.append(Market[a])
                add_Name_list.append(Name_list[a])
                add_Name_list2.append(Name_list2[a])
                add_Product_id_list.append(Product_id_list[a])
                add_Product_name_num.append(Product_name_num[a])
                add_Product_name_list.append(Product_name_list[a])
                add_Name_list6.append(Name_list6[a])
                add_Name_list3.append(Name_list3[a])
                add_Name_list4.append(Name_list4[a])
                add_Order_num.append(Order_num[a])
                add_Price.append(Price[a])

            a = a + 1
            if a == len(Product_id_list):
                break
        print('*************************')
        print('����¼��Ŀ������',len(add_Name_list5))
        for b in range(0,len(add_Name_list5)):

            sheet.cell(Apend_len, 1, add_Name_list5[b])
            sheet.cell(Apend_len, 2, today)
            sheet.cell(Apend_len, 3, add_Market[b])
            sheet.cell(Apend_len, 4, add_Name_list3[b])
            sheet.cell(Apend_len, 5, add_Name_list[b])
            sheet.cell(Apend_len, 6, add_Name_list2[b])
            sheet.cell(Apend_len, 7, add_Product_id_list[b])
            sheet.cell(Apend_len, 8, add_Product_name_num [b])
            sheet.cell(Apend_len, 9, add_Product_name_list[b])
            sheet.cell(Apend_len, 10, add_Order_num[b])
            sheet.cell(Apend_len, 11, add_Price[b])
            sheet.cell(Apend_len, 12, add_Name_list6[b])
            if add_Name_list4 [b] == 'x':
                sheet.cell(Apend_len, 16, '�µ�')
            else:

                print('��������',add_Name_list3[b])

            Apend_len += 1
        bg.save('����ͷ��Ŀ�����嵥��ϸ.xlsx')

    write()

if dic1[num] == '�����ջ���Ŀ�����嵥��ϸ.xlsx':
    df = pd.read_excel('�����ջ���Ŀ�����嵥��ϸ.xlsx', sheet_name = 'Sheet1')
    JDE_list = df['�г�/�ִ�'].values
    df = pd.read_excel('Bridge.xls', sheet_name = 'Sheet1')
    Name_list = df['�跽���˹�˾����'].values
    Name_list2 = df['�跽��Ʊ��λ'].values
    Name_list3 = df['����'].values
    Name_list4 = df['�ϱ�����'].values
    Name_list5 = df['��������'].values
    Name_list6 = df['JDE�ʲ���'].values
    Product_id_list = df['JDE������'].values
    Product_name_num = df['�豸JDE����'].values
    Product_name_list = df['��ʩ/�豸����'].values
    Order_num=df['�ɹ�����'].values
    Price = df['���ۣ�˰��'].values
    Market = df['�г�'].values
    today = datetime.date.today()


    def write():
        Apend_len = len(JDE_list) + 2
        bg = op.load_workbook('�����ջ���Ŀ�����嵥��ϸ.xlsx')
        sheet = bg['Sheet1']
        a = 0
        while True:
            if Product_name_list[a] == '˫��˫�Ȳ�������������-1��' or Product_name_list[a] == '˫���Ȳ�������������-1��' or Product_name_list[a] == '˫��˫�Ȳ���������������ɫ��-1��' or Product_name_list[a] == '˫���Ȳ���������������ɫ��-1��' or Product_name_list[a] == '������ľ������������-1��' or Product_name_list[a] == '��ˮ������֧��-1̨' or Product_name_list[a] == '�����U������֧��-1̨' or Product_name_list[a] == '������ľ��������������ɫ��-1��'  or Product_name_list[a] == '�����U����֧��-1̨' or Product_name_list[a] == '����˫��ľ��������������ɫ��-1��' or Product_name_list[a] == '����˫��ľ��������������ɫ��-1��':
                add_Name_list5.append(Name_list5[a])
                add_Market.append(Market[a])
                add_Name_list.append(Name_list[a])
                add_Name_list2.append(Name_list2[a])
                add_Product_id_list.append(Product_id_list[a])
                add_Product_name_num.append(Product_name_num[a])
                add_Product_name_list.append(Product_name_list[a])
                add_Name_list6.append(Name_list6[a])
                add_Name_list3.append(Name_list3[a])
                add_Name_list4.append(Name_list4[a])
                add_Order_num.append(Order_num[a])
                add_Price.append(Price[a])

            a = a + 1
            if a == len(Product_id_list):
                break
        print('*************************')
        print('����¼��Ŀ������',len(add_Name_list5))
        for b in range(0,len(add_Name_list5)):

            sheet.cell(Apend_len, 1, add_Name_list5[b])
            sheet.cell(Apend_len, 2, today)
            sheet.cell(Apend_len, 3, add_Market[b])
            sheet.cell(Apend_len, 4, add_Name_list3[b])
            sheet.cell(Apend_len, 5, add_Name_list[b])
            sheet.cell(Apend_len, 6, add_Name_list2[b])
            sheet.cell(Apend_len, 7, add_Product_id_list[b])
            sheet.cell(Apend_len, 8, add_Product_name_num [b])
            sheet.cell(Apend_len, 9, add_Product_name_list[b])
            sheet.cell(Apend_len, 10, add_Order_num[b])
            sheet.cell(Apend_len, 11, add_Price[b])
            sheet.cell(Apend_len, 12, add_Name_list6[b])
            if add_Name_list4 [b] == 'x':
                sheet.cell(Apend_len, 16, '�µ�')
            else:

                print('��������',add_Name_list3[b])

            Apend_len += 1
        bg.save('�����ջ���Ŀ�����嵥��ϸ.xlsx')

    write()

if dic1[num] == '������Ŀ�����嵥��ϸ.xlsx':
    df = pd.read_excel('������Ŀ�����嵥��ϸ.xlsx', sheet_name = 'Sheet1')
    JDE_list = df['�г�/�ִ�'].values
    df = pd.read_excel('Bridge.xls', sheet_name = 'Sheet1')
    Name_list = df['�跽���˹�˾����'].values
    Name_list2 = df['�跽��Ʊ��λ'].values
    Name_list3 = df['����'].values
    Name_list4 = df['�ϱ�����'].values
    Name_list5 = df['��������'].values
    Name_list6 = df['JDE�ʲ���'].values
    Product_id_list = df['JDE������'].values
    Product_name_num = df['�豸JDE����'].values
    Product_name_list = df['��ʩ/�豸����'].values
    Order_num=df['�ɹ�����'].values
    Price = df['���ۣ�˰��'].values
    Market = df['�г�'].values
    today = datetime.date.today()


    def write():
        Apend_len = len(JDE_list) + 2
        bg = op.load_workbook('������Ŀ�����嵥��ϸ.xlsx')
        sheet = bg['Sheet1']
        a = 0
        while True:
            if Product_name_list[a] == '�Ų�ѶDT50' or Product_name_list[a] == '�Ų�Ѷ�����ǣ���ֽĤ��֧�ܣ�' or Product_name_list[a] == '�±��� Up391 �ƶ���ǩ��ӡ��' or Product_name_list[a] == '�Ų�Ѷ������(��ֽĤ��֧��)':
                add_Name_list5.append(Name_list5[a])
                add_Market.append(Market[a])
                add_Name_list.append(Name_list[a])
                add_Name_list2.append(Name_list2[a])
                add_Product_id_list.append(Product_id_list[a])
                add_Product_name_num.append(Product_name_num[a])
                add_Product_name_list.append(Product_name_list[a])
                add_Name_list6.append(Name_list6[a])
                add_Name_list3.append(Name_list3[a])
                add_Name_list4.append(Name_list4[a])
                add_Order_num.append(Order_num[a])
                add_Price.append(Price[a])

            a = a + 1
            if a == len(Product_id_list):
                break
        print('*************************')
        print('����¼��Ŀ������',len(add_Name_list5))
        for b in range(0,len(add_Name_list5)):

            sheet.cell(Apend_len, 1, add_Name_list5[b])
            sheet.cell(Apend_len, 2, today)
            sheet.cell(Apend_len, 3, add_Market[b])
            sheet.cell(Apend_len, 4, add_Name_list3[b])
            sheet.cell(Apend_len, 5, add_Name_list[b])
            sheet.cell(Apend_len, 6, add_Name_list2[b])
            sheet.cell(Apend_len, 7, add_Product_id_list[b])
            sheet.cell(Apend_len, 8, add_Product_name_num [b])
            sheet.cell(Apend_len, 9, add_Product_name_list[b])
            sheet.cell(Apend_len, 10, add_Order_num[b])
            sheet.cell(Apend_len, 11, add_Price[b])
            sheet.cell(Apend_len, 12, add_Name_list6[b])
            if add_Name_list4 [b] == 'x':
                sheet.cell(Apend_len, 16, '�µ�')
            else:

                print('��������',add_Name_list3[b])

            Apend_len += 1
        bg.save('������Ŀ�����嵥��ϸ.xlsx')

    write()

if dic1[num] == '�ֱ�һ����Ŀ�����嵥��ϸ.xlsx':
    df = pd.read_excel('�ֱ�һ����Ŀ�����嵥��ϸ.xlsx', sheet_name = 'Sheet1')
    JDE_list = df['�г�/�ִ�'].values
    df = pd.read_excel('Bridge.xls', sheet_name = 'Sheet1')
    Name_list = df['�跽���˹�˾����'].values
    Name_list2 = df['�跽��Ʊ��λ'].values
    Name_list3 = df['����'].values
    Name_list4 = df['�ϱ�����'].values
    Name_list5 = df['��������'].values
    Name_list6 = df['JDE�ʲ���'].values
    Product_id_list = df['JDE������'].values
    Product_name_num = df['�豸JDE����'].values
    Product_name_list = df['��ʩ/�豸����'].values
    Order_num=df['�ɹ�����'].values
    Price = df['���ۣ�˰��'].values
    Market = df['�г�'].values
    today = datetime.date.today()


    def write():
        Apend_len = len(JDE_list) + 2
        bg = op.load_workbook('�ֱ�һ����Ŀ�����嵥��ϸ.xlsx')
        sheet = bg['Sheet1']
        a = 0
        while True:
            if Product_name_list[a] == '41mm�����ֱ� ���Ű�' or Product_name_list[a] == '41mm�����ֱ���ͨ��':
                add_Name_list5.append(Name_list5[a])
                add_Market.append(Market[a])
                add_Name_list.append(Name_list[a])
                add_Name_list2.append(Name_list2[a])
                add_Product_id_list.append(Product_id_list[a])
                add_Product_name_num.append(Product_name_num[a])
                add_Product_name_list.append(Product_name_list[a])
                add_Name_list6.append(Name_list6[a])
                add_Name_list3.append(Name_list3[a])
                add_Name_list4.append(Name_list4[a])
                add_Order_num.append(Order_num[a])
                add_Price.append(Price[a])

            a = a + 1
            if a == len(Product_id_list):
                break
        print('*************************')
        print('����¼��Ŀ������',len(add_Name_list5))
        for b in range(0,len(add_Name_list5)):

            sheet.cell(Apend_len, 1, add_Name_list5[b])
            sheet.cell(Apend_len, 2, today)
            sheet.cell(Apend_len, 3, add_Market[b])
            sheet.cell(Apend_len, 4, add_Name_list3[b])
            sheet.cell(Apend_len, 5, add_Name_list[b])
            sheet.cell(Apend_len, 6, add_Name_list2[b])
            sheet.cell(Apend_len, 7, add_Product_id_list[b])
            sheet.cell(Apend_len, 8, add_Product_name_num [b])
            sheet.cell(Apend_len, 9, add_Product_name_list[b])
            sheet.cell(Apend_len, 10, add_Order_num[b])
            sheet.cell(Apend_len, 11, add_Price[b])
            sheet.cell(Apend_len, 12, add_Name_list6[b])
            if add_Name_list4 [b] == 'x':
                sheet.cell(Apend_len, 16, '�µ�')
            else:

                print('��������',add_Name_list3[b])

            Apend_len += 1
        bg.save('�ֱ�һ����Ŀ�����嵥��ϸ.xlsx')

    write()

if dic1[num] == '�ֱ������Ŀ�����嵥��ϸ.xlsx':
    df = pd.read_excel('�ֱ������Ŀ�����嵥��ϸ.xlsx', sheet_name = 'Sheet1')
    JDE_list = df['�г�/�ִ�'].values
    df = pd.read_excel('Bridge.xls', sheet_name = 'Sheet1')
    Name_list = df['�跽���˹�˾����'].values
    Name_list2 = df['�跽��Ʊ��λ'].values
    Name_list3 = df['����'].values
    Name_list4 = df['�ϱ�����'].values
    Name_list5 = df['��������'].values
    Name_list6 = df['JDE�ʲ���'].values
    Product_id_list = df['JDE������'].values
    Product_name_num = df['�豸JDE����'].values
    Product_name_list = df['��ʩ/�豸����'].values
    Order_num=df['�ɹ�����'].values
    Price = df['���ۣ�˰��'].values
    Market = df['�г�'].values
    today = datetime.date.today()


    def write():
        Apend_len = len(JDE_list) + 2
        bg = op.load_workbook('�ֱ������Ŀ�����嵥��ϸ.xlsx')
        sheet = bg['Sheet1']
        a = 0
        while True:
            if Product_name_list[a] == '42':
                add_Name_list5.append(Name_list5[a])
                add_Market.append(Market[a])
                add_Name_list.append(Name_list[a])
                add_Name_list2.append(Name_list2[a])
                add_Product_id_list.append(Product_id_list[a])
                add_Product_name_num.append(Product_name_num[a])
                add_Product_name_list.append(Product_name_list[a])
                add_Name_list6.append(Name_list6[a])
                add_Name_list3.append(Name_list3[a])
                add_Name_list4.append(Name_list4[a])
                add_Order_num.append(Order_num[a])
                add_Price.append(Price[a])

            a = a + 1
            if a == len(Product_id_list):
                break
        print('*************************')
        print('����¼��Ŀ������',len(add_Name_list5))
        for b in range(0,len(add_Name_list5)):

            sheet.cell(Apend_len, 1, add_Name_list5[b])
            sheet.cell(Apend_len, 2, today)
            sheet.cell(Apend_len, 3, add_Market[b])
            sheet.cell(Apend_len, 4, add_Name_list3[b])
            sheet.cell(Apend_len, 5, add_Name_list[b])
            sheet.cell(Apend_len, 6, add_Name_list2[b])
            sheet.cell(Apend_len, 7, add_Product_id_list[b])
            sheet.cell(Apend_len, 8, add_Product_name_num [b])
            sheet.cell(Apend_len, 9, add_Product_name_list[b])
            sheet.cell(Apend_len, 10, add_Order_num[b])
            sheet.cell(Apend_len, 11, add_Price[b])
            sheet.cell(Apend_len, 12, add_Name_list6[b])
            if add_Name_list4 [b] == 'x':
                sheet.cell(Apend_len, 16, '�µ�')
            else:

                print('��������',add_Name_list3[b])

            Apend_len += 1
        bg.save('�ֱ������Ŀ�����嵥��ϸ.xlsx')

    write()

if dic1[num] == 'RFID���ܻ��������Ŀ�����嵥��ϸ.xlsx':
    df = pd.read_excel('RFID���ܻ��������Ŀ�����嵥��ϸ.xlsx', sheet_name = 'Sheet1')
    JDE_list = df['�г�/�ִ�'].values
    df = pd.read_excel('Bridge.xls', sheet_name = 'Sheet1')
    Name_list = df['�跽���˹�˾����'].values
    Name_list2 = df['�跽��Ʊ��λ'].values
    Name_list3 = df['����'].values
    Name_list4 = df['�ϱ�����'].values
    Name_list5 = df['��������'].values
    Name_list6 = df['JDE�ʲ���'].values
    Product_id_list = df['JDE������'].values
    Product_name_num = df['�豸JDE����'].values
    Product_name_list = df['��ʩ/�豸����'].values
    Order_num=df['�ɹ�����'].values
    Price = df['���ۣ�˰��'].values
    Market = df['�г�'].values
    today = datetime.date.today()


    def write():
        Apend_len = len(JDE_list) + 2
        bg = op.load_workbook('RFID���ܻ��������Ŀ�����嵥��ϸ.xlsx')
        sheet = bg['Sheet1']
        a = 0
        while True:
            if Product_name_list[a] == 'RFID��ǩ����ϣ�280ֻ��׼װ��' or Product_name_list[a] == 'RFID��ǩ����ϣ�280ֻ��90ETC��' or Product_name_list[a] == 'RFID����ϣ�280ֻ��90ETC��-5��*56��/��' or Product_name_list[a] == 'RFID����ϣ�280ֻ��׼װ��-5��*56��/��' or Product_name_list[a] == '���ܻ������PAD��-1̨/��' or Product_name_list[a] == '���ܻ��������Ե���أ�-1̨/��' or Product_name_list[a] == '���ܻ���������ڻ�վ��-1̨/��' or Product_name_list[a] == '���ܻ�����������վ��-1̨/��' or Product_name_list[a] == '���ܻ�������Ž����ϵͳ��-1̨/��':
                add_Name_list5.append(Name_list5[a])
                add_Market.append(Market[a])
                add_Name_list.append(Name_list[a])
                add_Name_list2.append(Name_list2[a])
                add_Product_id_list.append(Product_id_list[a])
                add_Product_name_num.append(Product_name_num[a])
                add_Product_name_list.append(Product_name_list[a])
                add_Name_list6.append(Name_list6[a])
                add_Name_list3.append(Name_list3[a])
                add_Name_list4.append(Name_list4[a])
                add_Order_num.append(Order_num[a])
                add_Price.append(Price[a])

            a = a + 1
            if a == len(Product_id_list):
                break
        print('*************************')
        print('����¼��Ŀ������',len(add_Name_list5))
        for b in range(0,len(add_Name_list5)):

            sheet.cell(Apend_len, 1, add_Name_list5[b])
            sheet.cell(Apend_len, 2, today)
            sheet.cell(Apend_len, 3, add_Market[b])
            sheet.cell(Apend_len, 4, add_Name_list3[b])
            sheet.cell(Apend_len, 5, add_Name_list[b])
            sheet.cell(Apend_len, 6, add_Name_list2[b])
            sheet.cell(Apend_len, 7, add_Product_id_list[b])
            sheet.cell(Apend_len, 8, add_Product_name_num [b])
            sheet.cell(Apend_len, 9, add_Product_name_list[b])
            sheet.cell(Apend_len, 10, add_Order_num[b])
            sheet.cell(Apend_len, 11, add_Price[b])
            sheet.cell(Apend_len, 12, add_Name_list6[b])
            if add_Name_list4 [b] == 'x':
                sheet.cell(Apend_len, 16, '�µ�')
            else:

                print('��������',add_Name_list3[b])

            Apend_len += 1
        bg.save('RFID���ܻ��������Ŀ�����嵥��ϸ.xlsx')

    write()

if dic1[num] == 'PHAI��Ŀ�����嵥��ϸ.xlsx':
    df = pd.read_excel('PHAI��Ŀ�����嵥��ϸ.xlsx', sheet_name = 'Sheet1')
    JDE_list = df['�г�/�ִ�'].values
    df = pd.read_excel('Bridge.xls', sheet_name = 'Sheet1')
    Name_list = df['�跽���˹�˾����'].values
    Name_list2 = df['�跽��Ʊ��λ'].values
    Name_list3 = df['����'].values
    Name_list4 = df['�ϱ�����'].values
    Name_list5 = df['��������'].values
    Name_list6 = df['JDE�ʲ���'].values
    Product_id_list = df['JDE������'].values
    Product_name_num = df['�豸JDE����'].values
    Product_name_list = df['��ʩ/�豸����'].values
    Order_num=df['�ɹ�����'].values
    Price = df['���ۣ�˰��'].values
    Market = df['�г�'].values
    today = datetime.date.today()


    def write():
        Apend_len = len(JDE_list) + 2
        bg = op.load_workbook('PHAI��Ŀ�����嵥��ϸ.xlsx')
        sheet = bg['Sheet1']
        a = 0
        while True:
            if Product_name_list[a] == '��˴ʱ��POE����ģ�� TE_POE101' or Product_name_list[a] == '���Ծ�������� MLC2Y1' or Product_name_list[a] == '��ʤ�����̿����� MLCM01' or Product_name_list[a] == '��ʤ���ⱨ���� MLA001_L' or Product_name_list[a] == '��ʤ�����̿�������Դ������' :
                add_Name_list5.append(Name_list5[a])
                add_Market.append(Market[a])
                add_Name_list.append(Name_list[a])
                add_Name_list2.append(Name_list2[a])
                add_Product_id_list.append(Product_id_list[a])
                add_Product_name_num.append(Product_name_num[a])
                add_Product_name_list.append(Product_name_list[a])
                add_Name_list6.append(Name_list6[a])
                add_Name_list3.append(Name_list3[a])
                add_Name_list4.append(Name_list4[a])
                add_Order_num.append(Order_num[a])
                add_Price.append(Price[a])

            a = a + 1
            if a == len(Product_id_list):
                break
        print('*************************')
        print('����¼��Ŀ������',len(add_Name_list5))
        for b in range(0,len(add_Name_list5)):

            sheet.cell(Apend_len, 1, add_Name_list5[b])
            sheet.cell(Apend_len, 2, today)
            sheet.cell(Apend_len, 3, add_Market[b])
            sheet.cell(Apend_len, 4, add_Name_list3[b])
            sheet.cell(Apend_len, 5, add_Name_list[b])
            sheet.cell(Apend_len, 6, add_Name_list2[b])
            sheet.cell(Apend_len, 7, add_Product_id_list[b])
            sheet.cell(Apend_len, 8, add_Product_name_num [b])
            sheet.cell(Apend_len, 9, add_Product_name_list[b])
            sheet.cell(Apend_len, 10, add_Order_num[b])
            sheet.cell(Apend_len, 11, add_Price[b])
            sheet.cell(Apend_len, 12, add_Name_list6[b])
            if add_Name_list4 [b] == 'x':
                sheet.cell(Apend_len, 16, '�µ�')
            else:

                print('��������',add_Name_list3[b])

            Apend_len += 1
        bg.save('PHAI��Ŀ�����嵥��ϸ.xlsx')

    write()

if dic1[num] == 'ELO�к�����Ŀ�����嵥��ϸ.xlsx':
    df = pd.read_excel('ELO�к�����Ŀ�����嵥��ϸ.xlsx', sheet_name = 'Sheet1')
    JDE_list = df['�г�/�ִ�'].values
    df = pd.read_excel('Bridge.xls', sheet_name = 'Sheet1')
    Name_list = df['�跽���˹�˾����'].values
    Name_list2 = df['�跽��Ʊ��λ'].values
    Name_list3 = df['����'].values
    Name_list4 = df['�ϱ�����'].values
    Name_list5 = df['��������'].values
    Name_list6 = df['JDE�ʲ���'].values
    Product_id_list = df['JDE������'].values
    Product_name_num = df['�豸JDE����'].values
    Product_name_list = df['��ʩ/�豸����'].values
    Order_num=df['�ɹ�����'].values
    Price = df['���ۣ�˰��'].values
    Market = df['�г�'].values
    today = datetime.date.today()


    def write():
        Apend_len = len(JDE_list) + 2
        bg = op.load_workbook('ELO�к�����Ŀ�����嵥��ϸ.xlsx')
        sheet = bg['Sheet1']
        a = 0
        while True:
            if Product_name_list[a] == 'ELO 10��к�һ���':
                add_Name_list5.append(Name_list5[a])
                add_Market.append(Market[a])
                add_Name_list.append(Name_list[a])
                add_Name_list2.append(Name_list2[a])
                add_Product_id_list.append(Product_id_list[a])
                add_Product_name_num.append(Product_name_num[a])
                add_Product_name_list.append(Product_name_list[a])
                add_Name_list6.append(Name_list6[a])
                add_Name_list3.append(Name_list3[a])
                add_Name_list4.append(Name_list4[a])
                add_Order_num.append(Order_num[a])
                add_Price.append(Price[a])

            a = a + 1
            if a == len(Product_id_list):
                break
        print('*************************')
        print('����¼��Ŀ������',len(add_Name_list5))
        for b in range(0,len(add_Name_list5)):

            sheet.cell(Apend_len, 1, add_Name_list5[b])
            sheet.cell(Apend_len, 2, today)
            sheet.cell(Apend_len, 3, add_Market[b])
            sheet.cell(Apend_len, 4, add_Name_list3[b])
            sheet.cell(Apend_len, 5, add_Name_list[b])
            sheet.cell(Apend_len, 6, add_Name_list2[b])
            sheet.cell(Apend_len, 7, add_Product_id_list[b])
            sheet.cell(Apend_len, 8, add_Product_name_num [b])
            sheet.cell(Apend_len, 9, add_Product_name_list[b])
            sheet.cell(Apend_len, 10, add_Order_num[b])
            sheet.cell(Apend_len, 11, add_Price[b])
            sheet.cell(Apend_len, 12, add_Name_list6[b])
            if add_Name_list4 [b] == 'x':
                sheet.cell(Apend_len, 16, '�µ�')
            else:

                print('��������',add_Name_list3[b])

            Apend_len += 1
        bg.save('ELO�к�����Ŀ�����嵥��ϸ.xlsx')

    write()

if dic1[num] == 'IOT��Ŀ�����嵥��ϸ.xlsx':
    df = pd.read_excel('IOT��Ŀ�����嵥��ϸ.xlsx', sheet_name = 'Sheet1')
    JDE_list = df['�г�/�ִ�'].values
    df = pd.read_excel('Bridge.xls', sheet_name = 'Sheet1')
    Name_list = df['�跽���˹�˾����'].values
    Name_list2 = df['�跽��Ʊ��λ'].values
    Name_list3 = df['����'].values
    Name_list4 = df['�ϱ�����'].values
    Name_list5 = df['��������'].values
    Name_list6 = df['JDE�ʲ���'].values
    Product_id_list = df['JDE������'].values
    Product_name_num = df['�豸JDE����'].values
    Product_name_list = df['��ʩ/�豸����'].values
    Order_num=df['�ɹ�����'].values
    Price = df['���ۣ�˰��'].values
    Market = df['�г�'].values
    today = datetime.date.today()


    def write():
        Apend_len = len(JDE_list) + 2
        bg = op.load_workbook('IOT��Ŀ�����嵥��ϸ.xlsx')
        sheet = bg['Sheet1']
        a = 0
        while True:
            if Product_name_list[a] == 'IOT ������':
                add_Name_list5.append(Name_list5[a])
                add_Market.append(Market[a])
                add_Name_list.append(Name_list[a])
                add_Name_list2.append(Name_list2[a])
                add_Product_id_list.append(Product_id_list[a])
                add_Product_name_num.append(Product_name_num[a])
                add_Product_name_list.append(Product_name_list[a])
                add_Name_list6.append(Name_list6[a])
                add_Name_list3.append(Name_list3[a])
                add_Name_list4.append(Name_list4[a])
                add_Order_num.append(Order_num[a])
                add_Price.append(Price[a])

            a = a + 1
            if a == len(Product_id_list):
                break
        print('*************************')
        print('����¼��Ŀ������',len(add_Name_list5))
        for b in range(0,len(add_Name_list5)):

            sheet.cell(Apend_len, 1, add_Name_list5[b])
            sheet.cell(Apend_len, 2, today)
            sheet.cell(Apend_len, 3, add_Market[b])
            sheet.cell(Apend_len, 4, add_Name_list3[b])
            sheet.cell(Apend_len, 5, add_Name_list[b])
            sheet.cell(Apend_len, 6, add_Name_list2[b])
            sheet.cell(Apend_len, 7, add_Product_id_list[b])
            sheet.cell(Apend_len, 8, add_Product_name_num [b])
            sheet.cell(Apend_len, 9, add_Product_name_list[b])
            sheet.cell(Apend_len, 10, add_Order_num[b])
            sheet.cell(Apend_len, 11, add_Price[b])
            sheet.cell(Apend_len, 12, add_Name_list6[b])
            if add_Name_list4 [b] == 'x':
                sheet.cell(Apend_len, 16, '�µ�')
            else:

                print('��������',add_Name_list3[b])

            Apend_len += 1
        bg.save('IOT��Ŀ�����嵥��ϸ.xlsx')

    write()

if dic1[num] == 'IOT-485��Ŀ�����嵥��ϸ.xlsx':
        df = pd.read_excel('IOT-485��Ŀ�����嵥��ϸ.xlsx', sheet_name='Sheet1')
        JDE_list = df['�г�/�ִ�'].values
        df = pd.read_excel('Bridge.xls', sheet_name='Sheet1')
        Name_list = df['�跽���˹�˾����'].values
        Name_list2 = df['�跽��Ʊ��λ'].values
        Name_list3 = df['����'].values
        Name_list4 = df['�ϱ�����'].values
        Name_list5 = df['��������'].values
        Name_list6 = df['JDE�ʲ���'].values
        Product_id_list = df['JDE������'].values
        Product_name_num = df['�豸JDE����'].values
        Product_name_list = df['��ʩ/�豸����'].values
        Order_num = df['�ɹ�����'].values
        Price = df['���ۣ�˰��'].values
        Market = df['�г�'].values
        today = datetime.date.today()


        def write():
            Apend_len = len(JDE_list) + 2
            bg = op.load_workbook('IOT-485��Ŀ�����嵥��ϸ.xlsx')
            sheet = bg['Sheet1']
            a = 0
            while True:
                if Product_name_list[a] == 'IOT ��Ŀ�� 485Hub-1̨':
                    add_Name_list5.append(Name_list5[a])
                    add_Market.append(Market[a])
                    add_Name_list.append(Name_list[a])
                    add_Name_list2.append(Name_list2[a])
                    add_Product_id_list.append(Product_id_list[a])
                    add_Product_name_num.append(Product_name_num[a])
                    add_Product_name_list.append(Product_name_list[a])
                    add_Name_list6.append(Name_list6[a])
                    add_Name_list3.append(Name_list3[a])
                    add_Name_list4.append(Name_list4[a])
                    add_Order_num.append(Order_num[a])
                    add_Price.append(Price[a])

                a = a + 1
                if a == len(Product_id_list):
                    break
            print('*************************')
            print('����¼��Ŀ������', len(add_Name_list5))
            for b in range(0, len(add_Name_list5)):

                sheet.cell(Apend_len, 1, add_Name_list5[b])
                sheet.cell(Apend_len, 2, today)
                sheet.cell(Apend_len, 3, add_Market[b])
                sheet.cell(Apend_len, 4, add_Name_list3[b])
                sheet.cell(Apend_len, 5, add_Name_list[b])
                sheet.cell(Apend_len, 6, add_Name_list2[b])
                sheet.cell(Apend_len, 7, add_Product_id_list[b])
                sheet.cell(Apend_len, 8, add_Product_name_num[b])
                sheet.cell(Apend_len, 9, add_Product_name_list[b])
                sheet.cell(Apend_len, 10, add_Order_num[b])
                sheet.cell(Apend_len, 11, add_Price[b])
                sheet.cell(Apend_len, 12, add_Name_list6[b])
                if add_Name_list4[b] == 'x':
                    sheet.cell(Apend_len, 16, '�µ�')
                else:

                    print('��������', add_Name_list3[b])

                Apend_len += 1
            bg.save('IOT-485��Ŀ�����嵥��ϸ.xlsx')


        write()
num1=input('���س����ر�')