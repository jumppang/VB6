����

��������

ũ��(Byte)

ǥ������

��뿹

������

Byte

1

0 ~ 255

Dim data as Byte

Integer

2

-32,768 ~ 32,767 

?���ļ����� : % 

Dim data as Integer?

Dim data%

Long

4

-2,147,483,648 ~ 2,147,483,647 

?���ļ����� : & 

Dim data as Long 

?Dim data& 

�Ǽ���

Single

4

�� -3.402823E38 ~ -1.401298E-45
�� 1.401298E-45 ~ 3.402823E38 

?���ļ����� : ! 

Dim data as Single 

?Dim data!

Double

8

�� -1.79769313486232E308 ~ -4.94065645841247E-324
�� 4.94065645841247E-324 ~ 1.79769313486232E324 

?���ļ����� : # 

Dim data as Double 

?Dim data# 

����

Boolean

2

True �Ǵ� False (�ʱⰪ�� False)

Dim data as Boolean

��ȭ��

Currency

8

-922,337,203,685,477.5808 ~ 922,337,203,685,477.5807 

?���ļ����� : @ 

Dim data as Currency 

?Dim data@ 

��¥��

Date 

8


1000��1��1�� ~ 9999��10��31�� 0�� 0�� 0�� ~ 23�� 59�� 59��

�ʱⰪ�� 12�� 00�� 00�� �ݵ�� #�� # ���̿� ��� �ð��� :


Dim data as Date

data = #10/12/2011 12:30:00#?

���ڿ���

String

����/����

����: 10byte + ���ڿ�����(0~��2��)

����: 1~�� 65,400��

""��ٿ�ǥ���̿� ���, �ʱⰪ�� Null

����(1byte), ����(1byte), �ѱ�(2byte) 

?���ļ����� : $ 


Dim data as String

Dim data as String * 5(5byte���ڿ�) 

?Dim data$ 

������

Variant

16/24

���������� ������� �ʾ��� ���
16: Dobule������ ����
24: ���ڿ����� + 22byte(����)


Dim data as Variant

Dim data

��ü��

Object 

4


��� ��ü ����

Set���� ����, ��ü ��ü�� �Ҵ�

Class >= ��ü(Instance) >= ��ü(Object)

Dim data as Object

�����

������

���������

�����

��������

 

[Private | Public] Type varname
elementname [([subscripts])] As type
. . .
End Type
[��ó] [VB6] ���������� �������� ��ȯ�Լ� �� ���� ����|�ۼ��� ��ǳ��ġ�»��