유형

데이터형

크기(Byte)

표현범위

사용예

정수형

Byte

1

0 ~ 255

Dim data as Byte

Integer

2

-32,768 ~ 32,767 

?형식선언문자 : % 

Dim data as Integer?

Dim data%

Long

4

-2,147,483,648 ~ 2,147,483,647 

?형식선언문자 : & 

Dim data as Long 

?Dim data& 

실수형

Single

4

음 -3.402823E38 ~ -1.401298E-45
양 1.401298E-45 ~ 3.402823E38 

?형식선언문자 : ! 

Dim data as Single 

?Dim data!

Double

8

음 -1.79769313486232E308 ~ -4.94065645841247E-324
양 4.94065645841247E-324 ~ 1.79769313486232E324 

?형식선언문자 : # 

Dim data as Double 

?Dim data# 

논리형

Boolean

2

True 또는 False (초기값은 False)

Dim data as Boolean

통화형

Currency

8

-922,337,203,685,477.5808 ~ 922,337,203,685,477.5807 

?형식선언문자 : @ 

Dim data as Currency 

?Dim data@ 

날짜형

Date 

8


1000년1월1일 ~ 9999년10월31일 0시 0분 0초 ~ 23시 59분 59초

초기값은 12시 00분 00초 반드시 #과 # 사이에 기술 시간은 :


Dim data as Date

data = #10/12/2011 12:30:00#?

문자열형

String

가변/고정

가변: 10byte + 문자열길이(0~약2조)

고정: 1~약 65,400자

""겹다옴표사이에 기술, 초기값은 Null

숫자(1byte), 영문(1byte), 한글(2byte) 

?형식선언문자 : $ 


Dim data as String

Dim data as String * 5(5byte문자열) 

?Dim data$ 

가변형

Variant

16/24

데이터형이 선언되지 않았을 경우
16: Dobule범위내 숫자
24: 문자열길이 + 22byte(문자)


Dim data as Variant

Dim data

개체형

Object 

4


모든 개체 참조

Set문이 선행, 개체 자체를 할당

Class >= 개체(Instance) >= 객체(Object)

Dim data as Object

사용자

정의형

사용자정의

사용자

정의형식

 

[Private | Public] Type varname
elementname [([subscripts])] As type
. . .
End Type
[출처] [VB6] 데이터형과 데이터형 변환함수 및 변수 선언|작성자 폭풍우치는사네