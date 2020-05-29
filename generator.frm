VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "주민등록번호 생성기"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5775
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command3 
      Caption         =   "검사"
      Height          =   495
      Left            =   3840
      TabIndex        =   20
      Top             =   3720
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "번호 유효성 검사"
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   2535
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   1440
         TabIndex        =   18
         Text            =   "1234567"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   240
         TabIndex        =   17
         Text            =   "123456"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "-"
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Top             =   285
         Width           =   135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "기본 설정"
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2655
      Begin VB.ComboBox Combo5 
         Height          =   300
         Left            =   840
         TabIndex        =   23
         Text            =   "서울특별시"
         Top             =   2880
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   840
         TabIndex        =   21
         Text            =   "한국인"
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox Combo4 
         Height          =   300
         Left            =   840
         TabIndex        =   14
         Text            =   "10"
         Top             =   2520
         Width           =   855
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         Left            =   840
         TabIndex        =   12
         Text            =   "1900"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   840
         TabIndex        =   7
         Text            =   "85"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   840
         TabIndex        =   6
         Text            =   "09"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   840
         TabIndex        =   5
         Text            =   "22"
         Top             =   1800
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   840
         TabIndex        =   4
         Text            =   "남자"
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "지역 :"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2925
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "구분 :"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   765
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "갯수 :"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2565
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "모드 :"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   405
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "연도 :"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1125
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "   월 :"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1485
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "   일 :"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1845
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "성별 :"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2205
         Width           =   495
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "종료"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "생성"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   3255
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   0
      Text            =   "Form1.frx":0BC2
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

'주민등록번호 형식 123456-1234567
'000000-1234567
'1 : 성별 (1900년생 : 1,2 / 2000년 이후생 : 3,4)
'2 : 지역명 (지역 구분 : 0서울 1경기 2 강원 3,4 충청 5,6 전라도 7,8 경상도 9 제주)
'34 : 지역등록기관
'56 : 임의생성값
'7 : Checksum

'에러 검사
If IsNumeric(Text1.Text) = False Then
    Call MsgBox("연도는 숫자로만 입력되야 합니다.", vbCritical)
    Exit Sub
End If
If IsNumeric(Text2.Text) = False Then
    Call MsgBox("날짜는 숫자로만 입력되야 합니다.", vbCritical)
    Exit Sub
End If
If IsNumeric(Text3.Text) = False Then
    Call MsgBox("날짜는 숫자로만 입력되야 합니다.", vbCritical)
    Exit Sub
End If
If IsNumeric(Combo4.Text) = False Then
    Call MsgBox("갯수는 숫자로만 입력되야 합니다.", vbCritical)
    Exit Sub
End If

If Len(Text1.Text) <> 2 Then
    Call MsgBox("연도는 반드시 2자리여야 합니다.", vbCritical)
    Exit Sub
End If
If Len(Text2.Text) > 2 Then
    Call MsgBox("날짜는 반드시 2자리여야 합니다.", vbCritical)
    Exit Sub
End If
    If Len(Text3.Text) > 2 Then
    Call MsgBox("날짜는 반드시 2자리여야 합니다.", vbCritical)
    Exit Sub
End If

If Combo1.Text <> "남자" And Combo1.Text <> "여자" Then
    Call MsgBox("성별지정이 잘못되었습니다.", vbCritical)
    Exit Sub
End If
If Combo3.Text <> "1900" And Combo3.Text <> "2000" Then
    Call MsgBox("모드지정이 잘못되었습니다.", vbCritical)
    Exit Sub
End If
If Combo2.Text <> "한국인" And Combo2.Text <> "외국인" And Combo2.Text <> "재외국민" And Combo2.Text <> "외국국적동포" Then
    Call MsgBox("국적의 구분이 잘못 지정되었습니다.", vbCritical)
    Exit Sub
End If
If Combo5.Text <> "서울특별시" And Combo5.Text <> "경기도" And Combo5.Text <> "강원도" And Combo5.Text <> "충청도" And Combo5.Text <> "전라도" And Combo5.Text <> "경상도" And Combo5.Text <> "제주특별자치도" Then
    Call MsgBox("지역의 구분이 잘못 지정되었습니다.", vbCritical)
    Exit Sub
End If

If Combo4.Text > 1000 Then
    Call MsgBox("너무 무리하는거 아냐?", vbCritical)
    Exit Sub
End If

Text4 = vbCrLf & "생성중입니다." & vbCrLf & "잠시만 기다려 주십시요."

'갯수 지정만큼 루프를 돌려 주민등록번호를 생성한다.
For i = 1 To Combo4.Text

    '기본 설정에서 앞의 6자리 형식을 가져온다.
    '연도
    check1 = Mid(Text1.Text, 1, 1)
    check2 = Mid(Text1.Text, 2, 1)
    '월
    check3 = Mid(Text2.Text, 1, 1)
    check4 = Mid(Text2.Text, 2, 1)
    If Len(Text2.Text) <> 2 Then
        check3 = 0
        check4 = CInt(Text2.Text)
    End If
    monthd = CLng(check3 & check4)
    If monthd > 12 Or monthd <= 0 Then
        Call MsgBox("월 지정이 잘못되었습니다.", vbCritical)
        Exit For
        Exit Sub
    End If
    '일
    check5 = Mid(Text3.Text, 1, 1)
    check6 = Mid(Text3.Text, 2, 1)
    If Len(Text3.Text) <> 2 Then
        check5 = 0
        check6 = CInt(Text3.Text)
    End If
    dayd = CLng(check5 & check6)
    If monthd = 1 Or monthd = 3 Or monthd = 5 Or monthd = 7 Or monthd = 8 Or monthd = 10 Or monthd = 12 Then
        If dayd > 31 Then
            Call MsgBox("일 지정이 잘못되었습니다. 해당 월은 31일까지 입니다.", vbCritical)
            errorc = 1
            Exit For
        End If
    End If
    If monthd = 4 Or monthd = 6 Or monthd = 9 Or monthd = 11 Then
        If dayd > 30 Then
            Call MsgBox("일 지정이 잘못되었습니다. 해당 월은 30일까지 입니다.", vbCritical)
            errorc = 1
            Exit For
        End If
    End If
    If monthd = 2 And dayd > 29 Then
        Call MsgBox("일 지정이 잘못되었습니다. 해당 월은 29일(윤년 적용 시)까지 입니다.", vbCritical)
        errorc = 1
        Exit For
    End If
    '성별
    If Combo3.Text = 1900 And Combo1.Text = "남자" And Combo2.Text = "한국인" Then
        sex = 1
    End If
    If Combo3.Text = 1900 And Combo1.Text = "여자" And Combo2.Text = "한국인" Then
        sex = 2
    End If
    If Combo3.Text = 2000 And Combo1.Text = "남자" And Combo2.Text = "한국인" Then
        sex = 3
    End If
    If Combo3.Text = 2000 And Combo1.Text = "여자" And Combo2.Text = "한국인" Then
        sex = 4
    End If
    
    '외국인 등록번호 체계는 다음과 같다
    '000000-1234567
    '1 : 성별 (1900년생 : 5,6 / 2000년 이후생 : 7,8)
    '23 : 등록기관 (탈북자는 대부분 안성시)
    '45 : 임의생성값
    '6 : 외국인 구분 (7 : 해외국적동포, 8 : 재외국민, 9 : 순수 외국인)
    '7 : Checksum
    
    If Combo3.Text = 1900 And Combo1.Text = "남자" And Combo2.Text = "외국인" Then
        sex = 5
        id_type = 9
    End If
    If Combo3.Text = 1900 And Combo1.Text = "여자" And Combo2.Text = "외국인" Then
        sex = 6
        id_type = 9
    End If
    If Combo3.Text = 2000 And Combo1.Text = "남자" And Combo2.Text = "외국인" Then
        sex = 7
        id_type = 9
    End If
    If Combo3.Text = 2000 And Combo1.Text = "여자" And Combo2.Text = "외국인" Then
        sex = 8
        id_type = 9
    End If
    
    If Combo3.Text = 1900 And Combo1.Text = "남자" And Combo2.Text = "재외국민" Then
        sex = 5
        id_type = 8
    End If
    If Combo3.Text = 1900 And Combo1.Text = "여자" And Combo2.Text = "재외국민" Then
        sex = 6
        id_type = 8
    End If
    If Combo3.Text = 2000 And Combo1.Text = "남자" And Combo2.Text = "재외국민" Then
        sex = 7
        id_type = 8
    End If
    If Combo3.Text = 2000 And Combo1.Text = "여자" And Combo2.Text = "재외국민" Then
        sex = 8
        id_type = 8
    End If
    
    If Combo3.Text = 1900 And Combo1.Text = "남자" And Combo2.Text = "외국국적동포" Then
        sex = 5
        id_type = 7
    End If
    If Combo3.Text = 1900 And Combo1.Text = "여자" And Combo2.Text = "외국국적동포" Then
        sex = 6
        id_type = 7
    End If
    If Combo3.Text = 2000 And Combo1.Text = "남자" And Combo2.Text = "외국국적동포" Then
        sex = 7
        id_type = 7
    End If
    If Combo3.Text = 2000 And Combo1.Text = "여자" And Combo2.Text = "외국국적동포" Then
        sex = 8
        id_type = 7
    End If
    
    '랜덤함수 사용법 : Int((상한값 - 하한값 + 1) * Rnd() + 하한값)
    If Combo5.Text = "서울특별시" Then
        region = 0
    End If
    If Combo5.Text = "경기도" Then
        region = 1
    End If
    If Combo5.Text = "강원도" Then
        region = 2
    End If
    If Combo5.Text = "충청도" Then
        Randomize
        region = Int((4 - 3 + 1) * Rnd() + 3)
    End If
    If Combo5.Text = "전라도" Then
        Randomize
        region = Int((6 - 5 + 1) * Rnd() + 5)
    End If
    If Combo5.Text = "경상도" Then
        Randomize
        region = Int((8 - 7 + 1) * Rnd() + 7)
    End If
    If Combo5.Text = "제주특별자치도" Then
        region = 9
    End If
    
    '주민등록번호 형식 중 랜덤하게 생성할 5자리(0~9)를 만든다.
    Randomize
    rand1 = rand & Int(9 * Rnd())
    rand2 = rand & Int(9 * Rnd())
    rand3 = rand & Int(9 * Rnd())
    rand4 = rand & Int(9 * Rnd())
    If id_type Then
        rand4 = id_type
    End If

    '7번째 자리의 계산
    'xxxxxx - yyyyyyy
    '******   *******
    '234567   8923456
    '----------------
    '이 식으로 나온 값을 11로 나눈 나머지 값을 취한 뒤 11 - 나머지값 = X가 최종 7번째 자리수가 된다.
    magickey = 11 - (((check1 * 2) + (check2 * 3) + (check3 * 4) + (check4 * 5) + (check5 * 6) + (check6 * 7) + (sex * 8) + (region * 9) + (rand1 * 2) + (rand2 * 3) + (rand3 * 4) + (rand4 * 5)) Mod 11)
    If magickey = 10 Then
        magickey = 0
    End If
    If magickey = 11 Then
        magickey = 1
    End If

'최종 완성된 주민등록번호의 변수 정의
    If Len(i) = 1 Then
        i = 0 & i '순번을 1 2 에서 01 02 와 같은 형식으로 만들어준다.
    End If
    kcode = kcode & vbCrLf & " " & "[" & i & "]" & " " & check1 & check2 & check3 & check4 & check5 & check6 & "-" & sex & region & rand1 & rand2 & rand3 & rand4 & magickey

Next i

'결과 출력
Text4 = "         [생성 결과]" & vbCrLf & " " & kcode & vbCrLf
If errorc = 1 Then
    Text4 = vbCrLf & "지정된 값에 오류가 있으므로 생성이 중단되었습니다."
End If

End Sub

Private Sub Command2_Click()

'종료
Form1.Hide
Unload Me
End Sub

Private Sub Command3_Click()

If IsNumeric(Text5.Text) = False Then
    Call MsgBox("입력값은 숫자로만 입력되야 합니다.", vbCritical)
    Exit Sub
End If
If IsNumeric(Text6.Text) = False Then
    Call MsgBox("입력값은 숫자로만 입력되야 합니다.", vbCritical)
    Exit Sub
End If

If Len(Text5.Text) <> 6 Then
    Call MsgBox("첫번째 입력값은 반드시 6자리여야 합니다.", vbCritical)
    Exit Sub
End If
If Len(Text6.Text) <> 7 Then
    Call MsgBox("두번째 입력값은 반드시 7자리여야 합니다.", vbCritical)
    Exit Sub
End If

checkkey = 11 - (((CInt(Mid(Text5.Text, 1, 1)) * 2) + (CInt(Mid(Text5.Text, 2, 1)) * 3) + (CInt(Mid(Text5.Text, 3, 1)) * 4) + (CInt(Mid(Text5.Text, 4, 1)) * 5) + (CInt(Mid(Text5.Text, 5, 1)) * 6) + (CInt(Mid(Text5.Text, 6, 1)) * 7) + (CInt(Mid(Text6.Text, 1, 1)) * 8) + (CInt(Mid(Text6.Text, 2, 1)) * 9) + (CInt(Mid(Text6.Text, 3, 1)) * 2) + (CInt(Mid(Text6.Text, 4, 1)) * 3) + (CInt(Mid(Text6.Text, 5, 1)) * 4) + (CInt(Mid(Text6.Text, 6, 1)) * 5)) Mod 11)
If checkkey = 10 Then
    checkkey = 0
End If

If checkkey = 11 Then
    checkkey = 1
End If

'mid 함수로 text form에서 숫자를 따와도 문자형식으로 취급되기 때문에
'정수형 변수로 미리 선언을 하든지 정수형으로 변환해야 한다.
checker = CInt(Mid(Text6.Text, 7, 1))

If checker = checkkey Then
    Text4.Text = "[O] " & Text5.Text & "-" & Text6.Text & vbCrLf & "(올바른 주민등록번호)"
End If

If checker <> checkkey Then
    checker = Text5.Text & "-" & Mid(Text6.Text, 1, 6)
    Text4.Text = "[X] " & Text5.Text & "-" & Text6.Text & vbCrLf & "(틀린 주민등록번호)" & vbCrLf & vbCrLf & "올바른 주민등록번호는" & vbCrLf & checker & checkkey & " 입니다."
End If

End Sub

Private Sub Form_Load()

'콤보박스를 완성한다.
Combo3.AddItem 1900
Combo3.AddItem 2000
Combo2.AddItem "한국인"
Combo2.AddItem "외국인"
Combo2.AddItem "재외국민"
Combo2.AddItem "외국국적동포"
Combo1.AddItem "남자"
Combo1.AddItem "여자"
Combo4.AddItem 1
Combo4.AddItem 5
Combo4.AddItem 10
Combo4.AddItem 20
Combo4.AddItem 30
Combo4.AddItem 50
Combo4.AddItem 100
Combo4.AddItem 500
Combo4.AddItem 1000
Combo5.AddItem "서울특별시"
Combo5.AddItem "경기도"
Combo5.AddItem "강원도"
Combo5.AddItem "충청도"
Combo5.AddItem "전라도"
Combo5.AddItem "경상도"
Combo5.AddItem "제주특별자치도"
End Sub
