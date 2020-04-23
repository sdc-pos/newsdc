VERSION 5.00
Begin VB.Form F1010101 
   BackColor       =   &H00FFFFFF&
   Caption         =   "倉庫マスタメンテナンス(倉庫マスタのみ)"
   ClientHeight    =   6300
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   11295
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "商品化設定"
      Height          =   1695
      Left            =   6480
      TabIndex        =   41
      Top             =   2040
      Width           =   3615
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "商品化用以外"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   43
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "商品化用倉庫"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   42
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   7320
      MaxLength       =   3
      TabIndex        =   39
      Top             =   1440
      Width           =   495
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   2
      Left            =   6000
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   5640
      MaxLength       =   2
      TabIndex        =   12
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   4680
      MaxLength       =   2
      TabIndex        =   11
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   5640
      MaxLength       =   2
      TabIndex        =   10
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   4680
      MaxLength       =   2
      TabIndex        =   9
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   5640
      MaxLength       =   2
      TabIndex        =   8
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   4680
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1560
      Width           =   375
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   4
      Left            =   9120
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   5
      Left            =   8880
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      Left            =   3480
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   4  '全角ひらがな
      Index           =   0
      Left            =   1560
      MaxLength       =   16
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   3
      Left            =   6000
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command 
      Caption         =   "終  了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10320
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9480
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8640
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "データ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7800
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6480
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5640
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "削  除"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "更  新"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "（０～１００）"
      Height          =   255
      Index           =   14
      Left            =   7920
      TabIndex        =   40
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "発注％"
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   38
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "混載区分"
      Height          =   255
      Index           =   13
      Left            =   4920
      TabIndex        =   37
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   255
      Index           =   12
      Left            =   5280
      TabIndex        =   36
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "段"
      Height          =   255
      Index           =   11
      Left            =   4200
      TabIndex        =   35
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   255
      Index           =   10
      Left            =   5280
      TabIndex        =   34
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "連"
      Height          =   255
      Index           =   9
      Left            =   4200
      TabIndex        =   33
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   255
      Index           =   8
      Left            =   5280
      TabIndex        =   32
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "列"
      Height          =   255
      Index           =   7
      Left            =   4200
      TabIndex        =   31
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   255
      Index           =   6
      Left            =   8280
      TabIndex        =   30
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "使用可否"
      Height          =   255
      Index           =   5
      Left            =   7800
      TabIndex        =   29
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "倉庫分類"
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   28
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "倉庫名称"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   27
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "倉庫№"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   26
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "事業部区分"
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   25
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "F1010101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Text_Max    As Integer                 '画面項目別最大ｲﾝﾃﾞｯｸｽ
Dim Combo_Max   As Integer
Dim Command_Max As Integer
Dim Soko_Csv    As String

Private Function Data_Proc() As Integer

Dim FileNo          As Integer
Dim fileName        As String
Dim Ret             As Integer

Dim com             As Integer
Dim sts             As Integer

Dim c               As String * 128

    Call Input_Lock

    FileNo = FreeFile
    fileName = Soko_Csv
    
    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & Right(Trim(fileName), Len(Trim(fileName)) - Ret)

    On Error GoTo Error_Proc

    Open (fileName) For Output As FileNo
    
    Write #FileNo, "事業部区分", "倉庫№", "倉庫名称", "倉庫分類", "倉庫区分", "国内外", "使用可否", "混載可否", "列範囲", "連範囲", "段範囲", "発注点", "商品化倉庫"

    com = BtOpGetFirst
    Do
        DoEvents
        sts = BTRV(com, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "倉庫マスタ")
                Exit Function
        End Select
    
        Write #FileNo, StrConv(SOKOREC.JGYOBU, vbUnicode),
        
        
        If GetIni("SOKO_NO", StrConv(SOKOREC.Soko_No, vbUnicode), "SYS", c) Then
            Write #FileNo, StrConv(SOKOREC.Soko_No, vbUnicode),
        Else
            Write #FileNo, Trim(c),
        End If
        Write #FileNo, StrConv(SOKOREC.SOKO_NAME, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.SOKO_BUN, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.SOKO_KBN, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.NAIGAI, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.KAHI_KBN, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.KONS_KBN, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.RETU_START, vbUnicode) & "～" & StrConv(SOKOREC.RETU_END, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.REN_START, vbUnicode) & "～" & StrConv(SOKOREC.REN_END, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.DAN_START, vbUnicode) & "～" & StrConv(SOKOREC.DAN_END, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.ORDER_POINT, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.GOODS_ON_F, vbUnicode)
    
    
        com = BtOpGetNext
    Loop

    Close #FileNo
    
    Call Input_UnLock
    
    Beep
    MsgBox "「" & fileName & "」は正常に出力されました。"


    Exit Function
Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "が使用中です。"
        Data_Proc = False
    Else
        MsgBox "Err.Number= " & Err.Number
        Data_Proc = True
    End If


    Call Input_UnLock



End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    F1010101.MousePointer = vbHourglass

    Call Ctrl_Lock(F1010101)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1010101)


    F1010101.MousePointer = vbDefault

End Sub
                                    '全倉庫マスタの読み込み
Private Function List_Proc()
Dim sts As Integer
Dim com As Integer
    
    List_Proc = False
    
    Combo(0).Clear
    
    com = BtOpGetFirst
    Do
        sts = BTRV(com, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "倉庫マスタ")
                List_Proc = True
                Exit Function
        End Select
        
        Combo(0).AddItem (StrConv(SOKOREC.Soko_No, vbUnicode))
        com = BtOpGetNext
    Loop
    
End Function
                                    '画面初期状態を設定する
Private Sub Clear_Field(Mode As Integer)
Dim i As Integer

    If (Mode = 0) Then
        Combo(0).Text = ""
    End If
    Combo(1).Text = SOKO_BUN0$
    Combo(2).Text = KONS_KBN0$
    Combo(3).Text = "（なし）"
    Combo(4).Text = NAIGAI0$
    Combo(5).Text = KAHI_KBN0$
    
    For i = 0 To 7
        Text(i).Text = ""
    Next i
                
    Combo(2).Enabled = True
    Combo(3).Enabled = True
    Combo(4).Enabled = True
    For i = 1 To 6
        Text(i).Enabled = True
    Next i

    Option1(0).Value = False
    Option1(1).Value = True
    

End Sub

'                                       入力項目のエラーチェック
Private Function Err_Chk() As Integer
            
Dim RetBuf As String
Dim i As Integer


    Err_Chk = False
    If Len(Combo(0).Text) = 0 Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Combo(0).SelStart = 0
        Combo(0).SelLength = Len(Combo(0).Text)
        Combo(0).SetFocus
        Err_Chk = True
        Exit Function
    End If
            
    If Combo(2).Text = KONS_KBN1$ Then
        For i = 0 To UBound(JGYOBU_T)
            If Combo(3).Text = RTrim(JGYOBU_T(i).NAME) Then
                If JGYOBU_T(i).Code = "0" Then
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Combo(2).SetFocus
                    Err_Chk = True
                    Exit Function
                End If
            End If
        Next i
    
        If Combo(4).Text = NAIGAI0$ Then
            Beep
            MsgBox "入力した項目はエラーです。"
            Combo(2).SetFocus
            Err_Chk = True
            Exit Function
        End If
    End If
                    
            
            
            
    For i = 1 To 6
        If Not IsNumeric(Text(i).Text) Then
            Beep
            MsgBox "入力した項目はエラーです。"
            Err_Chk = True
            Text(i).SelStart = 0
            Text(i).SelLength = Len(Text(i).Text)
            Text(i).SetFocus
            Exit Function
        Else
            Text(i).Text = Format(CInt(Text(i).Text), "00")
        End If
    Next i
    
    For i = 1 To 5 Step 2
        If Text(i).Text > Text(i + 1).Text Then
            Beep
            MsgBox "入力した項目はエラーです。"
            Text(i).SelStart = 0
            Text(i).SelLength = Len(Text(i).Text)
            Text(i).SetFocus
            Err_Chk = True
            Exit Function
        End If
    Next i
    
    
    If Text(7).Text = "" Then
        Text(7).Text = "   "
    End If
    
    If Not IsNumeric(Text(7).Text) Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Err_Chk = True
        Text(7).SelStart = 0
        Text(7).SelLength = Len(Text(7).Text)
        Text(7).SetFocus
        Exit Function
    Else
        Text(7).Text = Format(CInt(Text(7).Text), "#0")
    End If


End Function

Private Function Item_Dsp() As Integer
Dim sts As Integer
Dim i As Integer

    Item_Dsp = False
    Call UniCode_Conv(K0_SOKO.Soko_No, Combo(0).Text)
    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    Select Case sts
        Case BtNoErr
            For i = 0 To UBound(JGYOBU_T)
                If JGYOBU_T(i).Code = StrConv(SOKOREC.JGYOBU, vbUnicode) Then
                    Combo(3).Text = RTrim(JGYOBU_T(i).NAME)
                    Exit For
                End If
                                                    '例外処理（ないはず）
                If JGYOBU_T(i).Code = " " Then
                    Combo(3).Text = "（なし）"
                    Exit For
                End If
            Next i
            Select Case StrConv(SOKOREC.SOKO_BUN, vbUnicode)
                Case BUN_JITU$
                    Combo(1).Text = SOKO_BUN0$
                Case BUN_KASO$
                    Combo(1).Text = SOKO_BUN1$
'                Case bun_AUTO$%
'                    Combo(1).Text = SOKO_bun2$
            End Select
            Select Case StrConv(SOKOREC.NAIGAI, vbUnicode)
                Case NAIGAI_NON$
                    Combo(4).Text = NAIGAI0$
                Case NAIGAI_NAI$
                    Combo(4).Text = NAIGAI1$
                Case NAIGAI_GAI$
                    Combo(4).Text = NAIGAI2$
            End Select
            Select Case StrConv(SOKOREC.KAHI_KBN, vbUnicode)
                Case KAHI_KBN_OK$
                    Combo(5).Text = KAHI_KBN0$
                Case KAHI_KBN_NG$
                    Combo(5).Text = KAHI_KBN1$
            End Select
            Select Case StrConv(SOKOREC.KONS_KBN, vbUnicode)
                Case KONS_KBN_OK$
                    Combo(2).Text = KONS_KBN0$
                Case KONS_KBN_NG$
                    Combo(2).Text = KONS_KBN1$
            End Select
        
            Text(0).Text = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
            Text(1).Text = StrConv(SOKOREC.RETU_START, vbUnicode)
            Text(2).Text = StrConv(SOKOREC.RETU_END, vbUnicode)
            Text(3).Text = StrConv(SOKOREC.REN_START, vbUnicode)
            Text(4).Text = StrConv(SOKOREC.REN_END, vbUnicode)
            Text(5).Text = StrConv(SOKOREC.DAN_START, vbUnicode)
            Text(6).Text = StrConv(SOKOREC.DAN_END, vbUnicode)
                        
                        
            If Not IsNumeric(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
                Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
            End If
            
            Text(7).Text = Format(CInt(StrConv(SOKOREC.ORDER_POINT, vbUnicode)), "#0")
                        
            If StrConv(SOKOREC.GOODS_ON_F, vbUnicode) = GOODS_ON Then
                Option1(0).Value = True
                Option1(1).Value = False
            Else
                Option1(0).Value = False
                Option1(1).Value = True
            End If
                        
            If Combo(1).Text = SOKO_BUN1$ Then
                Combo(2).Enabled = False
                Combo(3).Enabled = False
                Combo(4).Enabled = False
                For i = 1 To 6
                    Text(i).Enabled = False
                Next i
            
                Text(7).Enabled = False
            
                Frame1.Enabled = True
            Else
                Combo(2).Enabled = True
                Combo(3).Enabled = True
                Combo(4).Enabled = True
                For i = 1 To 6
                    Text(i).Enabled = True
                Next i
            
            
                Text(7).Enabled = True
            
                Frame1.Enabled = False
            
            
            End If
            
            If Combo(2).Text = KONS_KBN0$ Then
                Combo(3).Enabled = False
                Combo(4).Enabled = False
            Else
                Combo(3).Enabled = True
                Combo(4).Enabled = True
            End If
            
        Case BtErrKeyNotFound
            Call Clear_Field(1)
        Case Else
            Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
            Item_Dsp = True
    End Select

End Function

                                            '倉庫マスタの追加／訂正
Private Function Update_Proc() As Integer
Dim sts As Integer
Dim ans As Integer
Dim com As Integer
Dim i As Integer

    Call UniCode_Conv(K0_SOKO.Soko_No, Combo(0).Text)
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<SOKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Call Clear_Field(0)
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "倉庫マスタ")
                Update_Proc = True
        End Select
    
        DoEvents
    
    Loop
                                            'レコード内容編集
    For i = 0 To UBound(JGYOBU_T)
        If RTrim(JGYOBU_T(i).NAME) = Combo(3).Text Then
            Call UniCode_Conv(SOKOREC.JGYOBU, JGYOBU_T(i).Code)
            Exit For
        End If
                                             '例外処理（ないはず）
        If JGYOBU_T(i).Code = " " Then
            Call UniCode_Conv(SOKOREC.JGYOBU, "0")
            Exit For
        End If
    Next i
    
    
    If i > UBound(JGYOBU_T) Then
        Call UniCode_Conv(SOKOREC.JGYOBU, "0")
    End If
    
    Call UniCode_Conv(SOKOREC.Soko_No, Combo(0).Text)
    Call UniCode_Conv(SOKOREC.SOKO_NAME, Text(0).Text)

    Select Case RTrim(Combo(1).Text)
        Case RTrim(SOKO_BUN0$)
            Call UniCode_Conv(SOKOREC.SOKO_BUN, BUN_JITU$)
        Case RTrim(SOKO_BUN1$)
            Call UniCode_Conv(SOKOREC.SOKO_BUN, BUN_KASO$)
'        Case SOKO_KBN2$
'            Call UniCode_Conv(SOKOREC.SOKO_bun, bun_AUTO$)
    End Select
    
    Call UniCode_Conv(SOKOREC.SOKO_KBN, "")
    
    Select Case Combo(4).Text
        Case NAIGAI0$
            Call UniCode_Conv(SOKOREC.NAIGAI, NAIGAI_NON$)
        Case NAIGAI1$
            Call UniCode_Conv(SOKOREC.NAIGAI, NAIGAI_NAI$)
        Case NAIGAI2$
            Call UniCode_Conv(SOKOREC.NAIGAI, NAIGAI_GAI$)
    End Select
    Select Case Combo(5).Text
        Case KAHI_KBN0$
            Call UniCode_Conv(SOKOREC.KAHI_KBN, KAHI_KBN_OK$)
        Case KAHI_KBN1$
            Call UniCode_Conv(SOKOREC.KAHI_KBN, KAHI_KBN_NG$)
    End Select
    Select Case Combo(2).Text
        Case KONS_KBN0$
            Call UniCode_Conv(SOKOREC.KONS_KBN, KONS_KBN_OK$)
        Case KONS_KBN1$
            Call UniCode_Conv(SOKOREC.KONS_KBN, KONS_KBN_NG$)
    End Select

'    If (StrConv(SOKOREC.SOKO_KBN, vbUnicode) = KBN_KASO$) Then
'        Call UniCode_Conv(SOKOREC.RETU_START, "00")
'        Call UniCode_Conv(SOKOREC.RETU_END, "00")
'        Call UniCode_Conv(SOKOREC.REN_START, "00")
'        Call UniCode_Conv(SOKOREC.REN_END, "00")
'        Call UniCode_Conv(SOKOREC.DAN_START, "00")
'        Call UniCode_Conv(SOKOREC.DAN_END, "00")
'    Else
        Call UniCode_Conv(SOKOREC.RETU_START, Text(1).Text)
        Call UniCode_Conv(SOKOREC.RETU_END, Text(2).Text)
        Call UniCode_Conv(SOKOREC.REN_START, Text(3).Text)
        Call UniCode_Conv(SOKOREC.REN_END, Text(4).Text)
        Call UniCode_Conv(SOKOREC.DAN_START, Text(5).Text)
        Call UniCode_Conv(SOKOREC.DAN_END, Text(6).Text)
        Call UniCode_Conv(SOKOREC.FILLER, "")
'   End If
    
    If (StrConv(SOKOREC.SOKO_KBN, vbUnicode) = BUN_KASO$) Then
        Call UniCode_Conv(SOKOREC.ORDER_POINT, "")
    Else
        Call UniCode_Conv(SOKOREC.ORDER_POINT, Format(CInt(Text(7).Text), "000"))
    End If


    If Option1(0).Value = True Then
        Call UniCode_Conv(SOKOREC.GOODS_ON_F, GOODS_ON)
    Else
        Call UniCode_Conv(SOKOREC.GOODS_ON_F, GOODS_OFF)
    End If

    Do
        sts = BTRV(com, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<SOKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    Call Clear_Field(0)
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, com, "倉庫マスタ")
                Update_Proc = True
                Exit Function
        End Select
    Loop
                                        'リストボックス追加
    If com = BtOpInsert Then
        Combo(0).AddItem Combo(0).Text
    End If
                                        '画面クリアー
    Call Clear_Field(0)
End Function
                                            '倉庫マスタの削除
Private Function Delete_Proc() As Integer
Dim sts As Integer
Dim ans As Integer
Dim flg As Boolean
Dim i As Integer

    Call UniCode_Conv(K0_SOKO.Soko_No, Combo(0).Text)
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
                flg = True
                Exit Do
            Case BtErrKeyNotFound
                flg = False
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<SOKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Call Clear_Field(0)
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "倉庫マスタ")
                Delete_Proc = True
        End Select
    Loop

    If flg Then
        Do
            sts = BTRV(BtOpDelete, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<SOKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                        Call Clear_Field(0)
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpDelete, "倉庫マスタ")
                    Delete_Proc = True
            End Select
        Loop
    End If
                                        'リストボックス削除
    For i = 0 To Combo(0).ListCount - 1
        If Combo(0).Text = Combo(0).List(i) Then
            Combo(0).RemoveItem i
            Exit For
        End If
    Next i
                                        '画面クリアー
    Call Clear_Field(0)
End Function

Private Sub Combo_DblClick(Index As Integer)

    If (Index = 0) Then
        If Item_Dsp() Then
            Unload Me
        End If
                
        Text(0).SetFocus
    End If

End Sub

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i As Integer
Dim wk As String

    Select Case KeyCode
        Case vbKeyReturn
            Select Case Index
                Case 0
                    wk = Trim(Combo(Index).Text)
                    If Len(wk) <> 2 Then
                        Beep
                        MsgBox "倉庫№は２桁入力です。"
                        Combo(Index).SetFocus
                        Exit Sub
                    End If
                    If Item_Dsp() Then
                        Unload Me
                    End If
                        
                Case 1
                    If Combo(Index).Text = SOKO_BUN1$ Then
                        Combo(2).Text = KONS_KBN0$
'                        For i = 0 To UBound(JGYOBU_T)
'                            If RTrim(JGYOBU_T(i).CODE) = "0" Then
'                                Combo(3).Text = RTrim(JGYOBU_T(i).NAME)
'                                Exit For
'                            End If
'                        Next i
                        Combo(3).Text = "（なし）"
                        Combo(4).Text = NAIGAI0$
                        Combo(2).Enabled = False
                        Combo(3).Enabled = False
                        Combo(4).Enabled = False
                        For i = 1 To 6
                            Text(i).Text = "01"
                            Text(i).Enabled = False
                        Next i
                    
                    
                        Text(7).Text = "000"
                        Text(7).Enabled = False
                        
                        Frame1.Enabled = True
                    Else
                        Combo(2).Enabled = True
                        For i = 1 To 6
                            Text(i).Enabled = True
                        Next i
                    
                        Text(7).Enabled = True
                        Frame1.Enabled = False
                    End If
                
                Case 2
                
                    If Combo(Index).Text = KONS_KBN0$ Then
'                        For i = 0 To UBound(JGYOBU_T)
'                            If RTrim(JGYOBU_T(i).CODE) = "0" Then
'                                Combo(3).Text = RTrim(JGYOBU_T(i).NAME)
'                                Exit For
'                            End If
'                        Next i
                        Combo(3).Text = "（なし）"
                        Combo(4).Text = NAIGAI0$
                        Combo(3).Enabled = False
                        Combo(4).Enabled = False
                    Else
                        Combo(3).Enabled = True
                        Combo(4).Enabled = True
                
                    End If
            End Select
            If Index = 4 Then
                Text(0).SetFocus
            Else
                For i = Index + 1 To 4
                    If Combo(i).Enabled Then
                        Combo(i).SetFocus
                        Exit For
                    End If
                Next i
                If i > 4 Then
                    Text(0).SetFocus
                End If
            End If
    End Select
End Sub


Private Sub Combo_LostFocus(Index As Integer)

'    If (Index = 0) Then
'        If Item_Dsp() Then
'            Unload Me
'        End If
                
'        Text(0).SelStart = ZERO
'        Text(0).SelLength = Len(RTrim(Text(0).Text))
'        Text(0).SetFocus
'    End If


End Sub

Private Sub Command_Click(Index As Integer)
Dim yn As Integer
Dim sts As Integer

    Select Case Index
        Case 0
                                            'エラーチェック
            If Err_Chk() Then
                Exit Sub
            End If
            
            Beep
            yn = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            End If
            Combo(0).SetFocus
        Case 3
            Beep
            yn = MsgBox("削除しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                If Delete_Proc() Then
                   Unload Me
                End If
            End If
            Combo(0).SetFocus
        Case 8
            Beep
            yn = MsgBox("データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                If Data_Proc() Then
                    Unload Me
                End If
            End If
        Case 11
            Unload Me
        Case Else
            Beep
    End Select
End Sub


Private Sub Form_DblClick()
    PrintForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
Dim sts As Integer

    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
    
    Text_Max = 7                '画面項目別最大ｲﾝﾃﾞｯｸｽ
    Combo_Max = 5
    Command_Max = 11

    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                'ＣＳＶファイル名取り込み
    If GetIni("FILE", "SOKO_CSV", "SYS", c) Then
        Beep
        MsgBox "倉庫マスタデータ出力用ファイル[SOKO_CSV]の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    Soko_Csv = Trim(c)
                                
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

    Combo(3).AddItem "（なし）"
    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).Code = " " Then
            Exit For
        End If
        Combo(3).AddItem RTrim(JGYOBU_T(i).NAME)
    Next i

                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '倉庫分類設定
    Combo(1).AddItem SOKO_BUN0$
    Combo(1).AddItem SOKO_BUN1$
                                '混載可否
    Combo(2).AddItem KONS_KBN0$
    Combo(2).AddItem KONS_KBN1$
                                '国内外設定
    Combo(4).AddItem NAIGAI0$
    Combo(4).AddItem NAIGAI1$
    Combo(4).AddItem NAIGAI2$
                                '使用可否
    Combo(5).AddItem KAHI_KBN0$
    Combo(5).AddItem KAHI_KBN1$
    
    If List_Proc() Then
        Unload Me
    End If
    Beep
    MsgBox "この処理はシステム全体に影響するので十分注意して操作してください。"
                                '画面初期設定
    Call Clear_Field(0)
    
    Combo(0).SetFocus
    
    End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫マスタ")
        End If
    End If
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "倉庫マスタ")
    End If
    Set F1010101 = Nothing

    End
End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If

End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim RetBuf As String
Dim i As Integer


    Select Case KeyCode
        Case vbKeyReturn
            If (Index > 0) Then
                If Not IsNumeric(Text(Index).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Exit Sub
                Else
                    Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                End If
            End If
            If Index < Text_Max Then
                For i = Index + 1 To Text_Max
                    If Text(i).Enabled Then
                        Text(i).SetFocus
                        Exit For
                    End If
                Next i
            End If
    End Select
End Sub
