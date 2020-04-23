VERSION 5.00
Begin VB.Form PM000101 
   Caption         =   "管理マスタメンテナンス"
   ClientHeight    =   8250
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   13155
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
   ScaleHeight     =   8250
   ScaleWidth      =   13155
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   11
      Left            =   8295
      MaxLength       =   1
      TabIndex        =   11
      Top             =   2160
      Width           =   240
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   18
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   18
      Top             =   6240
      Width           =   240
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   16
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   16
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   15
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   15
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   4  '全角ひらがな
      Index           =   14
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   14
      Top             =   3840
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   4  '全角ひらがな
      Index           =   13
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   13
      Top             =   3240
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   12
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   12
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   17
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   17
      Top             =   5760
      Width           =   240
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   10
      Left            =   8295
      MaxLength       =   4
      TabIndex        =   10
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   9
      Left            =   6135
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2160
      Width           =   240
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   6135
      MaxLength       =   4
      TabIndex        =   8
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   8055
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   7215
      MaxLength       =   2
      TabIndex        =   6
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   6135
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   4
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   2280
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1320
      Width           =   1050
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   4  '全角ひらがな
      Index           =   0
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終 了"
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
      TabIndex        =   30
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Index           =   8
      Left            =   7800
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "請 求"
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
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Index           =   3
      Left            =   2640
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "更 新"
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
      TabIndex        =   19
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      Caption         =   "仕入金額　丸め"
      Height          =   255
      Index           =   24
      Left            =   420
      TabIndex        =   55
      Top             =   6360
      Width           =   1710
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      Caption         =   "売上金額　丸め"
      Height          =   255
      Index           =   22
      Left            =   420
      TabIndex        =   54
      Top             =   5880
      Width           =   1710
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      Caption         =   "FAX番号"
      Height          =   255
      Index           =   21
      Left            =   1200
      TabIndex        =   53
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      Caption         =   "電話番号"
      Height          =   255
      Index           =   20
      Left            =   1200
      TabIndex        =   52
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      Caption         =   "センター名"
      Height          =   255
      Index           =   19
      Left            =   840
      TabIndex        =   51
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      Caption         =   "会社名"
      Height          =   255
      Index           =   18
      Left            =   1440
      TabIndex        =   50
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label LblS_TANTO 
      Alignment       =   1  '右揃え
      Height          =   255
      Left            =   3120
      TabIndex        =   49
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      Caption         =   "承認者"
      Height          =   255
      Index           =   17
      Left            =   1440
      TabIndex        =   48
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "※丸め:(0:切捨て 5:四捨五入 9:切上げ)"
      Height          =   255
      Index           =   16
      Left            =   6720
      TabIndex        =   47
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label Label 
      Caption         =   "    丸め"
      Height          =   255
      Index           =   15
      Left            =   7215
      TabIndex        =   46
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "%"
      Height          =   255
      Index           =   14
      Left            =   9015
      TabIndex        =   45
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Label 
      Caption         =   "新　税率"
      Height          =   255
      Index           =   13
      Left            =   7215
      TabIndex        =   44
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "    丸め"
      Height          =   255
      Index           =   12
      Left            =   5055
      TabIndex        =   43
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "%"
      Height          =   255
      Index           =   11
      Left            =   6855
      TabIndex        =   42
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Label 
      Caption         =   "現　税率"
      Height          =   255
      Index           =   10
      Left            =   5055
      TabIndex        =   41
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "日"
      Height          =   255
      Index           =   9
      Left            =   8535
      TabIndex        =   40
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label 
      Caption         =   "月"
      Height          =   255
      Index           =   8
      Left            =   7695
      TabIndex        =   39
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label 
      Caption         =   "年"
      Height          =   255
      Index           =   7
      Left            =   6855
      TabIndex        =   38
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label 
      Caption         =   "消費税変更日付"
      Height          =   255
      Index           =   6
      Left            =   4335
      TabIndex        =   37
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      Caption         =   "売上№"
      Height          =   255
      Index           =   5
      Left            =   1200
      TabIndex        =   36
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      Caption         =   "注文№"
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   35
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "指図票№"
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   34
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "(末日=31)"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   33
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "月末締日"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   32
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "ﾚｺｰﾄﾞ№"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   31
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PM000101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const ptxRec_No% = 0                'ﾚｺｰﾄﾞ№(入力不可)
Private Const ptxSHIME_DD% = 1              '月末締め処理

Private Const ptxSASHIZU_NO% = 2            '指図書№
Private Const ptxORDER_NO% = 3              '発注№
Private Const ptxURIAGE_NO% = 4             '資材売上ﾚｺｰﾄﾞ№

Private Const ptxZEI_CHANGE_YY% = 5         '消費税変更日付 年
Private Const ptxZEI_CHANGE_MM% = 6         '消費税変更日付 月
Private Const ptxZEI_CHANGE_DD% = 7         '消費税変更日付 日

Private Const ptxNOW_ZEI_RITU% = 8          '新　消費税率
Private Const ptxNOW_MARUME% = 9            '　　丸め

Private Const ptxNEW_ZEI_RITU% = 10         '新　消費税率
Private Const ptxNEW_MARUME% = 11           '　　丸め

Private Const ptxSHONIN_CODE% = 12          '承認担当
Private Const ptxKAISHA_NAME% = 13          '会社名
Private Const ptxCENTER_NAME% = 14          'センター名
Private Const ptxTEL_NO% = 15               '電話番号
Private Const ptxFAX_NO% = 16               'FAX番号

Private Const ptxURI_MARUME% = 17           '売上金額　丸め
Private Const ptxSHI_MARUME% = 18           '仕入金額　丸め







Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PM000101.MousePointer = vbHourglass

    Call Ctrl_Lock(PM000101)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM000101)


    PM000101.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts As Integer
    
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        Case ptxSHIME_DD    '月末締め日
            If Not IsNumeric(Text1(ptxSHIME_DD).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxSHIME_DD).SetFocus
                Exit Function
            End If
    
            If CInt(Text1(ptxSHIME_DD).Text) < 1 Or CInt(Text1(ptxSHIME_DD).Text) > 31 Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxSHIME_DD).SetFocus
                Exit Function
            End If
    
            Text1(ptxSHIME_DD).Text = Format(CInt(Text1(ptxSHIME_DD).Text), "00")
    
        Case ptxSASHIZU_NO  '指図書№
            
            If Not IsNumeric(Text1(ptxSASHIZU_NO).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxSASHIZU_NO).SetFocus
                Exit Function
            End If
    
'            If CLng(Text1(ptxSASHIZU_NO).Text) < 0 Or CLng(Text1(ptxSASHIZU_NO).Text) > 99999 Then     '2007.11.28
            If CLng(Text1(ptxSASHIZU_NO).Text) < 0 Or CLng(Text1(ptxSASHIZU_NO).Text) > 99999999 Then   '2007.11.28
                MsgBox "入力した項目はエラーです。"
                Text1(ptxSASHIZU_NO).SetFocus
                Exit Function
            End If
            
'            Text1(ptxSASHIZU_NO).Text = Format(CLng(Text1(ptxSASHIZU_NO).Text), "00000")       '2007.11.28
            Text1(ptxSASHIZU_NO).Text = Format(CLng(Text1(ptxSASHIZU_NO).Text), "00000000")     '2007.11.28
    
        Case ptxORDER_NO    '発注№
            
            If Not IsNumeric(Text1(ptxORDER_NO).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxORDER_NO).SetFocus
                Exit Function
            End If
    
            If CLng(Text1(ptxORDER_NO).Text) < 0 Or CLng(Text1(ptxORDER_NO).Text) > 99999 Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxORDER_NO).SetFocus
                Exit Function
            End If
            
            Text1(ptxORDER_NO).Text = Format(CLng(Text1(ptxORDER_NO).Text), "00000")
    
        Case ptxURIAGE_NO   '資材売上ﾚｺｰﾄﾞ№
            
            If Not IsNumeric(Text1(ptxURIAGE_NO).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxURIAGE_NO).SetFocus
                Exit Function
            End If
    
            If CLng(Text1(ptxURIAGE_NO).Text) < 0 Or CLng(Text1(ptxURIAGE_NO).Text) > 99999 Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxURIAGE_NO).SetFocus
                Exit Function
            End If
            
            Text1(ptxURIAGE_NO).Text = Format(CLng(Text1(ptxURIAGE_NO).Text), "00000")
    
        Case ptxZEI_CHANGE_YY '消費税変更日付　年
            If Not IsNumeric(Text1(ptxZEI_CHANGE_YY).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxZEI_CHANGE_YY).SetFocus
                Exit Function
            End If
    
            If CInt(Text1(ptxZEI_CHANGE_YY).Text) < 1 Or CInt(Text1(ptxZEI_CHANGE_YY).Text) > 9999 Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxZEI_CHANGE_YY).SetFocus
                Exit Function
            End If
            
            Text1(ptxZEI_CHANGE_YY).Text = Format(CInt(Text1(ptxZEI_CHANGE_YY).Text), "0000")
    
        Case ptxZEI_CHANGE_MM '消費税変更日付　月
            If Not IsNumeric(Text1(ptxZEI_CHANGE_MM).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxZEI_CHANGE_MM).SetFocus
                Exit Function
            End If
    
            If CInt(Text1(ptxZEI_CHANGE_MM).Text) < 1 Or CInt(Text1(ptxZEI_CHANGE_MM).Text) > 12 Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxZEI_CHANGE_MM).SetFocus
                Exit Function
            End If
            
            Text1(ptxZEI_CHANGE_MM).Text = Format(CInt(Text1(ptxZEI_CHANGE_MM).Text), "00")
    
        Case ptxZEI_CHANGE_DD   '消費税変更日付　日
            If Not IsNumeric(Text1(ptxZEI_CHANGE_DD).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxZEI_CHANGE_DD).SetFocus
                Exit Function
            End If
    
            If CInt(Text1(ptxZEI_CHANGE_DD).Text) < 1 Or CInt(Text1(ptxZEI_CHANGE_DD).Text) > 31 Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxZEI_CHANGE_DD).SetFocus
                Exit Function
            End If
            
            Text1(ptxZEI_CHANGE_DD).Text = Format(CInt(Text1(ptxZEI_CHANGE_DD).Text), "00")
            '日付OK？
            If Not IsDate(Text1(ptxZEI_CHANGE_YY).Text & "/" & Text1(ptxZEI_CHANGE_MM).Text & "/" & Text1(ptxZEI_CHANGE_DD).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxZEI_CHANGE_YY).SetFocus
                Exit Function
            End If
    
        Case ptxNOW_ZEI_RITU    '現　税率
            If Not IsNumeric(Text1(ptxNOW_ZEI_RITU).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxNOW_ZEI_RITU).SetFocus
                Exit Function
            End If
    
            If CDbl(Text1(ptxNOW_ZEI_RITU).Text) < 0 Or CDbl(Text1(ptxNOW_ZEI_RITU).Text) > 99.9 Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxNOW_ZEI_RITU).SetFocus
                Exit Function
            End If
            
            Text1(ptxNOW_ZEI_RITU).Text = Format(CDbl(Text1(ptxNOW_ZEI_RITU).Text), "#0.0")
    
        Case ptxNOW_MARUME      '現　丸め
            If Not IsNumeric(Text1(ptxNOW_MARUME).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxNOW_MARUME).SetFocus
                Exit Function
            End If
    
            If CInt(Text1(ptxNOW_MARUME).Text) <> 0 And _
                CInt(Text1(ptxNOW_MARUME).Text) <> 5 And _
                CInt(Text1(ptxNOW_MARUME).Text) <> 9 Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxNOW_MARUME).SetFocus
                Exit Function
            End If
            
        Case ptxNEW_ZEI_RITU    '新　税率
            If Not IsNumeric(Text1(ptxNEW_ZEI_RITU).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxNOW_ZEI_RITU).SetFocus
                Exit Function
            End If
    
            If CDbl(Text1(ptxNEW_ZEI_RITU).Text) < 0 Or CDbl(Text1(ptxNEW_ZEI_RITU).Text) > 99.9 Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxNOW_ZEI_RITU).SetFocus
                Exit Function
            End If
            
            Text1(ptxNEW_ZEI_RITU).Text = Format(CDbl(Text1(ptxNEW_ZEI_RITU).Text), "#0.0")
    
        Case ptxNEW_MARUME      '現　丸め
            If Not IsNumeric(Text1(ptxNEW_MARUME).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxNEW_MARUME).SetFocus
                Exit Function
            End If
    
            If CInt(Text1(ptxNEW_MARUME).Text) <> 0 And _
                CInt(Text1(ptxNEW_MARUME).Text) <> 5 And _
                CInt(Text1(ptxNEW_MARUME).Text) <> 9 Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxNEW_MARUME).SetFocus
                Exit Function
            End If
    
        Case ptxSHONIN_CODE     '承認担当者ｺｰﾄﾞ
            If Trim(Text1(ptxSHONIN_CODE).Text) = "" Then
                LblS_TANTO.Caption = ""
            Else
                Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxSHONIN_CODE).Text)
            
                sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
                Select Case sts
                    Case BtNoErr
                        LblS_TANTO.Caption = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                    Case BtErrKeyNotFound
                        LblS_TANTO.Caption = ""
                        MsgBox "入力した項目はエラーです。"
                        Text1(ptxSHONIN_CODE).SetFocus
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                        Exit Function
                End Select
            
            
            End If
    
    
        Case ptxKAISHA_NAME     '会社名
        Case ptxCENTER_NAME     '会社名
        Case ptxTEL_NO          '電話番号
        Case ptxFAX_NO          'FAX番号
    
        Case ptxURI_MARUME      '売上金額　丸め
            If Not IsNumeric(Text1(ptxURI_MARUME).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxURI_MARUME).SetFocus
                Exit Function
            End If
    
            If CInt(Text1(ptxURI_MARUME).Text) <> 0 And _
                CInt(Text1(ptxURI_MARUME).Text) <> 5 And _
                CInt(Text1(ptxURI_MARUME).Text) <> 9 Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxURI_MARUME).SetFocus
                Exit Function
            End If
    
        Case ptxSHI_MARUME      '仕入金額　丸め
            If Not IsNumeric(Text1(ptxSHI_MARUME).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxSHI_MARUME).SetFocus
                Exit Function
            End If
    
            If CInt(Text1(ptxSHI_MARUME).Text) <> 0 And _
                CInt(Text1(ptxSHI_MARUME).Text) <> 5 And _
                CInt(Text1(ptxSHI_MARUME).Text) <> 9 Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxSHI_MARUME).SetFocus
                Exit Function
            End If
    
    
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function
Private Function Item_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   画面表示
'----------------------------------------------------------------------------
Dim sts As Integer
Dim i   As Integer

    Item_Disp_Proc = True
    
    '管理ﾏｽﾀ（KEY=0）読み込み
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
    
    
    
    Text1(ptxRec_No).Text = P_ST_KANRI_No   'ﾚｺｰﾄﾞ№
    
    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
            'ﾚｺｰﾄﾞ内容の表示
                                            '月末締め日
            Text1(ptxSHIME_DD).Text = StrConv(P_KANRIREC.SHIME_DD, vbUnicode)
                                            '指図書№
            Text1(ptxSASHIZU_NO).Text = StrConv(P_KANRIREC.SASHIZU_NO, vbUnicode)
                                            '発注№
            Text1(ptxORDER_NO).Text = StrConv(P_KANRIREC.ORDER_NO, vbUnicode)
                                            '資材売上ﾚｺｰﾄﾞ№
            Text1(ptxURIAGE_NO).Text = StrConv(P_KANRIREC.URIAGE_NO, vbUnicode)
                                            '消費税変更日付 年
            Text1(ptxZEI_CHANGE_YY).Text = Left(StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode), 4)
                                            '消費税変更日付 月
            Text1(ptxZEI_CHANGE_MM).Text = Mid(StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode), 5, 2)
                                            '消費税変更日付 日
            Text1(ptxZEI_CHANGE_DD).Text = Right(StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode), 2)
                                            '現　消費税率
            Text1(ptxNOW_ZEI_RITU).Text = Format(CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)), "#0.0")
                                            '現　丸め
            Text1(ptxNOW_MARUME).Text = StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)
                                            '新　消費税率
            Text1(ptxNEW_ZEI_RITU).Text = Format(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)), "#0.0")
                                            '新　丸め
            Text1(ptxNEW_MARUME).Text = StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)
                                            '承認担当者
            Text1(ptxSHONIN_CODE).Text = Trim(StrConv(P_KANRIREC.SHONIN_CODE, vbUnicode))
        
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxSHONIN_CODE).Text)
        
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    LblS_TANTO.Caption = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    LblS_TANTO.Caption = ""
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Exit Function
            End Select
                                            '会社名
            Text1(ptxKAISHA_NAME).Text = Trim(StrConv(P_KANRIREC.KAISHA_NAME, vbUnicode))
                                            'センター名
            Text1(ptxCENTER_NAME).Text = Trim(StrConv(P_KANRIREC.CENTER_NAME, vbUnicode))
                                            '電話番号
            Text1(ptxTEL_NO).Text = Trim(StrConv(P_KANRIREC.TEL_NO, vbUnicode))
                                            'FAX番号
            Text1(ptxFAX_NO).Text = Trim(StrConv(P_KANRIREC.FAX_NO, vbUnicode))
                                            '売上金額　丸め
            Text1(ptxURI_MARUME).Text = StrConv(P_KANRIREC.URI_MARUME, vbUnicode)
                                            '仕入金額　丸め
            Text1(ptxSHI_MARUME).Text = StrConv(P_KANRIREC.SHI_MARUME, vbUnicode)
        
        
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "管理マスタ")
            Exit Function
    End Select

    Item_Disp_Proc = False

End Function

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   管理マスタ出力
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim com     As Integer


    Update_Proc = True
    '管理ﾏｽﾀ（KEY=0）読み込み
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Update_Proc = True
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "管理マスタ")
                Exit Function
        
        End Select
    
    
    Loop
    '--------------------------------------------------------レコード内容編集
    
    Call UniCode_Conv(P_KANRIREC.REC_NO, P_ST_KANRI_No)                                 'ﾚｺｰﾄﾞ№
    Call UniCode_Conv(P_KANRIREC.SHIME_DD, Text1(ptxSHIME_DD).Text)                     '月末締め日
    
    Call UniCode_Conv(P_KANRIREC.xSASHIZU_NO, "")                                       '指図書№2007.11.28
    
    Call UniCode_Conv(P_KANRIREC.SASHIZU_NO, Text1(ptxSASHIZU_NO).Text)                 '指図書№
    Call UniCode_Conv(P_KANRIREC.ORDER_NO, Text1(ptxORDER_NO).Text)                     '注文№
    Call UniCode_Conv(P_KANRIREC.URIAGE_NO, Text1(ptxURIAGE_NO).Text)                   '資材売上ﾚｺｰﾄﾞ№
    Call UniCode_Conv(P_KANRIREC.ZEI_CHANGE_YMD, Text1(ptxZEI_CHANGE_YY).Text & _
                                            Text1(ptxZEI_CHANGE_MM).Text & _
                                            Text1(ptxZEI_CHANGE_DD).Text)               '消費税変更日付
    Call UniCode_Conv(P_KANRIREC.NOW_ZEI_RITU, Format(CDbl(Text1(ptxNOW_ZEI_RITU).Text), "00.0"))   '現　税率
    Call UniCode_Conv(P_KANRIREC.NOW_MARUME, Text1(ptxNOW_MARUME).Text)                             '現　まるめ
    Call UniCode_Conv(P_KANRIREC.NEW_ZEI_RITU, Format(CDbl(Text1(ptxNEW_ZEI_RITU).Text), "00.0"))   '新　税率
    Call UniCode_Conv(P_KANRIREC.NEW_MARUME, Text1(ptxNEW_MARUME).Text)                             '新　まるめ
    Call UniCode_Conv(P_KANRIREC.SHONIN_CODE, Text1(ptxSHONIN_CODE).Text)               '承認担当者
    Call UniCode_Conv(P_KANRIREC.KAISHA_NAME, Text1(ptxKAISHA_NAME).Text)               '会社名
    Call UniCode_Conv(P_KANRIREC.CENTER_NAME, Text1(ptxCENTER_NAME).Text)               'センター名
    Call UniCode_Conv(P_KANRIREC.TEL_NO, Text1(ptxTEL_NO).Text)                         '電話番号
    Call UniCode_Conv(P_KANRIREC.FAX_NO, Text1(ptxFAX_NO).Text)                         'FAX番号
    
    Call UniCode_Conv(P_KANRIREC.URI_MARUME, Text1(ptxURI_MARUME).Text)                 '売上金額　丸め
    Call UniCode_Conv(P_KANRIREC.SHI_MARUME, Text1(ptxSHI_MARUME).Text)                 '仕入金額　丸め
        
    
    Call UniCode_Conv(P_KANRIREC.FILLER, "")                                            'Filler
    
    
    Do
        
        DoEvents
        
        sts = BTRV(com, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, com, "管理マスタ")
                Exit Function
        End Select
    Loop

    Update_Proc = False


End Function


Private Sub Command1_Click(Index As Integer)

Dim ans As Integer
Dim i   As Integer

    Select Case Index
        Case P_CMD_Upd        '更新
            
            
            For i = ptxRec_No To ptxSHONIN_CODE
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
            
            Beep
            ans = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            End If
            Text1(ptxSASHIZU_NO).SetFocus
        Case P_CMD_DEL                      '削除
        Case P_CMD_DSP                      '検索/表示
        
        Case 5                              '請求関連項目へ 2008.02.13
        
            PM000102.Show vbModal
        
        
        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        Case P_CMD_PRT                      '印刷
        
        Case P_CMD_End                      '終了
            Unload Me
    End Select

End Sub

Private Sub Form_DblClick()
'    PrintForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command1(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

Dim c       As String * 128

    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If

                                'ログファイル名取り込み
'    If GetIni("FILE", "LOGF", "SYS", c) Then
    If GetIni(App.EXEName, "LOGF", App.EXEName, c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                
    PM000101.Caption = PM000101.Caption & LAST_UPDATE_DAY
                                
                                
                                '管理マスタＯＰＥＮ
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    Show
                                
                                '画面初期設定
    If Item_Disp_Proc() Then
        Unload Me
    End If
                                
    Text1(ptxSHIME_DD).SetFocus
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            '管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "管理マスタ")
        End If
    End If
    
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PM000101 = Nothing

    End
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    If Error_Check_Proc(Index) Then     'エラーチェック
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動
End Sub

