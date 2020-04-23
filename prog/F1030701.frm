VERSION 5.00
Begin VB.Form F1030701 
   BackColor       =   &H00FFFFFF&
   Caption         =   "出荷予定登録"
   ClientHeight    =   6015
   ClientLeft      =   2130
   ClientTop       =   2715
   ClientWidth     =   13455
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
   MaxButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   13455
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Height          =   360
      Index           =   15
      Left            =   2400
      MaxLength       =   8
      TabIndex        =   51
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   840
      MaxLength       =   12
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      Left            =   2640
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   1320
      Width           =   972
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   2
      Left            =   3600
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   17
      Top             =   4080
      Width           =   4950
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   14
      Left            =   8640
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   13
      Left            =   7920
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   12
      Left            =   7200
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   11
      Left            =   6480
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   10
      Left            =   5040
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   9
      Left            =   3360
      MaxLength       =   5
      TabIndex        =   11
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   2400
      MaxLength       =   5
      TabIndex        =   10
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   5040
      MaxLength       =   6
      TabIndex        =   9
      Top             =   2280
      Width           =   852
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   3240
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   1560
      MaxLength       =   4
      TabIndex        =   6
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   9840
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   6360
      MaxLength       =   25
      TabIndex        =   4
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   3720
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1320
      Width           =   2535
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   960
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   1320
      Width           =   972
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
      Left            =   10200
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5040
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
      Left            =   9360
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5040
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
      Left            =   8520
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5040
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
      Index           =   8
      Left            =   7680
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5040
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
      Left            =   6600
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5040
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
      Left            =   5760
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5040
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
      Left            =   4920
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5040
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
      Left            =   4080
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5040
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
      Index           =   3
      Left            =   3000
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5040
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
      Left            =   2160
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5040
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
      Left            =   1320
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5040
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
      Left            =   480
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "（未入力：自動発番）"
      Height          =   255
      Index           =   7
      Left            =   6120
      TabIndex        =   50
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "（未入力：自動発番）"
      Height          =   255
      Index           =   19
      Left            =   2520
      TabIndex        =   49
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IDNo"
      Height          =   255
      Index           =   18
      Left            =   240
      TabIndex        =   48
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   255
      Index           =   15
      Left            =   2640
      TabIndex        =   47
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   46
      Top             =   5400
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "向け先"
      Height          =   255
      Index           =   17
      Left            =   1560
      TabIndex        =   45
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   16
      Left            =   8400
      TabIndex        =   44
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   13
      Left            =   7680
      TabIndex        =   43
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   12
      Left            =   6960
      TabIndex        =   42
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ホスト棚番"
      Height          =   255
      Index           =   11
      Left            =   6480
      TabIndex        =   41
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ホスト倉庫"
      Height          =   255
      Index           =   10
      Left            =   5040
      TabIndex        =   40
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "（先）"
      Height          =   255
      Index           =   9
      Left            =   3360
      TabIndex        =   39
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "予算単位（元）"
      Height          =   255
      Index           =   8
      Left            =   1560
      TabIndex        =   38
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "伝票№"
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   37
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   36
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   35
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "伝票日付"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   34
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "出荷予定数"
      Height          =   255
      Index           =   0
      Left            =   9840
      TabIndex        =   33
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "注文区分"
      Height          =   255
      Index           =   14
      Left            =   960
      TabIndex        =   32
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品  名"
      Height          =   255
      Index           =   6
      Left            =   6720
      TabIndex        =   31
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品  番"
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   30
      Top             =   1080
      Width           =   735
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1030701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Const pcmbC_Kbn% = 0
Private Const pcmbNAIGAI% = 1
Private Const pcmbMUKE_CODE% = 2

Private Const ptxMAX% = 15

Private Const ptxID_No% = 0
Private Const ptxCode% = 1
Private Const ptxName% = 2
Private Const ptxS_Qty% = 3
Private Const ptxYY% = 4
Private Const ptxMM% = 5
Private Const ptxDD% = 6
Private Const ptxNo% = 7
Private Const ptxMoto% = 8
Private Const ptxSaki% = 9
Private Const ptxSoko% = 10
Private Const ptxS_No% = 11
Private Const ptxRetu% = 12
Private Const ptxRen% = 13
Private Const ptxDan% = 14
Private Const ptxMUKE_CODE% = 15         '向け先（コード入力用）
                                   
'Private Const LAST_UPDATE_DAY$ = "[F103070]2018.04.21 09:00"
'Private Const LAST_UPDATE_DAY$ = "[F103070]2018.04.27 16:45"
Private Const LAST_UPDATE_DAY$ = "[F103070]2020.04.14 14:00 更新後前回 表示残りを修正"
                                   
Private DEF_CYU_KBN As String * 1       '2009.04.14
Private OSAKA_MODE As String * 1        '2010.03.23
                                   
                                   
                                   
                                   '画面初期状態を設定する
Private Sub Clear_Field(Optional Start_Pos As Integer = 0)
Dim i As Integer
    
    For i = Start_Pos To ptxMAX
        Text(i).Text = ""
    Next i
    
End Sub
                                    '品目マスタより各項目を表示する
Private Function Item_Dsp() As Integer

Dim sts As Integer


    Item_Dsp = True
                                                '国内外チェック
                                                'まず外部品番で読み込み
    
    Text(ptxCode).Text = StrConv(Text(ptxCode).Text, vbUpperCase)
    
    
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(ptxCode))
        
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            
            Text(ptxName) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
            Text(ptxSoko) = StrConv(ITEMREC.BIKOU_SOKO, vbUnicode)
            Text(ptxS_No) = StrConv(ITEMREC.ST_SOKO, vbUnicode)
            Text(ptxRetu) = StrConv(ITEMREC.ST_RETU, vbUnicode)
            Text(ptxRen) = StrConv(ITEMREC.ST_REN, vbUnicode)
            Text(ptxDan) = StrConv(ITEMREC.ST_DAN, vbUnicode)
        
        Case BtErrKeyNotFound
                                                '内部品番で読み込み
            Call UniCode_Conv(K2_ITEM.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K2_ITEM.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
            Call UniCode_Conv(K2_ITEM.HIN_NAI, Text(ptxCode).Text)
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
            Select Case sts
                Case BtNoErr
                    Text(ptxCode).Text = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                    Text(ptxName).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                    Text(ptxSoko).Text = StrConv(ITEMREC.BIKOU_SOKO, vbUnicode)
                    Text(ptxS_No).Text = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    Text(ptxRetu).Text = StrConv(ITEMREC.ST_RETU, vbUnicode)
                    Text(ptxRen).Text = StrConv(ITEMREC.ST_REN, vbUnicode)
                    Text(ptxDan).Text = StrConv(ITEMREC.ST_DAN, vbUnicode)
                Case BtErrKeyNotFound
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Item_Dsp = SYS_ERR
                    Exit Function
            End Select
                
                
                
        Case Else
                
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Item_Dsp = SYS_ERR
            Exit Function
        
    End Select
    
    Item_Dsp = False
    
End Function

'                                       入力項目のエラーチェック
Private Function Err_Chk() As Integer

Dim yn          As Integer
Dim RetBuf      As String
Dim i           As Integer
Dim sts         As Integer
Dim Flg         As Boolean
Dim Qty         As Long
Dim W_CYU_KBN   As String

    Err_Chk = True
                                        '品目チェック
    
    If Trim(Text(ptxCode).Text) = "" Then   '2018.04.20
        Beep                                '2018.04.20
        MsgBox "品番は必須入力です。 "      '2018.04.20
        Text(ptxS_Qty).SetFocus             '2018.04.20
        Exit Function                       '2018.04.20
    End If                                  '2018.04.20
    
    
    
    
    sts = Item_Dsp
    Select Case sts
        Case False
        Case True
            Beep
            MsgBox "入力した項目はエラーです｡ "
            Text(ptxCode).SetFocus
            Exit Function
        Case Else
            Err_Chk = SYS_ERR
            Exit Function
    End Select
                                        '出荷予定数量チェック
    If Trim(Text(ptxS_Qty).Text) = "" Then
        Text(ptxS_Qty) = "0"
    End If
        
    If Not IsNumeric(Trim(Text(ptxS_Qty).Text)) Then
        Beep
        MsgBox "入力した項目はエラーです｡ "
        Text(ptxS_Qty).SetFocus
        Exit Function
    End If
    
    Text(ptxS_Qty).Text = Format(CLng(Text(ptxS_Qty).Text), "#0")
    If CLng(Text(ptxS_Qty).Text) <= 0 Then
        Beep
        MsgBox "出荷予定数が入力されていません"
        Text(ptxS_Qty).SetFocus
        Exit Function
    End If
                                        '伝票日付
    For i = ptxYY To ptxDD
        If Trim(Text(i)) = "" Then
            Text(i).Text = "0"
        End If
        
        If Not IsNumeric(Text(i).Text) Then
            Beep
            MsgBox "入力した項目はエラーです｡ "
            Text(i).SetFocus
            Exit Function
        Else
            RetBuf = Format(CLng(Text(i).Text), "0000")
            Text(i).Text = Right(RetBuf, Text(i).MaxLength)
        End If
    Next i
        
    If Not IsDate(Text(ptxYY) & "/" & Text(ptxMM) & "/" & Text(ptxDD)) Then
        Beep
        MsgBox "入力した項目はエラーです｡ "
        Text(ptxYY).SetFocus
        Err_Chk = True
        Exit Function
    End If
                
    If Not IsNumeric(Trim(Text(ptxNo))) Then
'        Beep
'        MsgBox "入力した項目はエラーです。"
'        Text(ptxNo).SetFocus
'        Err_Chk = True
'        Exit Function
    Else
        Text(ptxNo) = Format(CLng(Text(ptxNo)), "000000")
    End If
                                    '「０」以外が入力されたら登録済みチェック
                                'ＩＤ№
    If Len(Text(ptxID_No).Text) = 0 Then
    Else
                                                '自動発番以外が入力されたら登録済みチェック
        If Not IsNumeric(Text(ptxID_No).Text) Then
'            Beep                               '英字もエラーにしない
'            MsgBox "入力した項目はエラーです。"
'            Text(ptxDEN_NO).SetFocus
'            Exit Function
        Else
            Text(ptxID_No).Text = Format(CDbl(Text(ptxID_No).Text), "00000000")
        End If
        
        Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU)
'        Call UniCode_Conv(K0_Y_SYU.KEY_CYU_KBN, Right(Combo(pcmbC_Kbn).Text, 1))2004.04.08
        Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Text(ptxID_No).Text)
        sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
                Beep
                MsgBox "入力した項目はエラーです。（出荷予定登録済み）"
                Text(ptxID_No).SetFocus
                Exit Function
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "出荷予定")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
    
    End If
                                '伝票№
    If Len(Text(ptxNo).Text) = 0 Then
    Else
        If Not IsNumeric(Text(ptxNo).Text) Then
'            Beep                               '英字もエラーにしない
'            MsgBox "入力した項目はエラーです。"
'            Text(ptxDEN_NO).SetFocus
'            Exit Function
        Else
            Text(ptxNo).Text = Format(CLng(Text(ptxNo).Text), "000000")
        End If
    
    End If
                                                '向け先チェック
    Call UniCode_Conv(K0_MTS.MUKE_CODE, Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8))
                                                '直送先コード
    Call UniCode_Conv(K0_MTS.SS_CODE, Right(Combo(pcmbMUKE_CODE).Text, 8))
    
    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Beep
            MsgBox "入力した項目はエラーです。（向け先）"
            If MTS_Set_Proc() Then
                Err_Chk = SYS_ERR
                Exit Function
            End If
            Combo(pcmbMUKE_CODE).SetFocus
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "向け先マスタ")
            Err_Chk = SYS_ERR
            Exit Function
    End Select
    
    Err_Chk = False
    
    
End Function
                                            '出荷予定の追加
Private Function Update_Proc() As Integer

Dim sts     As Integer
Dim com     As Integer
Dim ans     As Integer

Dim ID_NO   As String * 12
Dim DEN_NO  As String * 6
    
    Update_Proc = True
                                            '出荷予定編集
    Call UniCode_Conv(Y_SYUREC.WEL_ID, "")                                  '使用子機ＩＤ
    Call UniCode_Conv(Y_SYUREC.PRG_ID, "")                                  '使用中プログラム
    Call UniCode_Conv(Y_SYUREC.KAN_KBN, "0")                                '完了区分
    Call UniCode_Conv(Y_SYUREC.DT_SYU, "R")                                 'データ種別
    Call UniCode_Conv(Y_SYUREC.JGYOBU, Last_JGYOBU)                         '事業部区分
    
                                                                            '注文区分
    Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, Right(Combo(pcmbC_Kbn).Text, 1))
    Call UniCode_Conv(Y_SYUREC.CYU_KBN, Right(Combo(pcmbC_Kbn).Text, 1))
    
                                                                            'ＩＤ№
    If Len(Trim(Text(ptxID_No).Text)) <> 0 Then
        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, Text(ptxID_No).Text)
        Call UniCode_Conv(Y_SYUREC.ID_NO, Text(ptxID_No).Text)
    Else
        
        
        If OSAKA_MODE = "1" Then
        
            sts = Den_No_Set_Proc(31, Last_JGYOBU, ID_NO)
            If sts Then
                Update_Proc = SYS_ERR
                Exit Function
            End If
            Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, Trim(ID_NO) & "01")
            Call UniCode_Conv(Y_SYUREC.ID_NO, Trim(ID_NO) & "01")
        Else
        
            sts = Den_No_Set_Proc(21, Last_JGYOBU, ID_NO)
            If sts Then
                Update_Proc = SYS_ERR
                Exit Function
            End If
            Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO)
            Call UniCode_Conv(Y_SYUREC.ID_NO, ID_NO)
        End If
    End If
                                                                            
    Call UniCode_Conv(Y_SYUREC.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))    '国内外
                                                                    
    Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, Text(ptxCode).Text)              '品目番号
    Call UniCode_Conv(Y_SYUREC.HIN_NO, Text(ptxCode).Text)                  '品目番号
                                                                            '得意先コード
    Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8))
    Call UniCode_Conv(Y_SYUREC.MUKE_CODE, Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8))
                                                                            '直送先コード
    Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, Right(Combo(pcmbMUKE_CODE).Text, 8))
    Call UniCode_Conv(Y_SYUREC.SS_CODE, Right(Combo(pcmbMUKE_CODE).Text, 8))
                                                                            '出荷日
    Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, (Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text))
    Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, (Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text))
    Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, (Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text))
                
    

    
    Call UniCode_Conv(Y_SYUREC.JGYOBA, "")                                  '事業場
    Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")                                'データ区分
    Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")                                '取引区分
                                                                            
    Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, "")
    Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, "")
                                                                            
                                                                            '伝票№
    If Len(Trim(Text(ptxNo).Text)) <> 0 Then
        Call UniCode_Conv(Y_SYUREC.DEN_NO, Text(ptxNo).Text)
    Else
        
        
        If OSAKA_MODE = "1" Then
            sts = Den_No_Set_Proc(32, Last_JGYOBU, ID_NO)
            If sts Then
                Update_Proc = SYS_ERR
                Exit Function
            End If
            Call UniCode_Conv(Y_SYUREC.DEN_NO, Trim(ID_NO))
        Else
        
            sts = Den_No_Set_Proc(20, Last_JGYOBU, DEN_NO)
            If sts Then
                Update_Proc = SYS_ERR
                Exit Function
            End If
            Call UniCode_Conv(Y_SYUREC.DEN_NO, DEN_NO)
        End If
    End If
                                                                            '出庫数量
    Call UniCode_Conv(Y_SYUREC.SURYO, Format(CLng(Text(ptxS_Qty).Text), "0000000"))
        
    Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, "")                             '出庫収支
    
    Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")
    
    
    Call UniCode_Conv(Y_SYUREC.ODER_NO, "")                                 'オーダー番号
    Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")                                 'アイテム番号
    Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")
    Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")
                                                                            '得意先名称
    Call UniCode_Conv(Y_SYUREC.MUKE_NAME, StrConv(MTSREC.MUKE_NAME, vbUnicode))
                                                                            '注文区分名称
    Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, Left(Combo(pcmbC_Kbn).Text, Len(Combo(pcmbC_Kbn).Text) - 1))
                                                                            '品名
    Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(ITEMREC.HIN_NAME, vbUnicode))
        
        
    
    Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
    Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
    Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
        
        
    Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, "")
    Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")
    Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")
    Call UniCode_Conv(Y_SYUREC.BIKOU2, "")
    Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")
    Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, "")
    Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, "")
    Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, "")
    Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, "")
    Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, "")
    Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.BIKOU1, "")
    Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, "")
    Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, "")
    Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(ITEMREC.HIN_NAME, vbUnicode))
    Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, "")
    Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, "")
    Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, "")
    Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, "")
    Call UniCode_Conv(Y_SYUREC.KISHU_CODE, "")
    Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, "")
    Call UniCode_Conv(Y_SYUREC.SS_CODE, "")
    Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, "")
        
        
        
        
        
        
        
                                                                            'ホスト棚番
    Call UniCode_Conv(Y_SYUREC.HTANABAN, Text(ptxS_No).Text & Text(ptxRetu).Text & Text(ptxRen).Text & Text(ptxDan).Text)
    
    Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")                               '完了日付
    Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")                                 '完了日付
    Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")                              '検品日付
    Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")                                 '特売り区分
    
    Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "00000000")                      '実績数量
    
                                                                            '更新日時
    Call UniCode_Conv(Y_SYUREC.INS_NOW, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, "")       '2006.07.20
    Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, "")              '2006.07.20
                                                            '2006.07.20
    Call UniCode_Conv(Y_SYUREC.LK_MUKE_CODE, Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8))
    Call UniCode_Conv(Y_SYUREC.LK_SEQ_NO, "")              '2006.07.20
    
    
    
    Call UniCode_Conv(Y_SYUREC.FILLER, "")

    Do
        sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case BtErrDuplicates
                If Len(Trim(Text(ptxID_No).Text)) = 0 Then
                                            '自動発番データ重複は再発行
                    sts = Den_No_Set_Proc(21, Last_JGYOBU, ID_NO)
                    If sts Then
                        Update_Proc = SYS_ERR
                        Exit Function
                    End If
    
                    Call UniCode_Conv(Y_SYUREC.ID_NO, ID_NO)
                
                Else
                    Call File_Error(sts, BtOpInsert, "出荷予定データ")
                    Update_Proc = SYS_ERR
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "出荷予定データ")
                Update_Proc = SYS_ERR
                Exit Function
        End Select
    Loop
    
    
    
    If OSAKA_MODE = "1" Then
    
    
    
        '-------------------------------------------    出荷予定(ﾎｽﾄｲﾒｰｼﾞ)
        'ID_NO
        Call UniCode_Conv(Y_SYU_HREC.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))
        '№
        Call UniCode_Conv(Y_SYU_HREC.SYUKA_NO, "")
        '出荷日
        Call UniCode_Conv(Y_SYU_HREC.SYUKA_YMD, Text(ptxYY).Text & _
                                                Text(ptxMM).Text & _
                                                Text(ptxDD).Text)
        '送り先名
        Call UniCode_Conv(Y_SYU_HREC.OKURISAKI, "")
        '売り伝
        Call UniCode_Conv(Y_SYU_HREC.URIDEN, "0")
        '伝票番号
        Call UniCode_Conv(Y_SYU_HREC.DEN_NO, Left(StrConv(Y_SYUREC.DEN_NO, vbUnicode), 7))
        '追番
        Call UniCode_Conv(Y_SYU_HREC.SEQ_NO, Right(Trim(StrConv(Y_SYUREC.DEN_NO, vbUnicode)), 1))
        '品番
        Call UniCode_Conv(Y_SYU_HREC.HIN_NO, Text(ptxCode).Text)
        '数量
        Call UniCode_Conv(Y_SYU_HREC.SURYO, Format(CLng(Text(ptxS_Qty).Text), "0000000"))
        '注文№
        Call UniCode_Conv(Y_SYU_HREC.ODER_NO, "")
        '得意先
        Call UniCode_Conv(Y_SYU_HREC.MUKE_CODE, Text(ptxMUKE_CODE).Text)
        '得意先名
        Call UniCode_Conv(Y_SYU_HREC.MUKE_NAME, StrConv(MTSREC.MUKE_NAME, vbUnicode))
        '備考
        Call UniCode_Conv(Y_SYU_HREC.BIKOU, "")
        '運送会社名
        Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, "")
        '取込み日時（入力日時）
        Call UniCode_Conv(Y_SYU_HREC.INS_NOW, Format(Now, "YYYYMMDDHHMMSS"))
        '出荷ﾗﾍﾞﾙ印刷日時（入力日時）
        Call UniCode_Conv(Y_SYU_HREC.PRINT_NOW, "")
        
        
        'ﾃﾞｰﾀ発生順
        Call UniCode_Conv(Y_SYU_HREC.DATA_CNT, "00001")
        '送り状№
        Call UniCode_Conv(Y_SYU_HREC.OKURI_NO, "")
        '検品日時
        Call UniCode_Conv(Y_SYU_HREC.KENPIN_NOW, "")
        '検品担当者
        Call UniCode_Conv(Y_SYU_HREC.KENPIN_TANTO_CODE, "")
        '口数
        Call UniCode_Conv(Y_SYU_HREC.xKUTI_SU, "")
        '強制完了ﾌﾗｸﾞ
        Call UniCode_Conv(Y_SYU_HREC.KYOSEI_END, "")
        'ｷｬﾝｾﾙﾌﾗｸﾞ
        Call UniCode_Conv(Y_SYU_HREC.CANCEL_F, "")
        '入力備考
        Call UniCode_Conv(Y_SYU_HREC.INPUT_BIKOU, "")
        '入力備考
        Call UniCode_Conv(Y_SYU_HREC.INS_BIN, "09")
        '入力備考
        Call UniCode_Conv(Y_SYU_HREC.KUTI_SU, "")
        '事業部
        Call UniCode_Conv(Y_SYU_HREC.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
        '国内外
        Call UniCode_Conv(Y_SYU_HREC.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
        '出庫表№
        Call UniCode_Conv(Y_SYU_HREC.SYU_NO, "")
        '出庫実績数
        Call UniCode_Conv(Y_SYU_HREC.J_SURYO, "")
        '集約送り先CD
        Call UniCode_Conv(Y_SYU_HREC.COL_OKURISAKI_CD, "")
        '送り先CD
        Call UniCode_Conv(Y_SYU_HREC.OKURISAKI_CD, "")
        '住所
        Call UniCode_Conv(Y_SYU_HREC.JYUSHO, "")
        '電話番号
        Call UniCode_Conv(Y_SYU_HREC.TEL_NO, "")
        '郵便番号
        Call UniCode_Conv(Y_SYU_HREC.YUBIN_NO, "")
        '重量
        Call UniCode_Conv(Y_SYU_HREC.JURYO, "")
        '才数
        Call UniCode_Conv(Y_SYU_HREC.SAI_SU, "")
        '送り状№　枝番
        Call UniCode_Conv(Y_SYU_HREC.OKURI_NO_SEQ, "")
        '梱包区分
        Call UniCode_Conv(Y_SYU_HREC.KONPOU_F, "")
        '口数(単体)
        Call UniCode_Conv(Y_SYU_HREC.KUTI_SU_TAN, "")
        '才数(単体)
        Call UniCode_Conv(Y_SYU_HREC.SAI_SU_TAN, "")
        '送り状№　枝番
        Call UniCode_Conv(Y_SYU_HREC.OKURI_NO_SEQ_TO, "")
        '才数(単体:修正不可)    2010.11.01
        Call UniCode_Conv(Y_SYU_HREC.SAI_SU_TAN_SAV, "")
        '才数計算値(梱包単位)   2010.11.01
        Call UniCode_Conv(Y_SYU_HREC.SAI_SU_CALC, "")
        '口数計算値(梱包単位)   2010.11.9
        Call UniCode_Conv(Y_SYU_HREC.KUTI_SU_CALC, "")
        '件管№　　　■管理№(上)   2011.04.30
        Call UniCode_Conv(Y_SYU_HREC.SEK_KEN_NO, "")
        '品管№　　　■管理№(下)   2011.04.30
        Call UniCode_Conv(Y_SYU_HREC.SEK_HIN_NO, "")
        '注文ﾃﾞｰﾀ照合担当       2011.05.02
        Call UniCode_Conv(Y_SYU_HREC.SEK_SHOGO_TANTO, "")
        '注文ﾃﾞｰﾀ照合日時       2011.05.02
        Call UniCode_Conv(Y_SYU_HREC.SEK_SHOGO_DATETIME, "")
        '検品実績　バラ     2012.10.02
        Call UniCode_Conv(Y_SYU_HREC.CNT_BARA_SU, "")
        '検品実績　箱       2012.10.02
        Call UniCode_Conv(Y_SYU_HREC.CNT_HAKO_SU, "")
        '外装入り数         2012.10.02
        Call UniCode_Conv(Y_SYU_HREC.GAISO_IRI_QTY, "")
        '品番読込み回数     2012.10.02
        Call UniCode_Conv(Y_SYU_HREC.Y_HIN_CHK_CNT, "")
        '品番読込み済み回数 2012.10.02
        Call UniCode_Conv(Y_SYU_HREC.J_HIN_CHK_CNT, "")
        '検品中品番         2012.10.02
        Call UniCode_Conv(Y_SYU_HREC.KEN_HINBAN, "")
        '着店コード         2017.02.08
        Call UniCode_Conv(Y_SYU_HREC.TYAKUTEN, "")
        'FILLER
        Call UniCode_Conv(Y_SYU_HREC.FILLER, "")
        '追加　担当者       2011.05.06
        Call UniCode_Conv(Y_SYU_HREC.INS_TANTO, "F103070")
        '追加　日時         2011.05.06
        Call UniCode_Conv(Y_SYU_HREC.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))
        '更新　担当者       2011.05.06
        Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, "")
        '更新　日時         2011.05.06
        Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, "")
        
        
        
        Do
            sts = BTRV(BtOpInsert, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
            
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                Case BtErrDuplicates
                    'ﾃﾞｰﾀの矛盾が発生するのでこのまま処理中断
                    ans = MsgBox("伝票IDが、重複しています。更新処理中止します", vbOK, "確認入力")
                    
                    Call File_Error(sts, BtOpInsert, "出荷予定(ﾎｽﾄﾃﾞｰﾀ)データ")
                    Update_Proc = SYS_CANCEL
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpInsert, "出荷予定(ﾎｽﾄﾃﾞｰﾀ)データ")
                    Update_Proc = SYS_ERR
                    Exit Function
            End Select
        Loop
    End If
    
    If SYUKA_LOG_ON Then
        Call SYUKA_LOG_OUT_PROC("INS", "AFT")
    End If

    Beep
    MsgBox "伝票№:" & StrConv(Y_SYUREC.DEN_NO, vbUnicode) & " ID:" & StrConv(Y_SYUREC.ID_NO, vbUnicode)

    Update_Proc = False
End Function


Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
                                            '入力項目のクリアーとみなす
    Select Case Index
        Case pcmbC_Kbn
            Combo(pcmbNAIGAI).SetFocus
        Case pcmbNAIGAI
            Text(ptxCode).SetFocus
            
        Case pcmbMUKE_CODE
            Text(ptxMUKE_CODE) = Trim(Right(Combo(pcmbMUKE_CODE).Text, 16))
    End Select

End Sub

Private Sub Combo_LostFocus(Index As Integer)

    Select Case Index
        Case pcmbMUKE_CODE
            Text(ptxMUKE_CODE) = Trim(Right(Combo(pcmbMUKE_CODE).Text, 16))
    
    
    End Select

End Sub

Private Sub Command_Click(Index As Integer)
Dim yn As Integer
Dim sts As Integer

    Select Case Index
        Case 0                      '更新
                                    'エラーチェック
            sts = Err_Chk()
            Select Case sts
                Case False
                Case True
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
            
            Beep
            yn = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                sts = Update_Proc()
                Select Case sts
                    Case False, True
                    Case SYS_ERR
                        Unload Me
                End Select
            End If
            'Call Clear_Field()
            Text(ptxCode) = ""
            Text(ptxNo) = ""
            Text(ptxName) = ""      '2020/04/14 品名空白
            Text(ptxS_Qty) = ""     '2020/04/14 出荷予定数空白
            Text(ptxS_No) = ""      '2020/04/14 倉庫空白
            Text(ptxRetu) = ""      '2020/04/14 列空白
            Text(ptxRen) = ""       '2020/04/14 連空白
            Text(ptxDan) = ""       '2020/04/14 段空白
            Text(ptxCode).SetFocus
            Call Text_GotFocus(ptxCode)
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
    
    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                
                                '出荷ログファイル名取り込み
    If SYUKA_LOGF_GET_PROC() Then
        Beep
        MsgBox "出荷ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

                                '事業部取り込み
    If JGYOB_TB_Set() Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If Trim(JGYOBU_T(i).CODE) = "" Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1030701.Caption = "出荷予定登録（" + RTrim(JGYOBU_T(i).NAME) + ") " & LAST_UPDATE_DAY
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)



                                'ﾃﾞﾌｫﾙﾄ注文区分取り込み 2009.04.14
    If GetIni(App.EXEName, "DEF_CYU_KBN", "SYS", c) Then
        DEF_CYU_KBN = CYU_KBN_TUK
    Else
        DEF_CYU_KBN = Trim(c)
    End If
                                '大阪？ 2010.03.23
    If GetIni(App.EXEName, "OSAKA_MODE", "SYS", c) Then
        OSAKA_MODE = "0"
    Else
        OSAKA_MODE = Trim(c)
    End If





                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '向け先管理マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '発番マスタＯＰＥＮ
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定データファイルＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                
                                '出荷予定データファイルＯＰＥＮ
    If Y_SYU_H_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '画面初期設定
    Combo(pcmbC_Kbn).Clear
    Combo(pcmbC_Kbn).AddItem CYU_KBN_1$ & Space(5) & CYU_KBN_TUK$
    Combo(pcmbC_Kbn).AddItem CYU_KBN_2$ & Space(5) & CYU_KBN_SPO$
    Combo(pcmbC_Kbn).AddItem CYU_KBN_3$ & Space(5) & CYU_KBN_HJU$
    Combo(pcmbC_Kbn).AddItem CYU_KBN_E$ & Space(5) & CYU_KBN_BOU$
    '↓97.08.06 「特売（緊急）」の予定は存在しない
'    Combo(pcmbC_Kbn).AddItem CYU_KBN_T$ & Space(5) & CYU_KBN_KIN$
    'Combo(pcmbC_Kbn).Text = CYU_KBN_1$
    '↓2001.03.28 「特売」の予定も登録可とした！
'    Combo(pcmbC_Kbn).AddItem CYU_KBN_4$ & Space(5) & CYU_KBN_TOK$
    
    
    '2009.04.14
    Combo(pcmbC_Kbn).ListIndex = 0
    For i = 0 To Combo(pcmbC_Kbn).ListCount - 1
        If DEF_CYU_KBN = Right(Combo(pcmbC_Kbn).List(i), 1) Then
            Combo(pcmbC_Kbn).ListIndex = i
            Exit For
        End If
    Next i
    
    Combo(pcmbNAIGAI).AddItem NAIGAI1 & Space(5) & NAIGAI_NAI
    Combo(pcmbNAIGAI).AddItem NAIGAI2 & Space(5) & NAIGAI_GAI
    Combo(pcmbNAIGAI).ListIndex = 0
                        '向け先設定
    If MTS_Set_Proc() Then
        Unload Me
    End If
    
    Call Clear_Field
    Text(ptxYY) = Mid(Date, 1, 4)
    Text(ptxMM) = Mid(Date, 6, 2)
    Text(ptxDD) = Mid(Date, 9, 2)  '2020/04/14 伝票日付を本日日付自動表示に変更
    
    Text(ptxID_No).SetFocus

    End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '向け先管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先管理マスタ")
        End If
    End If
                                            '発番マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "発番マスタ")
        End If
    End If
                                            '出荷予定データファイルＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定データファイル")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1030701 = Nothing

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    'メニューより終了要求
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1030701.Caption = "出荷予定登録（" + RTrim(JGYOBU_T(Index).NAME) + ") " & LAST_UPDATE_DAY
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If

End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim RetBuf  As String
Dim i       As Integer
Dim sts     As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub
            
    Select Case Index
        Case ptxCode                '品目コード
            
            
            
            If Trim(Text(ptxCode).Text) = "" Then   '2018.04.20
                Beep                                '2018.04.20
                MsgBox "品番は必須入力です。 "      '2018.04.20
                Text(ptxCode).SetFocus              '2018.04.20
                Exit Sub                            '2018.04.20
            End If                                  '2018.04.20
            
            
            
            
            sts = Item_Dsp()
            Select Case sts
                Case False
                Case True
                    Beep
                    MsgBox "入力した項目はエラーです｡ (未登録エラー)"
                    Text(Index).SetFocus
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
        Case ptxMUKE_CODE         '向け先（コード入力用）
            Call UniCode_Conv(K2_MTS.MUKE_CODE, Text(Index).Text)
            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K2_MTS, Len(K2_MTS), 2)
            Select Case sts
                Case BtNoErr
                    If Len(Trim(StrConv(MTSREC.SS_CODE, vbUnicode))) <> 0 Then
                        Beep
                        MsgBox "入力した項目はエラーです。(向け先コード)"
                        Exit Sub
                    End If
                                
                Case BtErrKeyNotFound
                                
                    Call UniCode_Conv(K3_MTS.SS_CODE, Text(Index).Text)
                                                        
                    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K3_MTS, Len(K3_MTS), 3)
                    Select Case sts
                        Case BtNoErr
                                        
                        Case BtErrKeyNotFound
                            Beep
                            MsgBox "入力した項目はエラーです。(向け先コード)"
                            Exit Sub
                                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                            Unload Me
                    End Select

                Case Else
                    Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                    Unload Me
            End Select


            For i = 0 To Combo(pcmbMUKE_CODE).ListCount - 1 '向け先
    
                If Right(Combo(pcmbMUKE_CODE).List(i), 16) = StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode) Then
                    Combo(pcmbMUKE_CODE).ListIndex = i
                    Exit For
                End If
            
    
            Next

            
    End Select
    
    For i = Index + 1 To ptxMAX
        If Text(i).Visible And Text(i).Enabled And Not Text(i).Locked Then
            Text(i).SetFocus
            Exit Sub
        End If
    Next i
    Combo(pcmbMUKE_CODE).SetFocus
    
End Sub
Private Function MTS_Set_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer
Dim Edit        As String


    MTS_Set_Proc = True
    
    com = BtOpGetFirst
    
    Combo(pcmbMUKE_CODE).Clear
    
    Do
        sts = BTRV(com, MTS_POS, MTSREC, Len(MTSREC), K1_MTS, Len(K1_MTS), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "向け先マスタ")
                MTS_Set_Proc = SYS_ERR
                Exit Function
        End Select
    
    
        Edit = StrConv(MTSREC.MUKE_NAME, vbUnicode) & "   "
        Edit = Edit & StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode)
        Combo(pcmbMUKE_CODE).AddItem Edit
    
    
        com = BtOpGetNext
    Loop




    If Combo(pcmbMUKE_CODE).ListCount = 0 Then
    Else
        Combo(pcmbMUKE_CODE).ListIndex = 0
    End If


    MTS_Set_Proc = False
End Function


Private Sub Text_LostFocus(Index As Integer)
    
    If Index = ptxCode Then
        Text(ptxCode).Text = StrConv(Text(ptxCode).Text, vbUpperCase)
    End If
End Sub
