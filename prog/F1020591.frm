VERSION 5.00
Begin VB.Form F1020591 
   BackColor       =   &H00FFFFFF&
   Caption         =   "[美的品番管理]入庫現品票印刷"
   ClientHeight    =   7470
   ClientLeft      =   2025
   ClientTop       =   2940
   ClientWidth     =   12900
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
   ScaleHeight     =   7470
   ScaleWidth      =   12900
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   3840
      MaxLength       =   70
      TabIndex        =   2
      Top             =   960
      Width           =   8535
   End
   Begin VB.TextBox text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   3810
      MaxLength       =   4
      TabIndex        =   9
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox text 
      Height          =   375
      Index           =   12
      Left            =   7890
      MaxLength       =   8
      TabIndex        =   12
      Top             =   3060
      Width           =   1092
   End
   Begin VB.ListBox List2 
      Height          =   300
      Left            =   6435
      Sorted          =   -1  'True
      TabIndex        =   51
      Top             =   5220
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      ItemData        =   "F1020591.frx":0000
      Left            =   7860
      List            =   "F1020591.frx":0002
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   6
      Top             =   2040
      Width           =   2790
   End
   Begin VB.ComboBox Combo1 
      Height          =   336
      Left            =   2880
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   14
      Top             =   5760
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "用紙選択"
      Height          =   975
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   1575
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A4"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A5"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   11
      Left            =   4440
      MaxLength       =   3
      TabIndex        =   47
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   10
      Left            =   8280
      MaxLength       =   3
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5760
      Width           =   732
   End
   Begin VB.TextBox text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   9
      Left            =   7920
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   5730
      MaxLength       =   2
      TabIndex        =   11
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   4920
      MaxLength       =   2
      TabIndex        =   10
      Top             =   3120
      Width           =   375
   End
   Begin VB.ComboBox Combo 
      Height          =   336
      Index           =   0
      Left            =   2235
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   5460
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox text 
      Height          =   375
      Index           =   5
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   8
      Top             =   2520
      Width           =   2532
   End
   Begin VB.TextBox text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   3840
      MaxLength       =   3
      TabIndex        =   7
      Top             =   2520
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   1260
      Left            =   690
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3840
      Width           =   9096
   End
   Begin VB.TextBox text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   3810
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1500
      Width           =   2655
   End
   Begin VB.TextBox text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   7890
      MaxLength       =   40
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   420
      Width           =   3135
   End
   Begin VB.TextBox text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   3840
      MaxLength       =   20
      TabIndex        =   1
      Top             =   420
      Width           =   2535
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
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "印  刷"
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
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Top             =   6360
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "確  定"
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
      TabStop         =   0   'False
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "美的品番"
      Height          =   255
      Index           =   18
      Left            =   2280
      TabIndex        =   54
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品名"
      Height          =   255
      Index           =   17
      Left            =   7170
      TabIndex        =   53
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "仕入先"
      Height          =   255
      Index           =   16
      Left            =   6915
      TabIndex        =   52
      Top             =   3180
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "原産国"
      Height          =   255
      Index           =   15
      Left            =   6960
      TabIndex        =   50
      Top             =   2160
      Width           =   750
   End
   Begin VB.Label lblST_TANABAN 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   49
      Top             =   2100
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "標準棚番"
      Height          =   255
      Index           =   13
      Left            =   2655
      TabIndex        =   48
      Top             =   2100
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "枚数合計"
      Height          =   252
      Index           =   12
      Left            =   7200
      TabIndex        =   46
      Top             =   5880
      Width           =   972
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "印刷備考"
      Height          =   252
      Index           =   11
      Left            =   7296
      TabIndex        =   45
      Top             =   3600
      Width           =   1092
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "枚数"
      Height          =   252
      Index           =   10
      Left            =   6456
      TabIndex        =   44
      Top             =   3600
      Width           =   612
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "入数"
      Height          =   252
      Index           =   9
      Left            =   5616
      TabIndex        =   43
      Top             =   3600
      Width           =   612
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "入数"
      Height          =   255
      Index           =   8
      Left            =   7170
      TabIndex        =   42
      Top             =   1620
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "日"
      Height          =   255
      Index           =   7
      Left            =   6210
      TabIndex        =   41
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "月"
      Height          =   255
      Index           =   6
      Left            =   5370
      TabIndex        =   40
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "年"
      Height          =   255
      Index           =   5
      Left            =   4530
      TabIndex        =   39
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "入 荷 日"
      Height          =   255
      Index           =   4
      Left            =   2610
      TabIndex        =   38
      Top             =   3240
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
      Height          =   312
      Left            =   120
      TabIndex        =   37
      Top             =   6960
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   252
      Index           =   3
      Left            =   1356
      TabIndex        =   36
      Top             =   5580
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   180
      TabIndex        =   35
      Top             =   5820
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "印刷備考"
      Height          =   255
      Index           =   2
      Left            =   6735
      TabIndex        =   34
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "印刷枚数"
      Height          =   255
      Index           =   1
      Left            =   2610
      TabIndex        =   33
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品番（内部）"
      Height          =   255
      Index           =   14
      Left            =   2250
      TabIndex        =   32
      Top             =   1620
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品番（外部）"
      Height          =   255
      Index           =   0
      Left            =   2250
      TabIndex        =   31
      Top             =   540
      Width           =   1455
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1020591"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NormalFont As New StdFont           '印刷フォント
Dim Code39Font As New StdFont           '印刷フォント


Private Type Print_tbl_tag              '印刷用テーブル

    NAIGAI          As String * 2
    HIN_GAI         As String * 20
    HIN_NAI         As String * 13
    HIN_NAME        As String
    IRI_QTY         As String * 8
    ST_SOKO         As String * 2
    ST_SOKO_NAME    As String * 5
    ST_RETU         As String * 2
    ST_REN          As String * 2
    ST_DAN          As String * 2
'    BIKOU           As String * 15
'    BIKOU           As String * 20
    BIKOU           As String
    GENSAN          As String * 22
'2010.10.07
    SHIIRE_WORK_CENTER As _
                       String * 8

'    B_HIN_CODE      As String * 70     2013.12.18 41桁までとする
    B_HIN_CODE      As String * 41      '2013.12.18
End Type

Dim Print_tbl(0 To 6, 0 To 1) _
                    As Print_tbl_tag





Private Const ptxHin_Gai% = 0       '品番(外)
Private Const ptxHin_Name% = 1      '品名

Private Const ptxB_Hin_Code% = 2    '美的品番



Private Const ptxHin_Nai% = 3       '品番(内)
Private Const ptxMaiSuu% = 4        '印刷枚数
Private Const ptxBikou% = 5         '印刷備考
Private Const ptxNyuka_YY% = 6      '入荷日　年
Private Const ptxNyuka_MM% = 7      '入荷日　月
Private Const ptxNyuka_DD% = 8      '入荷日　日
Private Const ptxIriSuu% = 9        '入り数
Private Const ptxGoukei% = 10       '合計

Private Const ptxwkMaiSuu% = 11     '保存用印刷枚数

                                    '保存用印刷枚数 2010.10.07
Private Const ptxSHIIRE_WORK_CENTER% = 12

                                    '2010.10.07
Private SHIIRE_WORK_CENTER_F  As Integer


Dim JGYOBU_NAME As String

Dim Printer_tbl() As String
Dim Max_Gyo     As Integer

Private Const Last_Update_Day$ = "(F102059 2013.12.19 16:20) "

Private Function Print_Proc() As Integer

Dim Maisu       As Integer
Dim sts         As Integer
Dim flg         As Boolean

Dim wk_LOOP     As Integer

Dim Gyo         As Integer


Dim Retu        As Integer

Dim wk_Naigai   As String * 1

Dim Wk_Printer As Printer

    Print_Proc = False

'指定帳票用プリンタ情報取得
    For Each Wk_Printer In Printers
        If RTrim(Wk_Printer.DeviceName) = RTrim(Combo1.text) Then
                Set Printer = Wk_Printer
                Exit For
        End If
    Next

    If Option1(0).Value = True Then
        Printer.PaperSize = vbPRPSA5
        Printer.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷
        Max_Gyo = 2
    Else
        Printer.PaperSize = vbPRPSA4
        Printer.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷
        Max_Gyo = 5
    End If


    



    For Gyo = 0 To UBound(Print_tbl)
        For Retu = 0 To 1
        
            Print_tbl(Gyo, Retu).HIN_GAI = " "
        
        Next Retu
    Next Gyo

    Gyo = 0
    Retu = 0


    For wk_LOOP = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(wk_LOOP).CODE = Last_JGYOBU Then
            JGYOBU_NAME = JGYOBU_T(wk_LOOP).NAME
            Exit For
        End If
    Next wk_LOOP



    For wk_LOOP = 0 To List1.ListCount - 1
        wk_Naigai = Left(List1.List(wk_LOOP), 1)
        
Item_Read:
        
        Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_ITEM.NAIGAI, wk_Naigai)
        Call UniCode_Conv(K0_ITEM.HIN_GAI, Mid(List1.List(wk_LOOP), 3, 20))
        flg = False
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                flg = True
            Case Else
                
                
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.19
                If sts > 3000 Or sts = 3 Then
                
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)

                    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    If sts Then
                        Call File_Error(sts, BtOpReset, "棚マスタ")
                        Beep
                        MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
                    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.31
'                                                '倉庫マスタＯＰＥＮ
'                    If SOKO_Open(0) Then
'                        Beep
'                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'                        Unload Me
'                    End If
'                                                '品目マスタＯＰＥＮ
'                    If ITEM_Open(0) Then
'                        Beep
'                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'                        Unload Me
'                    End If
'
'                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.05.30
'                                                'PNマスタＯＰＥＮ
'                    If PN_M_Open(0) Then
'                        Beep
'                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'                        Unload Me
'                    End If
'                    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'                                                '原産国マスタＯＰＥＮ
'                    If GENSAN_Open(0) Then
'                        Beep
'                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'                        Unload Me
'                    End If
                    Call File_Open_Proc
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.31
                
                
                
                
                    GoTo Item_Read

                
                End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.19
                
                
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        End Select
        
        For Maisu = 1 To CInt(Mid(List1.List(wk_LOOP), 49, 3))
            
            DoEvents
            
            If wk_Naigai = NAIGAI_NAI Then
                Print_tbl(Gyo, Retu).NAIGAI = NAIGAI1
            Else
                Print_tbl(Gyo, Retu).NAIGAI = NAIGAI2
            End If
            Print_tbl(Gyo, Retu).HIN_GAI = Mid(List1.List(wk_LOOP), 3, 20)
            If Not flg Then
                Print_tbl(Gyo, Retu).HIN_NAI = StrConv(ITEMREC.HIN_NAI, vbUnicode)
                Print_tbl(Gyo, Retu).HIN_NAME = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                Print_tbl(Gyo, Retu).ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                Print_tbl(Gyo, Retu).ST_RETU = StrConv(ITEMREC.ST_RETU, vbUnicode)
                Print_tbl(Gyo, Retu).ST_REN = StrConv(ITEMREC.ST_REN, vbUnicode)
                Print_tbl(Gyo, Retu).ST_DAN = StrConv(ITEMREC.ST_DAN, vbUnicode)
    
                Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
                        Print_tbl(Gyo, Retu).ST_SOKO_NAME = Left(StrConv(SOKOREC.SOKO_NAME, vbUnicode), 5)
                    Case BtErrKeyNotFound
                        Print_tbl(Gyo, Retu).ST_SOKO_NAME = " "
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                        Beep
                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
                        Unload Me
                End Select
            
            
'                Print_tbl(Gyo, Retu).GENSAN = Trim(Right(Left(List1.List(wk_LOOP), 22), 31))
                
                Print_tbl(Gyo, Retu).GENSAN = Trim(Left(Right(List1.List(wk_LOOP), 39), 22))
                
                
                '2010.10.07
                Print_tbl(Gyo, Retu).SHIIRE_WORK_CENTER = Trim(Left(Right(List1.List(wk_LOOP), 16), 8))     '2013.08.23
                '2010.10.07
            
            
                Call UniCode_Conv(K0_B_ITEM.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_B_ITEM.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_B_ITEM.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                        
                sts = BTRV(BtOpGetEqual, B_ITEM_POS, B_ITEMREC, Len(B_ITEMREC), K0_B_ITEM, Len(K0_B_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Print_tbl(Gyo, Retu).B_HIN_CODE = StrConv(B_ITEMREC.B_HIN_CODE, vbUnicode)
                    Case BtErrKeyNotFound
                        Print_tbl(Gyo, Retu).B_HIN_CODE = ""
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "美的品番ﾃﾞｰﾀ")
                        Beep
                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
                        Unload Me
                End Select
            
            
'
            
            Else
                Print_tbl(Gyo, Retu).HIN_NAI = " "
                Print_tbl(Gyo, Retu).HIN_NAME = " "
                Print_tbl(Gyo, Retu).ST_SOKO = " "
                Print_tbl(Gyo, Retu).ST_RETU = " "
                Print_tbl(Gyo, Retu).ST_REN = " "
                Print_tbl(Gyo, Retu).ST_DAN = " "
                Print_tbl(Gyo, Retu).ST_SOKO_NAME = " "
                Print_tbl(Gyo, Retu).GENSAN = ""
            
                '2010.10.07
                Print_tbl(Gyo, Retu).SHIIRE_WORK_CENTER = ""
                '2010.10.07
            
                Print_tbl(Gyo, Retu).B_HIN_CODE = ""
            
            
            End If
    
            Print_tbl(Gyo, Retu).IRI_QTY = Mid(List1.List(wk_LOOP), 38, 8)
            Print_tbl(Gyo, Retu).BIKOU = Mid(List1.List(wk_LOOP), 55, 20)
    
            Retu = Retu + 1
            If Retu > 1 Then
                Gyo = Gyo + 1
                If Gyo > Max_Gyo Then
                    Call Print_Sub_Proc
                    Printer.NewPage
                    For Gyo = 0 To Max_Gyo
                        For Retu = 0 To 1
        
                            Print_tbl(Gyo, Retu).HIN_GAI = " "
        
                        Next Retu
                    Next Gyo

                    Gyo = 0
                End If
                Retu = 0
            End If
        Next Maisu

    
    Next wk_LOOP
    
    Call Print_Sub_Proc
        
End Function
                                    
                                    '画面初期状態を設定する
Private Sub Clear_Field()
Dim i As Integer
    
    For i = 0 To 4
        text(i).text = ""
    Next i
    text(ptxIriSuu).text = ""
    
    text(ptxSHIIRE_WORK_CENTER).text = ""

    text(ptxGoukei).text = "0"
    text(ptxwkMaiSuu).text = "0"

    lblST_TANABAN(0).Caption = ""

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.31
    Combo(1).Clear
    Combo(1).ListIndex = -1
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.31




End Sub

Private Sub Combo_Click(Index As Integer)
        
        text(ptxHin_Gai).SelStart = 0
        text(ptxHin_Gai).SelLength = Len(RTrim(text(ptxHin_Gai).text))
        text(ptxHin_Gai).SetFocus
End Sub

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            
            
            Select Case Index
                Case 0
                    text(ptxHin_Gai).SetFocus
                Case 1
                    text(ptxMaiSuu).SetFocus
            End Select
        
        
        
        
        Case vbKeyF9
            Command(8).Value = True
        Case vbKeyF12
            Command(11).Value = True
    End Select

End Sub



Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case vbKeyReturn
'            Select Case Index
'                Case 0
'                    Call Clear_Field(0)
'                    List1.Clear
'                    text(0).SetFocus
'            End Select
'    End Select
'
End Sub

Private Sub Command_Click(Index As Integer)

Dim yn              As Integer
Dim RetBuf          As String
Dim sts             As Integer

Dim wkList_Box      As String
Dim wk_Kbn          As String * 1
Dim wk_Bikou        As String * 20
Dim wk_Maisuu       As Integer

Dim wk_IRI_QTY      As String * 8
Dim wk_MAISU        As String * 3


Dim wkGENSAN        As String * 22


Dim wkSHIIRE_WORK_CENTER As String * 8

Dim wkHin_Nai       As String * 13

Dim wkKEPPIN_QTY    As String * 8       '2013.08.23


Select Case Index
        Case 0                              '確定
                                            
                                            '外部品番で読み込み
'            If Len(text(ptxHin_Gai).text) <> 0 Then
                
            '2010.11.25
            If Len(text(ptxHin_Gai).text) <> 0 And Len(text(ptxHin_Nai).text) = 0 Then
                
                
Item_Read:
                Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                If Combo(0).text = NAIGAI1$ Then
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI$)
                    wk_Kbn = NAIGAI_NAI
                Else
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI$)
                    wk_Kbn = NAIGAI_GAI
                End If
                Call UniCode_Conv(K0_ITEM.HIN_GAI, RTrim(text(ptxHin_Gai).text))
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        text(ptxHin_Name).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                        text(ptxHin_Nai).text = StrConv(ITEMREC.HIN_NAI, vbUnicode)
                    Case BtErrKeyNotFound
                        MsgBox "入力したコードは登録されていません。"
                        Exit Sub
                    Case Else
                        
                        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.19
                        If sts > 3000 Or sts = 3 Then
                        
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                        
                            sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            If sts Then
                                Call File_Error(sts, BtOpReset, "棚マスタ")
                                Beep
                                MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
                            End If
                        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.31
        '
        '                                                '倉庫マスタＯＰＥＮ
        '                    If SOKO_Open(0) Then
        '                        Beep
        '                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        '                        Unload Me
        '                    End If
        '                                                '品目マスタＯＰＥＮ
        '                    If ITEM_Open(0) Then
        '                        Beep
        '                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        '                        Unload Me
        '                    End If
        '
        '                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.05.30
        '                                                'PNマスタＯＰＥＮ
        '                    If PN_M_Open(0) Then
        '                        Beep
        '                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        '                        Unload Me
        '                    End If
        '                    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        '                                                '原産国マスタＯＰＥＮ
        '                    If GENSAN_Open(0) Then
        '                        Beep
        '                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        '                        Unload Me
        '                    End If
        
        
                    Call File_Open_Proc
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.31
                
                
                
                
                    GoTo Item_Read

                
                End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.19
                        
                        
                        
                        
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Beep
                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
                        Unload Me
                End Select
                                                        
            Else                            '内部品番で読み込み
                
                                
                
Item_Read2:
                
                '2010.11.25
                If Len(text(ptxHin_Gai).text) = 0 And Len(text(ptxHin_Nai).text) = 0 Then
                    #If Center_chk Then
                        Call UniCode_Conv(K3_ITEM.JGYOBU, Last_JGYOBU)
                        If Combo(0).text = NAIGAI1$ Then
                            Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI_NAI$)
                            wk_Kbn = NAIGAI_NAI
                        Else
                            Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI_GAI$)
                            wk_Kbn = NAIGAI_GAI
                        End If
                        Call UniCode_Conv(K2_ITEM.HIN_NAI, RTrim(text(ptxHin_Nai).text))
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                    #Else
                        Call UniCode_Conv(K1_ITEM.JGYOBU, Last_JGYOBU)
                        If Combo(0).text = NAIGAI1$ Then
                            Call UniCode_Conv(K1_ITEM.NAIGAI, NAIGAI_NAI$)
                            wk_Kbn = NAIGAI_NAI
                        Else
                            Call UniCode_Conv(K1_ITEM.NAIGAI, NAIGAI_GAI$)
                            wk_Kbn = NAIGAI_GAI
                        End If
                        Call UniCode_Conv(K1_ITEM.HIN_NAI, RTrim(text(ptxHin_Nai).text))
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K1_ITEM, Len(K1_ITEM), 1)
                    #End If
                    Select Case sts
                        Case BtNoErr
                            text(ptxHin_Gai).text = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                            text(ptxHin_Name).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                        Case BtErrKeyNotFound
                            MsgBox "入力したコードは登録されていません。"
                            Exit Sub
                        Case Else
                            
                            
                            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.19
                If sts > 3000 Or sts = 3 Then
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)

                
                    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    If sts Then
                        Call File_Error(sts, BtOpReset, "棚マスタ")
                        Beep
                        MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
                    End If
                
                
                                                '倉庫マスタＯＰＥＮ
                    If SOKO_Open(0) Then
                        Beep
                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
                        Unload Me
                    End If
                                                '品目マスタＯＰＥＮ
                    If ITEM_Open(0) Then
                        Beep
                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
                        Unload Me
                    End If
                    
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.05.30
                                                'PNマスタＯＰＥＮ
                    If PN_M_Open(0) Then
                        Beep
                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
                        Unload Me
                    End If
                    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                                '原産国マスタＯＰＥＮ
                    If GENSAN_Open(0) Then
                        Beep
                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
                        Unload Me
                    End If
                
                
                
                
                    GoTo Item_Read2

                
                End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.19
                            
                            
                            
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Beep
                            MsgBox "システム異常が発生しました。処理を中止して下さい。"
                            Unload Me
                    End Select
                End If
            End If
                                            'エラーチェック
            If Len(RTrim(text(ptxHin_Gai).text)) = 0 Then
                Beep
                MsgBox "入力した項目はエラーです。"
                text(ptxHin_Gai).SetFocus
                Exit Sub
            End If
    
        
            If Len(text(ptxMaiSuu).text) = 0 Then
                text(ptxMaiSuu).text = "1"
            End If
            
            
            If Not IsNumeric(text(ptxMaiSuu).text) Then
                Beep
                MsgBox "入力した項目はエラーです。"
                text(ptxMaiSuu).SetFocus
                Exit Sub
            Else
                text(ptxMaiSuu).text = Format(CInt(text(ptxMaiSuu).text), "#0")
            
            End If
            If CInt(text(ptxMaiSuu).text) < 1 Then
                Beep
                MsgBox "入力した項目はエラーです。"
                text(ptxMaiSuu).SetFocus
                Exit Sub
            End If
            
            If text(ptxIriSuu).text = "" Then
            Else
                If Len(Trim(text(ptxIriSuu).text)) = 0 Then
                Else
                    If Not IsNumeric(text(ptxIriSuu).text) Then
                        Beep
                        MsgBox "入力した項目はエラーです。"
                        text(ptxIriSuu).SetFocus
                        Exit Sub
                    End If
                End If
            End If
            
            Beep
            yn = MsgBox("確定しますか？", vbYesNo + vbQuestion, "確認入力")
            
            If yn = vbYes Then
                wk_Kbn = NAIGAI_NAI
                                
                wkList_Box = wk_Kbn & " " & StrConv(ITEMREC.HIN_GAI, vbUnicode) + " "


                '2010.11.25
'                wkList_Box = wkList_Box & Left(StrConv(ITEMREC.HIN_NAI, vbUnicode), 13) + " "
                wkHin_Nai = text(ptxHin_Nai).text
                wkList_Box = wkList_Box & wkHin_Nai + " "
                '2010.11.25
                
                
                If Not IsNumeric(text(ptxIriSuu).text) Then
                    wk_IRI_QTY = ""
                Else
                    wk_IRI_QTY = Format(CLng(text(ptxIriSuu).text), "#0")
                End If
                wk_IRI_QTY = Space(Len(wk_IRI_QTY) - Len(Trim(wk_IRI_QTY))) & Trim(wk_IRI_QTY)
                
                wkList_Box = wkList_Box & wk_IRI_QTY & "   "
                
                wk_MAISU = Format(CLng(text(ptxMaiSuu).text), "#0")
                wk_MAISU = Space(Len(wk_MAISU) - Len(Trim(wk_MAISU))) & Trim(wk_MAISU)
                
                wkList_Box = wkList_Box & wk_MAISU & "   "
                wk_Bikou = text(ptxBikou).text
                wkList_Box = wkList_Box & wk_Bikou & "   "
                wkList_Box = wkList_Box & StrConv(ITEMREC.HIN_NAME, vbUnicode) + " "
                
                If Combo(1).ListCount > 1 Then
                    
                    wkGENSAN = Left(Combo(1).text, 20) & "*" & Format(Combo(1).ListCount, "0")
                    wkList_Box = wkList_Box & wkGENSAN & " "
                Else
                     
                    wkGENSAN = Left(Combo(1).text, 20)
                    wkList_Box = wkList_Box & wkGENSAN & " "
                End If
                
                '2010.10.07
                
                wkSHIIRE_WORK_CENTER = text(ptxSHIIRE_WORK_CENTER).text
                wkList_Box = wkList_Box & wkSHIIRE_WORK_CENTER
                
                
                
                
                List1.AddItem wkList_Box
            End If
                        
            If Item_Update_Proc() Then
                Unload Me
            End If
            
            wk_Maisuu = CInt(text(ptxGoukei).text) - CInt(text(ptxwkMaiSuu).text) + CInt(text(ptxMaiSuu).text)
            
            Call Clear_Field
            text(ptxGoukei).text = Format(wk_Maisuu, "#0")
            text(ptxB_Hin_Code).SetFocus
        Case 8                              '印刷
            Beep
            yn = MsgBox("印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                
                sts = Print_Proc()
                Printer.EndDoc
                Call Clear_Field
                List1.Clear
            End If
            
            text(ptxB_Hin_Code).SetFocus
            
        Case 11                             '終了
            If List1.ListCount = 0 Then
                Unload Me
            End If
            Beep
            yn = MsgBox("終了しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                Unload Me
            End If
            text(ptxB_Hin_Code).SetFocus
            
        Case Else
            Beep
    End Select
    
End Sub


Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim i           As Integer
Dim c           As String * 128
Dim sts         As Integer
Dim Pri_Name    As Printer
Dim DEF         As String
    
    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
    
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If

'2012.12.15    For i = 0 To UBound(JGYOBU_T) - 1
    For i = 0 To UBound(JGYOBU_T)   '2012.12.15
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1020591.Caption = "[美的品番管理]入庫現品票印刷（" + RTrim(JGYOBU_T(i).NAME) + ") " & Last_Update_Day
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)

                                
    
    Call File_Open_Proc
    
    
    
    
                                'デフォルト用紙サイズ取り込み
    If GetIni(App.EXEName, "DEF", App.EXEName, c) Then
        c = ""
    End If
    DEF = RTrim(c)
                                
                                
                                '仕入れ先更新可否   2010.10.07
    If GetIni(App.EXEName, "SHIIRE_WORK_CENTER_F", App.EXEName, c) Then
        SHIIRE_WORK_CENTER_F = True
    Else
    
        If Trim(c) = "1" Then
            SHIIRE_WORK_CENTER_F = False
        Else
            SHIIRE_WORK_CENTER_F = True
        End If
    End If
    text(ptxSHIIRE_WORK_CENTER).Locked = SHIIRE_WORK_CENTER_F
                                
                                
                                '印刷フォント設定
    With Code39Font
        .NAME = Label1.FontName
        .Size = Label1.FontSize
    End With
    Set Printer.Font = Code39Font
                                '印刷フォント設定
    With NormalFont
        .NAME = F1020591.FontName
        .Size = F1020591.FontSize
    End With
    Set Printer.Font = NormalFont
                                
                                '画面初期設定
    
    If DEF = Trim(Option1(0).Caption) Then
        Option1(0).Value = True
        Option1(1).Value = False
    Else
        If DEF = Trim(Option1(1).Caption) Then
            Option1(0).Value = False
            Option1(1).Value = True
        Else
            Option1(0).Value = True
            Option1(1).Value = False
        End If
    End If
    
    Combo(0).AddItem NAIGAI1$
    Combo(0).AddItem NAIGAI2$
    Combo(0).text = NAIGAI1$
    
    text(ptxNyuka_YY).text = Mid(Date, 1, 4)
    text(ptxNyuka_MM).text = Mid(Date, 6, 2)
    text(ptxNyuka_DD).text = Mid(Date, 9, 2)
    
    
    Call Clear_Field
    List1.Clear
    
    Combo1.Clear
    
    For Each Pri_Name In Printers
        If Pri_Name.DeviceName = Printer.DeviceName Then
            Combo1.AddItem Pri_Name.DeviceName
        End If
    Next
    
    For Each Pri_Name In Printers
        If Pri_Name.DeviceName <> Printer.DriverName Then
            Combo1.AddItem Pri_Name.DeviceName
        End If
    Next
    
    
    Combo1.ListIndex = 0
    
    text(ptxB_Hin_Code).SetFocus
    
    End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer


    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.05.30
                                            'PNマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, PN_M_POS, PN_MREC, Len(PN_MREC), K0_PN_M, Len(K0_PN_M), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "PNマスタ")
            Beep
            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        End If
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    
    
                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "棚マスタ")
            Beep
            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
            Beep
            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        End If
    End If
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "棚マスタ")
        Beep
        MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
    End If

    End
End Sub

Private Sub List1_DblClick()

Dim ans     As Integer

    
    ans = MsgBox("指定行を削除しますか？", vbYesNo + vbDefaultButton1, "確認入力")
    
    If ans = vbYes Then
        List1.RemoveItem List1.ListIndex
    End If

'Dim sts As Integer
'Dim sts_QTY
'Dim Mode As Integer
'Dim wk_Index As Integer
'Dim RetBuf As String
'
'Dim wk_Naigai   As String * 1
'
'                                                'リストボックスより項目表示
'    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
'    wk_Naigai = Right(List1.List(List1.ListIndex), 1)
'    If wk_Naigai = "1" Then
'        Combo(0).ListIndex = 0
'    Else
'        Combo(0).ListIndex = 1
'    End If
'    Call UniCode_Conv(K0_ITEM.NAIGAI, wk_Naigai)
'
'    '97.10.12
'    wk_Index = List1.ListIndex
'    Call UniCode_Conv(K0_ITEM.HIN_GAI, Mid$(List1.List(List1.ListIndex), 1, 13))
'                                                '外部品番で読み込み
'    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'    Select Case sts
'        Case BtNoErr
'            '97.10.12
'            Text(0) = Mid$(List1.List(List1.ListIndex), 1, 13)
'            Text(1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
'            Text(2) = RTrim(StrConv(ITEMREC.HIN_NAI, vbUnicode))
'            Text(3) = Mid$(List1.List(List1.ListIndex), 66, 3)
'            Text(10) = Mid$(List1.List(List1.ListIndex), 66, 3)
'            Text(4) = Trim(Mid$(List1.List(List1.ListIndex), 72, 10))
'            Text(8) = Trim(Mid$(List1.List(List1.ListIndex), 55, 8))
'            Text(8).SetFocus
'            List1.RemoveItem wk_Index
'
'        Case BtErrKeyNotFound           'これは無いはず
'            MsgBox "マスタ内容が変更されています。最新情報を再表示します。"
'            If Len(Text(0).Text) <> 0 Then
'                Mode = 0
'            Else
'                Mode = 1
'            End If
'
'        Case Else
'            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
'            Beep
'            MsgBox "システム異常が発生しました。処理を中止して下さい。"
'            Unload Me
'    End Select

End Sub


Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer

'2012.12.22    For i = 0 To UBound(JGYOBU_T) - 1
    For i = 0 To UBound(JGYOBU_T)                   '2012.12.22
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1020591.Caption = "[美的品番管理]入庫現品票印刷（" + RTrim(JGYOBU_T(Index).NAME) + ") " & Last_Update_Day
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub

Private Sub Text_GotFocus(Index As Integer)
    If text(Index).TabStop = True Then
        text(Index) = Trim(text(Index).text)
        text(Index).SelStart = 0
        text(Index).SelLength = Len(text(Index).text)
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim RetBuf      As String
Dim i           As Integer
Dim sts         As Integer
Dim sts_QTY     As Integer



    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown
            Select Case Index
                Case 0
                    If Len(text(ptxHin_Gai).text) <> 0 Then
                                                
    
                        text(Index).text = RTrim(StrConv(text(Index).text, vbUpperCase))
    
    
                                                
                                                '外部品番で読み込み
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                        If Combo(0).text = NAIGAI1$ Then
                            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI$)
                        Else
                            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI$)
                        End If
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, RTrim(text(ptxHin_Gai).text))
                        
Item_Read:
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                                text(ptxHin_Name).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                
                                
                                
                                
                                
                                
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    美的品番
                                Call UniCode_Conv(K0_B_ITEM.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                                Call UniCode_Conv(K0_B_ITEM.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                                Call UniCode_Conv(K0_B_ITEM.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
Item_Read_B:
                                
                                sts = BTRV(BtOpGetEqual, B_ITEM_POS, B_ITEMREC, Len(B_ITEMREC), K0_B_ITEM, Len(K0_B_ITEM), 0)
                                Select Case sts
                                    Case BtNoErr
                                    Case BtErrKeyNotFound
                                    
                                        Call UniCode_Conv(B_ITEMREC.B_HIN_CODE, "")
                                    
                                    Case Else
                                
                                        If sts > 3000 Or sts = 3 Then
                                        
                                            Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                        
                                            sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                                            If sts Then
                                                Call File_Error(sts, BtOpReset, "棚マスタ")
                                                Beep
                                                MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
                                            End If
                                        
                        
                                            Call File_Open_Proc
                                        
                                            GoTo Item_Read_B
                                        
                                        End If
                                
                                End Select
                                
                                
                                text(ptxB_Hin_Code).text = RTrim(StrConv(B_ITEMREC.B_HIN_CODE, vbUnicode))
                                
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    美的品番
                                
                                
                                
                                text(ptxHin_Nai).text = StrConv(ITEMREC.HIN_NAI, vbUnicode)
                                
                                
                                '2010.10.07
                                'text(ptxBikou).text = StrConv(ITEMREC.BIKOU, vbUnicode)
'                                If Trim(StrConv(ITEMREC.BIKOU20, vbUnicode)) = "" Or _
'                                    Mid(StrConv(ITEMREC.BIKOU20, vbUnicode), 1, 1) < " " Then
'
'                                    Call UniCode_Conv(ITEMREC.BIKOU20, StrConv(ITEMREC.BIKOU, vbUnicode))
'
'                                End If
                                text(ptxBikou).text = StrConv(ITEMREC.BIKOU20, vbUnicode)
                                '2010.10.07
                                
                                
                                
                                
                                If IsNumeric(StrConv(ITEMREC.IRI_QTY, vbUnicode)) Then
                                    text(ptxIriSuu).text = Format(CLng(StrConv(ITEMREC.IRI_QTY, vbUnicode)), "#0")
                                Else
                                    text(ptxIriSuu).text = ""
                                End If
                                
                                
                                
                                '2010.07.16
                                lblST_TANABAN(0).Caption = StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                            StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                                
                                                            '2012.01.30 引数追加
                                If GENSANKOKU_SET_Proc(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) Then
                                    Unload Me
                                End If
                                '2010.07.16
                                
                                
                                
                                '2010.10.07
                                text(ptxSHIIRE_WORK_CENTER).text = Trim(StrConv(ITEMREC.TORI_SHIIRE_WORK_CENTER, vbUnicode))
                                '2010.10.07
                                
                                
                                
                                
                                
'                                text(ptxMaiSuu).SetFocus
                                text(ptxIriSuu).SetFocus
                                Call Text_GotFocus(ptxIriSuu)
                  
'                                text(ptxHin_Nai).SetFocus   '2010.10.18
                  
                  
                            Case BtErrKeyNotFound
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> MT_2009.05.30
                                'MsgBox "入力したコードは登録されていません。"
                                'Text(0).SetFocus
                                If PN_CHK(text(Index), "G", "FLABEL", 1) Then
                                    text(Index).SetFocus
                                    Call Text_GotFocus(Index)
                                    Exit Sub
                                End If
                                
                                text(ptxHin_Name).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                text(ptxHin_Nai).text = StrConv(ITEMREC.HIN_NAI, vbUnicode)
                                
                                '2010.10.07
                                'text(ptxBikou).text = StrConv(ITEMREC.BIKOU, vbUnicode)
'                                If Trim(StrConv(ITEMREC.BIKOU20, vbUnicode)) = "" Or _
'                                    Mid(StrConv(ITEMREC.BIKOU20, vbUnicode), 1, 1) < " " Then
'
'                                    Call UniCode_Conv(ITEMREC.BIKOU20, StrConv(ITEMREC.BIKOU, vbUnicode))
'
'                                End If
                                text(ptxBikou).text = StrConv(ITEMREC.BIKOU20, vbUnicode)
                                '2010.10.07
                                If IsNumeric(StrConv(ITEMREC.IRI_QTY, vbUnicode)) Then
                                    text(ptxIriSuu).text = Format(CLng(StrConv(ITEMREC.IRI_QTY, vbUnicode)), "#0")
                                Else
                                    text(ptxIriSuu).text = ""
                                End If
                                
                                
                                
                                
                                
                                '2010.07.16
                                lblST_TANABAN(0).Caption = StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                            StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                                
                                
                                
                                                            '2012.01.30 引数追加
                                If GENSANKOKU_SET_Proc(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) Then
                                    Unload Me
                                End If
                                '2010.07.16
                                
                                
                                '2010.10.07
                                text(ptxSHIIRE_WORK_CENTER).text = Trim(StrConv(ITEMREC.TORI_SHIIRE_WORK_CENTER, vbUnicode))
                                '2010.10.07
                                
                                
                                
                                
'''                                text(ptxHin_Nai).SetFocus   '2010.10.18
                                
                                text(ptxIriSuu).SetFocus
                                Call Text_GotFocus(ptxIriSuu)
                                    
                                '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                Exit Sub
                            Case Else
                                
                                
                                If sts > 3000 Or sts = 3 Then
                                
                                    Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                
                                    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                                    If sts Then
                                        Call File_Error(sts, BtOpReset, "棚マスタ")
                                        Beep
                                        MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
                                    End If
                                
                
                                    Call File_Open_Proc
                                
                                    GoTo Item_Read
                
                                
                                End If
                                
                                
                                
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                Beep
                                MsgBox "システム異常が発生しました。処理を中止して下さい。"
                                Unload Me
                        End Select
                    End If
                
                
                
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 美的品番
                Case 2
                
                
                    If Trim(text(ptxHin_Gai).text) = "" Then
                                                
    
                        text(Index).text = RTrim(StrConv(text(Index).text, vbUpperCase))
                                                '美的品番で読み込み
                        Call UniCode_Conv(K1_B_ITEM.B_HIN_CODE, RTrim(text(ptxB_Hin_Code).text))
                        
B_Item_Read2:
                        sts = BTRV(BtOpGetEqual, B_ITEM_POS, B_ITEMREC, Len(B_ITEMREC), K1_B_ITEM, Len(K1_B_ITEM), 1)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                    
                                MsgBox "美的品番が登録されていません"
                                text(Index).SetFocus
                                Call Text_GotFocus(Index)
                                Exit Sub
                            
                            Case Else
                        
                        
                        
                                If sts > 3000 Or sts = 3 Then
                                
                                    Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                
                                    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                                    If sts Then
                                        Call File_Error(sts, BtOpReset, "")
                                        Beep
                                        MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
                                    End If
                                
                
                                    Call File_Open_Proc
                                
                                    GoTo B_Item_Read2
                
                                
                                End If
                                
                                
                                
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                Beep
                                MsgBox "システム異常が発生しました。処理を中止して下さい。"
                                Unload Me
                        
                        
                        
                        
                        End Select
                        
                        
                        
                        
                        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(B_ITEMREC.JGYOBU, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(B_ITEMREC.NAIGAI, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(B_ITEMREC.HIN_GAI, vbUnicode))
                        
Item_Read3:
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                                
                                text(ptxHin_Gai).text = RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                text(ptxHin_Name).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                text(ptxHin_Nai).text = RTrim(StrConv(ITEMREC.HIN_NAI, vbUnicode))
                                
                                
                                text(ptxBikou).text = RTrim(StrConv(ITEMREC.BIKOU20, vbUnicode))
                                
                                
                                
                                
                                If IsNumeric(StrConv(ITEMREC.IRI_QTY, vbUnicode)) Then
                                    text(ptxIriSuu).text = Format(CLng(StrConv(ITEMREC.IRI_QTY, vbUnicode)), "#0")
                                Else
                                    text(ptxIriSuu).text = ""
                                End If
                                lblST_TANABAN(0).Caption = StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                            StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                                
                                                            '2012.01.30 引数追加
                                If GENSANKOKU_SET_Proc(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) Then
                                    Unload Me
                                End If
                                '2010.07.16
                                
                                
                                
                                '2010.10.07
                                text(ptxSHIIRE_WORK_CENTER).text = Trim(StrConv(ITEMREC.TORI_SHIIRE_WORK_CENTER, vbUnicode))
                                '2010.10.07
                                
                                
                                
                                
                                
'                                text(ptxMaiSuu).SetFocus
                                text(ptxIriSuu).SetFocus
                                Call Text_GotFocus(ptxIriSuu)
                  
                  
                  
                            Case BtErrKeyNotFound
                                
                                MsgBox ("美的品番と品目の関連が異常です。確認して下さい")
                                text(ptxB_Hin_Code).SetFocus
                                Call Text_GotFocus(ptxB_Hin_Code)
                                Exit Sub
                            Case Else
                                
                                
                                If sts > 3000 Or sts = 3 Then
                                
                                    Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                
                                    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                                    If sts Then
                                        Call File_Error(sts, BtOpReset, "")
                                        Beep
                                        MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
                                    End If
                                
                
                                    Call File_Open_Proc
                                
                                    GoTo Item_Read3
                
                                
                                End If
                                
                                
                                
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                Beep
                                MsgBox "システム異常が発生しました。処理を中止して下さい。"
                                Unload Me
                        End Select
                    End If
                
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 美的品番
                
                Case 3
                    If Len(text(ptxHin_Gai).text) = 0 Then
                        If Len(text(ptxHin_Nai).text) <> 0 Then
                                                
                            text(Index).text = RTrim(StrConv(text(Index).text, vbUpperCase))
                                                
Item_Read2:
                                                '内部品番で読み込み
                            #If Center_chk Then
                                Call UniCode_Conv(K2_ITEM.JGYOBU, Last_JGYOBU)
                                If Combo(0).text = NAIGAI1$ Then
                                    Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI_NAI$)
                                Else
                                    Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI_GAI$)
                                End If
                                Call UniCode_Conv(K2_ITEM.HIN_NAI, RTrim(text(ptxHin_Nai).text))
                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                            #Else
                                Call UniCode_Conv(K1_ITEM.JGYOBU, Last_JGYOBU)
                                If Combo(0).text = NAIGAI1$ Then
                                    Call UniCode_Conv(K1_ITEM.NAIGAI, NAIGAI_NAI$)
                                Else
                                    Call UniCode_Conv(K1_ITEM.NAIGAI, NAIGAI_GAI$)
                                End If
                                Call UniCode_Conv(K1_ITEM.HIN_NAI, RTrim(text(ptxHin_Nai).text))
                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K1_ITEM, Len(K1_ITEM), 1)
                            #End If
                            Select Case sts
                                Case BtNoErr
                                    text(ptxHin_Gai).text = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                                    text(ptxHin_Name).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
'                                    text(ptxBikou).text = Left(StrConv(ITEMREC.BIKOU, vbUnicode), 10)
                                    text(ptxBikou).text = Left(StrConv(ITEMREC.BIKOU20, vbUnicode), 20)
                                    If IsNumeric(StrConv(ITEMREC.IRI_QTY, vbUnicode)) Then
                                        text(ptxIriSuu).text = Format(CLng(StrConv(ITEMREC.IRI_QTY, vbUnicode)), "#0")
                                    Else
                                        text(ptxIriSuu).text = ""
                                    End If
                                    
                                    
                                    '2010.07.16
                                    lblST_TANABAN(0).Caption = StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_DAN, vbUnicode)
                                    
                                                        '2012.01.30 引数追加
                                    If GENSANKOKU_SET_Proc(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) Then
                                        Unload Me
                                    End If
                                    '2010.07.16
                                    
                                    
                                    
                                    
                                    
                                    text(ptxMaiSuu).SetFocus

'''                                    text(ptxHin_Nai).SetFocus   '2010.10.18

                                Case BtErrKeyNotFound
                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> MT_2009.05.30
                                    'MsgBox "入力したコードは登録されていません。"
                                    'Text(0).SetFocus
                                    
                                    If PN_CHK(text(Index), "N", "FLABEL", 1, 1) Then
                                        text(Index).SetFocus
                                        Call Text_GotFocus(Index)
                                        Exit Sub
                                    End If
                                    text(ptxHin_Gai).text = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                                    text(ptxHin_Name).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                    text(ptxHin_Nai).text = StrConv(ITEMREC.HIN_NAI, vbUnicode)
                                    
'                                    text(ptxBikou).text = StrConv(ITEMREC.BIKOU, vbUnicode)
                                    text(ptxBikou).text = StrConv(ITEMREC.BIKOU20, vbUnicode)
                                    If IsNumeric(StrConv(ITEMREC.IRI_QTY, vbUnicode)) Then
                                        text(ptxIriSuu).text = Format(CLng(StrConv(ITEMREC.IRI_QTY, vbUnicode)), "#0")
                                    Else
                                        text(ptxIriSuu).text = ""
                                    End If
                                    
                                    
                                    
                                    
                                        
                                    '2010.07.16
                                    lblST_TANABAN(0).Caption = StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_DAN, vbUnicode)
                                                                    '2012.01.30 引数追加
                                    If GENSANKOKU_SET_Proc(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) Then
                                        Unload Me
                                    End If
                                    '2010.07.16
                                    
                                    
                                    
                                    text(ptxIriSuu).SetFocus
                                    Call Text_GotFocus(ptxIriSuu)
                                    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                
                                    Exit Sub
'''                                    text(ptxHin_Nai).SetFocus   '2010.10.18
                                Case Else
                                    
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.19
                                    If sts > 3000 Or sts = 3 Then
                                    
                                        Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                    
                                        sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                                        If sts Then
                                            Call File_Error(sts, BtOpReset, "棚マスタ")
                                            Beep
                                            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
                                        End If
                                    
                                    
                                        Call File_Open_Proc
                                    
                                        GoTo Item_Read2
                    
                                    
                                    End If
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.19
                                    
                                    
                                    
                                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                    Beep
                                    MsgBox "システム異常が発生しました。処理を中止して下さい。"
                                    Unload Me
                            End Select
                        Else
                            MsgBox "入力した項目はエラーです。"
                            Exit Sub
                        End If
                    End If
            End Select
            
            
            
            
            If Index < 3 Then
                text(ptxIriSuu).SetFocus
            End If
            If Index = ptxIriSuu Then
                text(ptxMaiSuu).SetFocus
            End If
            If Index > 2 Then
                If Index < 8 Then
                    text(Index + 1).SetFocus
                End If
            End If
       
'''             Call Tab_Ctrl(Shift)        '移動
       
        Case vbKeyUp
            For i = Index - 1 To 0 Step -1
                If text(i).Enabled Then
                    text(i).SetFocus
                    Exit For
                End If
            Next i
        Case vbKeyF1
            Command(0).Value = True
        Case vbKeyF9
            Command(8).Value = True
        Case vbKeyF12
            Command(11).Value = True
    End Select
End Sub


Private Sub Print_Sub_Proc()
                                            
Dim Gyo         As Integer
Dim wk_IRI_QTY  As String * 5
                                            
Dim wkGENSAN    As String * 15
                                            
'    Printer.NewPage
                                            
    On Error GoTo Err_Proc
                                            
    For Gyo = 0 To 5
                                            
                                            
        If Len(Trim(Print_tbl(Gyo, 0).HIN_GAI)) = 0 Then
            Exit For
        End If


'------------------------------------------------   1行目   ------------------
        With NormalFont
            .NAME = F1020591.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(20);
        
        Printer.Print "入庫現品票";
        Printer.Print Tab(47);
        Printer.Print Trim(JGYOBU_NAME);

        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            
            Printer.Print Tab(80);
            Printer.Print "入庫現品票";
            Printer.Print Tab(104);
            Printer.Print Trim(JGYOBU_NAME)
        End If
'------------------------------------------------   2行目   ------------------
        
        If Gyo = Max_Gyo Then
        
        
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 2
            End With
            Set Printer.Font = NormalFont
            Printer.Print
        Else
    
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 6
            End With
            Set Printer.Font = NormalFont
            Printer.Print
    
        End If
    
'------------------------------------------------   3行目   ------------------
        Set Printer.Font = Code39Font
        Printer.Print Tab(2);
        Printer.Print "*" + Trim(Left(Print_tbl(Gyo, 0).HIN_GAI, 14)) + "*";
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(23);
            Printer.Print "*" + Trim(Left(Print_tbl(Gyo, 1).HIN_GAI, 14)) + "*"
        End If
'------------------------------------------------   4行目   ------------------
        With NormalFont
            .NAME = F1020591.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   5行目   ------------------
       With NormalFont
            .NAME = F1020591.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F1020591.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
        Printer.Print "品番";
        With NormalFont
            .NAME = F1020591.FontName
            .Size = 14                          '18-14
        End With
        Set Printer.Font = NormalFont
        Printer.Print " " & Left(Print_tbl(Gyo, 0).HIN_GAI, 14);
        With NormalFont
            .NAME = F1020591.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
        Printer.Print "(" & Left(Print_tbl(Gyo, 0).HIN_NAI, 14) & ")";
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
            
            
            Printer.Print "品番";
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 14                      '18-14
            End With
            Set Printer.Font = NormalFont
            Printer.Print " " & Left(Print_tbl(Gyo, 1).HIN_GAI, 14);
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
            Printer.Print "(" & Left(Print_tbl(Gyo, 1).HIN_NAI, 14) & ")"
        End If
'------------------------------------------------   6行目   ------------------
        If Gyo = Max_Gyo Then
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 2
            End With
            Set Printer.Font = NormalFont
            Printer.Print
        Else
    
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 4
            End With
            Set Printer.Font = NormalFont
            Printer.Print
        End If
'------------------------------------------------   7行目   ------------------
        With NormalFont
            .NAME = F1020591.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F1020591.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
        Printer.Print "品名" & " " & LeftB(Print_tbl(Gyo, 0).HIN_NAME, 80);
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
            Printer.Print "品名" & " " & LeftB(Print_tbl(Gyo, 1).HIN_NAME, 80)
        End If
'------------------------------------------------   8行目   ------------------
        If Gyo = Max_Gyo Then
            
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 2
            End With
            Set Printer.Font = NormalFont
            Printer.Print
        Else
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 4
            End With
            Set Printer.Font = NormalFont
            Printer.Print
        End If
'------------------------------------------------   9行目   ------------------
        With NormalFont
            .NAME = F1020591.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        
        With NormalFont
            .NAME = F1020591.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print "　　入数" & ":";
        With NormalFont
            .NAME = F1020591.FontName
            .Size = 14                      '18-14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Format(Print_tbl(Gyo, 0).IRI_QTY, "#0");
        With NormalFont
            .NAME = F1020591.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(30);
        Printer.Print "入荷日" & ":";
        With NormalFont
            .NAME = F1020591.FontName
            .Size = 14                      '18-14
        End With
        Set Printer.Font = NormalFont
        Printer.Print text(ptxNyuka_YY).text & "/" & text(ptxNyuka_MM).text & "/" & text(ptxNyuka_DD).text;
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            Set Printer.Font = NormalFont
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 10
            End With
            Set Printer.Font = NormalFont
            
            Printer.Print "　　入数" & ":";
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 14                  '18-14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Format(Print_tbl(Gyo, 1).IRI_QTY, "#0");
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 10
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(88);
            Printer.Print "入荷日" & ":";
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 14                  '18-14
            End With
            Set Printer.Font = NormalFont
            Printer.Print text(ptxNyuka_YY).text & "/" & text(ptxNyuka_MM).text & "/" & text(ptxNyuka_DD).text
        End If
'------------------------------------------------   10行目   ------------------
        With NormalFont
            .NAME = F1020591.FontName
            .Size = 4
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   11行目   ------------------
        With NormalFont
            .NAME = F1020591.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F1020591.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print "標準棚番" & ":" & Print_tbl(Gyo, 0).ST_SOKO & "-" & Print_tbl(Gyo, 0).ST_RETU & "-" & Print_tbl(Gyo, 0).ST_REN & "-" & Print_tbl(Gyo, 0).ST_DAN;
        Printer.Print Tab(30);
        Printer.Print "　備考" & ":" & RTrim(LeftB(Print_tbl(Gyo, 0).BIKOU, 40));
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 10
            End With
            Printer.Print "標準棚番" & ":" & Print_tbl(Gyo, 1).ST_SOKO & "-" & Print_tbl(Gyo, 1).ST_RETU & "-" & Print_tbl(Gyo, 1).ST_REN & "-" & Print_tbl(Gyo, 1).ST_DAN;
            Printer.Print Tab(88);
            Printer.Print "　備考" & ":" & RTrim(LeftB(Print_tbl(Gyo, 1).BIKOU, 40))
        End If
'------------------------------------------------   12行目   ------------------
        With NormalFont
            .NAME = F1020591.FontName
                .Size = 14                      '14-->10                                    '2013.12.06
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F1020591.FontName
            .Size = 10                          '10-8
        End With
        Set Printer.Font = NormalFont
        
        
        
        wkGENSAN = Left(Print_tbl(Gyo, 0).GENSAN, 13) & Right(Print_tbl(Gyo, 0).GENSAN, 2)
        
        
        
'        Printer.Print "　原産国" & ":" & Left(Print_tbl(Gyo, 0).GENSAN, 15);
        Printer.Print "　原産国" & ":" & wkGENSAN;
        
        
        
        Printer.Print Tab(30);
        Printer.Print "仕入先" & ":" & Print_tbl(Gyo, 0).SHIIRE_WORK_CENTER;
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 14                      '14-->10                                    '2013.12.06
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            With NormalFont
                .NAME = F1020591.FontName
                .Size = 10                       '10-8
            End With
'            Printer.Print "　原産国" & ":" & Left(Print_tbl(Gyo, 1).GENSAN, 15);
            
        wkGENSAN = Left(Print_tbl(Gyo, 1).GENSAN, 13) & Right(Print_tbl(Gyo, 1).GENSAN, 2)
'        Printer.Print "　原産国" & ":" & Left(Print_tbl(Gyo, 0).GENSAN, 15);
        Printer.Print "　原産国" & ":" & wkGENSAN;
            
            
            Printer.Print Tab(88);
'            Printer.Print "仕入先" & ":" & Print_tbl(Gyo, 1).SHIIRE_WORK_CENTER;   '2013.12.06
            Printer.Print "仕入先" & ":" & Print_tbl(Gyo, 1).SHIIRE_WORK_CENTER     '2013.12.06
        End If




'------------------------------------------------   13行目   ------------------


'2013.12.06        With NormalFont
'2013.12.06            .NAME = F1020591.FontName
'2013.12.06            .Size = 10  '8->10  '2013.10.22
'2013.12.06        End With
'2013.12.06        Set Printer.Font = NormalFont
'2013.12.06        Printer.Print
        
        
        With NormalFont                                                                         '2013.10.22
            .NAME = F1020591.FontName                                                           '2013.10.22
            .Size = 10 '8-7 2013.12.06 7->12 2013.12.18                                         '2013.10.22
        End With                                                                                '2013.10.22
        Set Printer.Font = NormalFont                                                           '2013.10.22
        
'201312.06        Printer.Print Tab(2);                                                                   '2013.10.22
        Printer.Print Print_tbl(Gyo, 0).B_HIN_CODE;                                             '2013.10.22
        
                
        '41桁まで印刷 2013.12.18
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(58); '74 - 80 2013.12.06 80-->53 2013.12.18                                                       '2013.10.22
            Printer.Print Print_tbl(Gyo, 1).B_HIN_CODE                                          '2013.10.22
        End If
'------------------------------------------------   1行目   ------------------
'        Set Printer.Font = Code39Font
'        Printer.Print Tab(2);
'        Printer.Print "*" + Print_tbl(Gyo, 0).HIN_GAI + "*";
'        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
'            Printer.Print
'        Else
'            Printer.Print Tab(20);
'            Printer.Print "*" + Print_tbl(Gyo, 1).HIN_GAI + "*"
'        End If
'------------------------------------------------   2行目   ------------------
'        With NormalFont
'            .NAME = F1020501.FontName
'            .Size = 14
'        End With
'        Set Printer.Font = NormalFont
'        Printer.Print Tab(4);
'        Printer.Print "[" & Trim(JGYOBU_NAME) & "]";
'        With NormalFont
'            .NAME = F1020501.FontName
'            .Size = 12
'        End With
'        Set Printer.Font = NormalFont
'        Printer.Print Tab(18);
'        Printer.Print "[" & Print_tbl(Gyo, 0).NAIGAI & "]";
'        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
'            Printer.Print
'        Else
'            Printer.Print Tab(53);
'            With NormalFont
'                .NAME = F1020501.FontName
'                .Size = 14
'            End With
'            Set Printer.Font = NormalFont
'            Printer.Print "[" & Trim(JGYOBU_NAME) & "]";
'            With NormalFont
'                .NAME = F1020501.FontName
'                .Size = 12
'            End With
'            Set Printer.Font = NormalFont
'            Printer.Print Tab(67);
'            Printer.Print "[" & Print_tbl(Gyo, 1).NAIGAI & "]"
'        End If
''2010.07.21        Printer.Print
'------------------------------------------------   3行目   ------------------
'        Printer.Print Tab(4);
'        Printer.Print "[入庫現品票]" & "          ";
'        Printer.Print text(ptxNyuka_YY).text & "/" & text(ptxNyuka_MM).text & "/" & text(ptxNyuka_DD).text;
'        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
'            Printer.Print
'        Else
'            Printer.Print Tab(53);
'            Printer.Print "[入庫現品票]" & "          ";
'            Printer.Print text(ptxNyuka_YY).text & "/" & text(ptxNyuka_MM).text & "/" & text(ptxNyuka_DD).text
'        End If
'------------------------------------------------   4行目   ------------------
'        With NormalFont
'            .NAME = F1020501.FontName
'            .Size = 14
'        End With
'        Set Printer.Font = NormalFont
'        Printer.Print Tab(4);
'        Printer.Print "品番" & "  ";
'        Printer.Print Print_tbl(Gyo, 0).HIN_GAI & " (";
'        Printer.Print Print_tbl(Gyo, 0).HIN_NAI & ")";
'        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
'            Printer.Print
'        Else
'            Printer.Print Tab(46);
'            Printer.Print "品番" & "  ";
'            Printer.Print Print_tbl(Gyo, 1).HIN_GAI & " (";
'            Printer.Print Print_tbl(Gyo, 1).HIN_NAI & ")"
'        End If
'------------------------------------------------   5行目   ------------------
'        With NormalFont
'            .NAME = F1020501.FontName
'            .Size = 12
'        End With
'        Set Printer.Font = NormalFont
'        Printer.Print Tab(4);
'        Printer.Print "品名  ";
'        Printer.Print Print_tbl(Gyo, 0).HIN_NAME;
'        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
'            Printer.Print
'        Else
'            Printer.Print Tab(53);
'            Printer.Print "品名  ";
'            Printer.Print Print_tbl(Gyo, 1).HIN_NAME
'        End If
'------------------------------------------------   6行目   ------------------
'        Printer.Print Tab(13);
'        Printer.Print "入数：";
'        If IsNumeric(Print_tbl(Gyo, 0).IRI_QTY) Then
'            wk_IRI_QTY = Right(Format(CLng(Print_tbl(Gyo, 0).IRI_QTY), "###0"), 5)
'            wk_IRI_QTY = Space(Len(wk_IRI_QTY) - Len(Trim(wk_IRI_QTY))) & Trim(wk_IRI_QTY)
'
'            Printer.Print StrConv(wk_IRI_QTY, vbWide);
'        Else
'            Printer.Print "　　　　　";
'        End If
'        Printer.Print "  " & Print_tbl(Gyo, 0).BIKOU;
'        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
'            Printer.Print
'        Else
'            Printer.Print Tab(62);
'            Printer.Print "入数：";
'            If IsNumeric(Print_tbl(Gyo, 1).IRI_QTY) Then
'                wk_IRI_QTY = Right(Format(CLng(Print_tbl(Gyo, 1).IRI_QTY), "###0"), 5)
'                wk_IRI_QTY = Space(Len(wk_IRI_QTY) - Len(Trim(wk_IRI_QTY))) & Trim(wk_IRI_QTY)
'
'                Printer.Print StrConv(wk_IRI_QTY, vbWide);
'            Else
'                Printer.Print "　　　　　";
'            End If
'            Printer.Print "  " & Print_tbl(Gyo, 1).BIKOU
'        End If
'------------------------------------------------   6行目   ------------------
'        Printer.Print Tab(4);
'        Printer.Print "標準入庫棚  ";
'        Printer.Print Print_tbl(Gyo, 0).ST_SOKO & ":";
'        Printer.Print Print_tbl(Gyo, 0).ST_SOKO_NAME;
'        Printer.Print Tab(37);
'        Printer.Print Print_tbl(Gyo, 0).ST_RETU & "-" & Print_tbl(Gyo, 0).ST_REN & "-" & Print_tbl(Gyo, 0).ST_DAN;
'        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
'            Printer.Print
'        Else
'            Printer.Print Tab(53);
'            Printer.Print "標準入庫棚  ";
'            Printer.Print Print_tbl(Gyo, 1).ST_SOKO & ":";
'            Printer.Print Print_tbl(Gyo, 1).ST_SOKO_NAME;
'            Printer.Print Tab(86);
'            Printer.Print Print_tbl(Gyo, 1).ST_RETU & "-" & Print_tbl(Gyo, 1).ST_REN & "-" & Print_tbl(Gyo, 1).ST_DAN
'        End If
'
'
'
'------------------------------------------------   7行目   ------------------
'        Printer.Print Tab(4);
'        Printer.Print "　　原産国  ";
'        Printer.Print Print_tbl(Gyo, 0).GENSAN;
'        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
'            Printer.Print ;
'        Else
'            Printer.Print Tab(53);
'            Printer.Print "　　原産国  ";
'            Printer.Print Print_tbl(Gyo, 1).GENSAN;
'        End If
'
'
'
        If Gyo <> Max_Gyo Then


            With NormalFont
                .NAME = F1020591.FontName
                .Size = 8        '2013.12.19 10-->6
            End With
            Set Printer.Font = NormalFont
            Printer.Print
            
            



            If Max_Gyo <> 2 Then
            
                With NormalFont
                    .NAME = F1020591.FontName
                    .Size = 4      '6-->4　2013.12.19
                End With
                Set Printer.Font = NormalFont
                Printer.Print '2013.10.23-->12.19
                
                
                With NormalFont
                    .NAME = F1020591.FontName
                    .Size = 2      '6-->4　2013.12.19
                End With
                Printer.Print '2013.10.23-->12.19
'                Printer.Print '2013.10.23-->12.19
                
                
                
                
            Else
                With NormalFont
                    .NAME = F1020591.FontName
                    .Size = 4
                End With
                Set Printer.Font = NormalFont
                Printer.Print
                With NormalFont
                    .NAME = F1020591.FontName
                    .Size = 6
                End With
                Set Printer.Font = NormalFont
                Printer.Print
            
            
            End If

'            With NormalFont
'                .NAME = F1020501.FontName
'                .Size = 14
'            End With
'            Set Printer.Font = NormalFont
'            Printer.Print
''        With NormalFont
''            .NAME = F1020501.FontName
''            .Size = 18
''        End With
''        Set Printer.Font = NormalFont
'            With NormalFont
'                .NAME = F1020501.FontName
'                .Size = 18
'            End With
'            Set Printer.Font = NormalFont
'            Printer.Print
'
'
'
''2010.07.21
'            With NormalFont
'                .NAME = F1020501.FontName
'                .Size = 14
'            End With
'            Set Printer.Font = NormalFont
'            Printer.Print
''2010.07.21


        End If
    Next Gyo

    Exit Sub

Err_Proc:

    If Err.Number = 482 Then
        MsgBox "プリンターエラーが発生しました。"
    Else
        MsgBox "実行時エラー：" & Err.Number
    End If
End Sub

Private Function Item_Update_Proc() As Integer

Dim sts         As Integer
Dim ans         As Integer
Dim wk_Naigai   As String * 1

    Item_Update_Proc = True

    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    
    If Combo(0).text = NAIGAI1 Then
        wk_Naigai = NAIGAI_NAI
    Else
        wk_Naigai = NAIGAI_GAI
    End If
    
    Call UniCode_Conv(K0_ITEM.NAIGAI, wk_Naigai)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, text(ptxHin_Gai).text)
Item_Read:
    Do
        sts = BTRV(BtOpGetEqual + 200, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                MsgBox "他でデータ変更されています。更新処理を中止します。"
                Item_Update_Proc = False
                Exit Function
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Item_Update_Proc = False
                    Exit Function
                End If
            Case Else
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.19
                If sts > 3000 Or sts = 3 Then
                
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                
                    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    If sts Then
                        Call File_Error(sts, BtOpReset, "棚マスタ")
                        Beep
                        MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
                    End If
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.31
'
'                                                '倉庫マスタＯＰＥＮ
'                    If SOKO_Open(0) Then
'                        Beep
'                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'                        Unload Me
'                    End If
'                                                '品目マスタＯＰＥＮ
'                    If ITEM_Open(0) Then
'                        Beep
'                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'                        Unload Me
'                    End If
'
'                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.05.30
'                                                'PNマスタＯＰＥＮ
'                    If PN_M_Open(0) Then
'                        Beep
'                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'                        Unload Me
'                    End If
'                    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'                                                '原産国マスタＯＰＥＮ
'                    If GENSAN_Open(0) Then
'                        Beep
'                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'                        Unload Me
'                    End If
                
                
                    Call File_Open_Proc
                
                    GoTo Item_Read

                
                End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.19
                
                
                
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                MsgBox "システム異常が発生しました。処理を中止して下さい。"
                Exit Function
        End Select
    Loop


    Call UniCode_Conv(ITEMREC.HIN_NAI, text(ptxHin_Nai).text)

    Call UniCode_Conv(ITEMREC.BIKOU, "")
    Call UniCode_Conv(ITEMREC.BIKOU20, text(ptxBikou).text)
    
    
    If text(ptxIriSuu).text = "" Then
        Call UniCode_Conv(ITEMREC.IRI_QTY, "")
    Else
        If Len(Trim(text(ptxIriSuu).text)) = 0 Then
            Call UniCode_Conv(ITEMREC.IRI_QTY, "")
        Else
            Call UniCode_Conv(ITEMREC.IRI_QTY, Format(CLng(text(ptxIriSuu).text), "00000000"))
        End If
    End If


    Call UniCode_Conv(ITEMREC.UPD_TANTO, "2050")                            '追加　担当者

    Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))  '追加　日時



    Do
        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Item_Update_Proc = False
                    Exit Function
                End If
            Case Else
                


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.19
                If sts > 3000 Or sts = 3 Then
                
                    Call File_Error(sts, BtOpUpdate, "品目マスタ", 0)

                    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    If sts Then
                        Call File_Error(sts, BtOpReset, "棚マスタ")
                        Beep
                        MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
                    End If
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.31
'                                                '倉庫マスタＯＰＥＮ
'                    If SOKO_Open(0) Then
'                        Beep
'                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'                        Unload Me
'                    End If
'                                                '品目マスタＯＰＥＮ
'                    If ITEM_Open(0) Then
'                        Beep
'                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'                        Unload Me
'                    End If
'
'                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.05.30
'                                                'PNマスタＯＰＥＮ
'                    If PN_M_Open(0) Then
'                        Beep
'                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'                        Unload Me
'                    End If
'                    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'                                                '原産国マスタＯＰＥＮ
'                    If GENSAN_Open(0) Then
'                        Beep
'                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'                        Unload Me
'                    End If
                    Call File_Open_Proc
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.31
                
                    GoTo Item_Read

                
                End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.19
                
                
                
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                MsgBox "システム異常が発生しました。処理を中止して下さい。"
                Exit Function
        End Select
    Loop

    Item_Update_Proc = False


End Function


Private Sub Text_LostFocus(Index As Integer)

    If Index = 0 Or Index = 2 Or Index = 3 Then
    
        text(Index).text = RTrim(StrConv(text(Index).text, vbUpperCase))
    
    
    End If


End Sub


Private Function GENSANKOKU_SET_Proc(TORI_GENSANKOKU As String) As Integer
'   2012.01.30 引数追加
Dim sts     As Integer
Dim com     As Integer
Dim i       As Integer

Dim wkTORI_GENSANKOKU   As String   '2013.03.31


    GENSANKOKU_SET_Proc = True
    
    'NULL 空白変換  2013.03.31
    wkTORI_GENSANKOKU = ""
    For i = 1 To Len(TORI_GENSANKOKU)
                
        If Mid(TORI_GENSANKOKU, i, 1) < " " Then
            wkTORI_GENSANKOKU = wkTORI_GENSANKOKU & " "
        Else
            wkTORI_GENSANKOKU = wkTORI_GENSANKOKU & Mid(TORI_GENSANKOKU, i, 1)
        End If
    
    Next i
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    TORI_GENSANKOKUの有無チェック＆書き込み   2012.01.31
    
    If Trim(wkTORI_GENSANKOKU) = "" Then                '2013.03.31
    Else
        Call UniCode_Conv(K0_GENSAN.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_GENSAN.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_GENSAN.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
        'Call UniCode_Conv(K0_GENSAN.GENSANKOKU, TORI_GENSANKOKU)           '2013.03.31
        Call UniCode_Conv(K0_GENSAN.GENSANKOKU, wkTORI_GENSANKOKU)          '2013.03.31
        
        sts = BTRV(BtOpGetEqual, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
            
                Call UniCode_Conv(GENSANREC.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(GENSANREC.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(GENSANREC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                'Call UniCode_Conv(GENSANREC.GENSANKOKU, TORI_GENSAKOKU)        '2013.03.31
                Call UniCode_Conv(GENSANREC.GENSANKOKU, wkTORI_GENSANKOKU)      '2013.03.31
                Call UniCode_Conv(GENSANREC.FILLER, "")
        
                Call UniCode_Conv(GENSANREC.INS_TANTO, "PLABEL")
                Call UniCode_Conv(GENSANREC.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))
        
                Call UniCode_Conv(GENSANREC.UPD_TANTO, "")
                Call UniCode_Conv(GENSANREC.UPD_DATETIME, "")
            
            
                sts = BTRV(BtOpInsert, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                Select Case sts
                
                    Case BtNoErr
                    Case BtErrDuplicates
                    Case Else
                        Call File_Error(sts, com, "原産国マスタ")
                        Exit Function
                End Select
            
            
            
            
            Case Else
                Exit Function
        End Select
    End If
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    TORI_GENSANKOKUの有無チェック＆書き込み   2012.01.31
    
    
    
    
    
    
    
    
    Combo(1).Clear
    List2.Clear
    
    
    
    Call UniCode_Conv(K0_GENSAN.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_GENSAN.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_GENSAN.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))

    Call UniCode_Conv(K0_GENSAN.GENSANKOKU, "")
    
    com = BtOpGetGreater
    
    Do
        
        DoEvents
        
        sts = BTRV(com, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> StrConv(GENSANREC.JGYOBU, vbUnicode) Or _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) <> StrConv(GENSANREC.NAIGAI, vbUnicode) Or _
                    StrConv(ITEMREC.HIN_GAI, vbUnicode) <> StrConv(GENSANREC.HIN_GAI, vbUnicode) Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Exit Function
        End Select
    
        
'        List2.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)        2013.01.28
        
        If Trim(StrConv(GENSANREC.UPD_DATETIME, vbUnicode)) = "" Then                                               '2013.01.28
            List2.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2013.01.28
        Else                                                                                                        '2013.01.28
            List2.AddItem StrConv(GENSANREC.UPD_DATETIME, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2013.01.28
        End If
        
        com = BtOpGetNext
    Loop
    
        
    If List2.ListCount > 0 Then
'''        For i = 0 To List2.ListCount - 1
        For i = List2.ListCount - 1 To 0 Step -1
            Combo(1).AddItem Right(List2.List(i), 20)
        
        Next i
    
        Combo(1).ListIndex = 0
    End If
    
    GENSANKOKU_SET_Proc = False


End Function

Private Sub File_Open_Proc()
                                
Dim c           As String * 128     '2013.8.23
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenRead) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.05.30
                                'PNマスタＯＰＥＮ
    If PN_M_Open(BtOpenRead) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                '原産国マスタＯＰＥＮ
    If GENSAN_Open(BtOpenNomal) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If


    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<    2012.01.31  -->　削除 2012.02.06
'                                'カントリーマスタ ＯＰＥＮ
'    If Country_Open(BtOpenRead) Then
'        Unload Me
'    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<    2012.01.31  -->　削除 2012.02.06


    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                '美的品番データＯＰＥＮ
    If B_ITEM_Open(BtOpenNomal) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If

End Sub
