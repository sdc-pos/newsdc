VERSION 5.00
Begin VB.Form F1020501 
   BackColor       =   &H00FFFFFF&
   Caption         =   "入庫現品票印刷"
   ClientHeight    =   10230
   ClientLeft      =   2025
   ClientTop       =   2940
   ClientWidth     =   17175
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
   ScaleHeight     =   10230
   ScaleWidth      =   17175
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   135
      Left            =   7680
      TabIndex        =   72
      Top             =   1320
      Width           =   15
   End
   Begin VB.ListBox List3 
      Height          =   300
      Left            =   13080
      TabIndex        =   71
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox text 
      Height          =   375
      Index           =   12
      Left            =   9480
      MaxLength       =   40
      TabIndex        =   70
      Top             =   2520
      Width           =   4935
   End
   Begin VB.TextBox text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   3810
      MaxLength       =   4
      TabIndex        =   8
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox text 
      Height          =   375
      Index           =   11
      Left            =   9450
      MaxLength       =   8
      TabIndex        =   11
      Top             =   2940
      Width           =   1092
   End
   Begin VB.ListBox List2 
      Height          =   300
      Left            =   6435
      Sorted          =   -1  'True
      TabIndex        =   50
      Top             =   7740
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      ItemData        =   "F1020501.frx":0000
      Left            =   9420
      List            =   "F1020501.frx":0002
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   5
      Top             =   1560
      Width           =   2790
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   2880
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   13
      Top             =   8280
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "用紙選択"
      Height          =   975
      Left            =   450
      TabIndex        =   17
      Top             =   420
      Width           =   1575
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A4"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A5"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   10
      Left            =   4410
      MaxLength       =   3
      TabIndex        =   46
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   9
      Left            =   10080
      MaxLength       =   3
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   8280
      Width           =   732
   End
   Begin VB.TextBox text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   9450
      MaxLength       =   8
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   5730
      MaxLength       =   2
      TabIndex        =   10
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   4890
      MaxLength       =   2
      TabIndex        =   9
      Top             =   2640
      Width           =   375
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   2235
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   7740
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox text 
      Height          =   375
      Index           =   4
      Left            =   9450
      MaxLength       =   20
      TabIndex        =   7
      Top             =   2040
      Width           =   2532
   End
   Begin VB.TextBox text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   3810
      MaxLength       =   3
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3840
      Width           =   16770
   End
   Begin VB.TextBox text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   3810
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1020
      Width           =   2655
   End
   Begin VB.TextBox text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   9450
      MaxLength       =   40
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   3135
   End
   Begin VB.TextBox text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   3810
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
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   8880
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
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   8880
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
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   8880
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
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   8880
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   8880
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
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   8880
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   8880
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
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8880
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8880
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   8880
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   8880
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   8880
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "備　考２"
      Height          =   255
      Index           =   29
      Left            =   8400
      TabIndex        =   69
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblMAISUU 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   12240
      TabIndex        =   68
      Top             =   8400
      Width           =   120
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "シート"
      Height          =   255
      Index           =   28
      Left            =   11400
      TabIndex        =   67
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label lblSIZE 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12480
      TabIndex        =   66
      Top             =   8040
      Width           =   975
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   27
      Left            =   12360
      TabIndex        =   65
      Top             =   9120
      Width           =   120
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ラベル紙"
      Height          =   255
      Index           =   26
      Left            =   11400
      TabIndex        =   64
      Top             =   8040
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "原産国"
      Height          =   255
      Index           =   25
      Left            =   14280
      TabIndex        =   63
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品　　　　名"
      Height          =   255
      Index           =   24
      Left            =   9240
      TabIndex        =   62
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品番(内部)"
      Height          =   255
      Index           =   23
      Left            =   2640
      TabIndex        =   61
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品番(外部)"
      Height          =   255
      Index           =   21
      Left            =   240
      TabIndex        =   60
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "選択プリンター"
      Height          =   255
      Index           =   19
      Left            =   1080
      TabIndex        =   59
      Top             =   8400
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "※ダブルクリックで行削除"
      Height          =   255
      Left            =   13680
      TabIndex        =   58
      Top             =   7200
      Width           =   3015
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "個"
      Height          =   252
      Index           =   22
      Left            =   5160
      TabIndex        =   57
      Top             =   3240
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label lblKEPPIN_QTY 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   4800
      TabIndex        =   56
      Top             =   3240
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "件"
      Height          =   252
      Index           =   20
      Left            =   4320
      TabIndex        =   55
      Top             =   3240
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label lblKEPPIN_CNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   3840
      TabIndex        =   54
      Top             =   3240
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "欠品"
      Height          =   252
      Index           =   18
      Left            =   2616
      TabIndex        =   53
      Top             =   3240
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品名"
      Height          =   255
      Index           =   17
      Left            =   8850
      TabIndex        =   52
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "仕入先"
      Height          =   255
      Index           =   16
      Left            =   8610
      TabIndex        =   51
      Top             =   3060
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "原産国"
      Height          =   255
      Index           =   15
      Left            =   8595
      TabIndex        =   49
      Top             =   1680
      Width           =   750
   End
   Begin VB.Label lblST_TANABAN 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3810
      TabIndex        =   48
      Top             =   1620
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "標 準 棚 番"
      Height          =   255
      Index           =   13
      Left            =   2250
      TabIndex        =   47
      Top             =   1620
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "現品票枚数合計"
      Height          =   255
      Index           =   12
      Left            =   8280
      TabIndex        =   45
      Top             =   8400
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "印刷備考"
      Height          =   255
      Index           =   11
      Left            =   6480
      TabIndex        =   44
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "枚数"
      Height          =   255
      Index           =   10
      Left            =   5610
      TabIndex        =   43
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "入数"
      Height          =   255
      Index           =   9
      Left            =   4770
      TabIndex        =   42
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "入数"
      Height          =   255
      Index           =   8
      Left            =   8850
      TabIndex        =   41
      Top             =   1020
      Width           =   495
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "日"
      Height          =   255
      Index           =   7
      Left            =   6210
      TabIndex        =   40
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "月"
      Height          =   255
      Index           =   6
      Left            =   5370
      TabIndex        =   39
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "年"
      Height          =   255
      Index           =   5
      Left            =   4530
      TabIndex        =   38
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "　 入 荷 日"
      Height          =   255
      Index           =   4
      Left            =   2250
      TabIndex        =   37
      Top             =   2760
      Width           =   1455
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
      Left            =   120
      TabIndex        =   36
      Top             =   9480
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   255
      Index           =   3
      Left            =   1350
      TabIndex        =   35
      Top             =   7860
      Visible         =   0   'False
      Width           =   855
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
      Height          =   255
      Left            =   180
      TabIndex        =   34
      Top             =   8340
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "備　考１"
      Height          =   255
      Index           =   2
      Left            =   8370
      TabIndex        =   33
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "印 刷 枚 数"
      Height          =   255
      Index           =   1
      Left            =   2250
      TabIndex        =   32
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品番（内部）"
      Height          =   255
      Index           =   14
      Left            =   2250
      TabIndex        =   31
      Top             =   1140
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品番（外部）"
      Height          =   255
      Index           =   0
      Left            =   2250
      TabIndex        =   30
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
Attribute VB_Name = "F1020501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NormalFont As New StdFont           '印刷フォント
Dim Code39Font As New StdFont           '印刷フォント

Private WS_NO       As String           'ﾜｰｸｽﾃｰｼｮﾝ番号  2016.12.28


Private Type Print_tbl_tag              '印刷用テーブル

    NAIGAI          As String * 2
    HIN_GAI         As String * 20
'    HIN_NAI         As String * 13     '2018.09.21
    HIN_NAI         As String * 20      '2018.09.21
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
    BIKOU2          As String * 40      '2019.01.21
    GENSAN          As String * 22
'2010.10.07
    SHIIRE_WORK_CENTER As _
                       String * 8

    KEPPIN_QTY      As String * 8       '2013.08.23


    GAI_BUHIN       As String * 1       '2017.03.03
End Type

Dim Print_tbl(0 To 6, 0 To 1) _
                    As Print_tbl_tag




Dim HIN_GAI_LTRIM   As Integer          '2016.12.27
Dim HIN_NAI_LTRIM   As Integer          '2017.01.10



Private MENU_NO     As String * 2       'ﾒﾆｭｰ�ａ@   2016.12.27
Private RIRK_ID     As String * 2       '要因　     2016.12.27
Private MEMO        As String           'メモ       2016.12.27




Private Const ptxHin_Gai% = 0       '品番(外)
Private Const ptxHin_Name% = 1      '品名
Private Const ptxHin_Nai% = 2       '品番(内)




Private Const ptxMaiSuu% = 3        '印刷枚数
Private Const ptxBikou% = 4         '印刷備考
Private Const ptxNyuka_YY% = 5      '入荷日　年
Private Const ptxNyuka_MM% = 6      '入荷日　月
Private Const ptxNyuka_DD% = 7      '入荷日　日
Private Const ptxIriSuu% = 8        '入り数
Private Const ptxGoukei% = 9        '合計

Private Const ptxwkMaiSuu% = 10     '保存用印刷枚数

                                    '保存用印刷枚数 2010.10.07
Private Const ptxSHIIRE_WORK_CENTER% = 11

Private Const ptxBikou2% = 12         '印刷備考2 '2018.02.03
                                    
                                    
                                    '2010.10.07
Private SHIIRE_WORK_CENTER_F  As Integer


Dim JGYOBU_NAME As String

Dim Printer_tbl() As String
Dim Max_Gyo     As Integer

Dim Last_Printer    As String   '2016.09.30


Dim GENSAN_KOKU_F   As Integer  '2017.03.03

Dim CLEAR_BUTTON    As Integer  '2018.12.04



Dim Print_Flg       As Integer  '2019.01.21


'Private Const Last_Update_Day$ = "(F102050 2019.04.02 07:45)"
'Private Const Last_Update_Day$ = "(F102050 2019.04.05 13:45)"
'Private Const Last_Update_Day$ = "(F102050 2019.06.20 15:45)"
'Private Const Last_Update_Day$ = "(F102050 2019.08.26 10:30) 対内品番()削除"
Private Const Last_Update_Day$ = "(F102050 2019.11.08 14:30) 対外品番16桁対応"

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


    On Error GoTo Error_Proc        '2018.10.24



    If Option1(0).Value = True Then
        Printer.PaperSize = vbPRPSA5
        Printer.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷
        Max_Gyo = 2
    Else
        Printer.PaperSize = vbPRPSA4
        Printer.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷
        Max_Gyo = 5
    End If

    On Error GoTo 0                 '2018.10.24
    



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
        'wk_Naigai = Left(List1.List(wk_LOOP), 1)           '2016.10.17
        wk_Naigai = NAIGAI_NAI                              '2016.10.17
        
Item_Read:
        
        Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_ITEM.NAIGAI, wk_Naigai)
'        Call UniCode_Conv(K0_ITEM.HIN_GAI, Mid(List1.List(wk_LOOP), 3, 20))        '2016.10.17
        Call UniCode_Conv(K0_ITEM.HIN_GAI, Mid(List1.List(wk_LOOP), 1, 20))         '2016.10.17
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

                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2015.04.20
                    'sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    'If sts Then
                    '    Call File_Error(sts, BtOpReset, "棚マスタ")
                    '    Beep
                    '    MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
                    'End If
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
                    
                    Do  '2015.04.20
                    
                        If Not File_Open_Proc Then
                            Exit Do
                        End If
                    Loop
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.31
                
                
                
                
                    GoTo Item_Read

                
                End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.19
                
                
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        End Select
        
        
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  作業ﾛｸﾞ出力    '2016.12.27
        If Trim(MENU_NO) <> "" Then
        
            If P_SAGYO_LOG_OUTPUT_PROC(WS_NO, _
                                        "", _
                                        Last_JGYOBU, _
                                        "1", _
                                        MENU_NO, _
                                        RIRK_ID, _
                                        Mid(List1.List(wk_LOOP), 1, 20), _
                                        0, _
                                        CLng(CInt(Mid(List1.List(wk_LOOP), 47, 3))), _
                                        "", _
                                        "", , , , , , , , , , MEMO) Then
                Exit Function
            End If
    
        End If
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  作業ﾛｸﾞ出力    '2016.12.27
        
        
        
        
        
'        For Maisu = 1 To CInt(Mid(List1.List(wk_LOOP), 49, 3))     '2016.10.17
        For Maisu = 1 To CInt(Mid(List1.List(wk_LOOP), 47, 3))      '2016.10.17
            
            DoEvents
            
            If wk_Naigai = NAIGAI_NAI Then
                Print_tbl(Gyo, Retu).NAIGAI = NAIGAI1
            Else
                Print_tbl(Gyo, Retu).NAIGAI = NAIGAI2
            End If
'            Print_tbl(Gyo, Retu).HIN_GAI = Mid(List1.List(wk_LOOP), 3, 20)         '2016.10.17
            Print_tbl(Gyo, Retu).HIN_GAI = Mid(List1.List(wk_LOOP), 1, 20)          '2016.10.17
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
'                Print_tbl(Gyo, Retu).SHIIRE_WORK_CENTER = Trim(Right(List1.List(wk_LOOP), 8))              '2013.08.23
                Print_tbl(Gyo, Retu).SHIIRE_WORK_CENTER = Trim(Left(Right(List1.List(wk_LOOP), 16), 8))     '2013.08.23
                '2010.10.07
                Print_tbl(Gyo, Retu).KEPPIN_QTY = Trim(Right(List1.List(wk_LOOP), 8))                        '2013.08.23
            
                Print_tbl(Gyo, Retu).GAI_BUHIN = StrConv(ITEMREC.GAI_BUHIN, vbUnicode)      '2017.03.03
'
            
            
                '>>>>>>>>>  2018.02.03
                
                Call UniCode_Conv(K0_ITEM_CHG.N_JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K0_ITEM_CHG.N_NAIGAI, "1")
                Call UniCode_Conv(K0_ITEM_CHG.N_HIN_GAI, Print_tbl(Gyo, Retu).HIN_GAI)
                
                sts = BTRV(BtOpGetEqual, ITEM_CHG_POS, ITEM_CHG_REC, Len(ITEM_CHG_REC), K0_ITEM_CHG, Len(K0_ITEM_CHG), 0)
                Select Case sts
                    
                    Case BtNoErr
                        Print_tbl(Gyo, Retu).BIKOU2 = StrConv(ITEM_CHG_REC.O_HIN_GAI, vbUnicode)
                    Case BtErrKeyNotFound
                        Print_tbl(Gyo, Retu).BIKOU2 = " "
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目読み替え")
                        Unload Me
                End Select
                '>>>>>>>>>  2018.02.03
            
            
            
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
            
                Print_tbl(Gyo, Retu).KEPPIN_QTY = ""        '2013.08.23
                
                Print_tbl(Gyo, Retu).GAI_BUHIN = ""         '2017.03.03
            
            
            End If
    
            'Print_tbl(Gyo, Retu).IRI_QTY = Mid(List1.List(wk_LOOP), 38, 8)     '2016.10.16
            'Print_tbl(Gyo, Retu).BIKOU = Mid(List1.List(wk_LOOP), 55, 20)      '2016.10.16
    
            'Print_tbl(Gyo, Retu).IRI_QTY = Mid(List1.List(wk_LOOP), 36, 8)      '2016.10.16  2018.09.21
            Print_tbl(Gyo, Retu).IRI_QTY = Mid(List1.List(wk_LOOP), 38, 6)      '             2018.09.21
            Print_tbl(Gyo, Retu).BIKOU = Mid(List1.List(wk_LOOP), 53, 20)       '2016.10.16
    
            Print_tbl(Gyo, Retu).BIKOU2 = Trim(List3.List(wk_LOOP))       '2016.10.16
    
    
            Retu = Retu + 1
            If Retu > 1 Then
                Gyo = Gyo + 1
                If Gyo > Max_Gyo Then
                    
                    
                    If Print_Flg = 1 Then                       '2019.01.21
                        If Max_Gyo = 2 Then                     '2019.01.21
                            Call New_Print_Sub_A5_Proc          '2019.01.21
                        Else                                    '2019.01.21
                           Call New_Print_Sub_Proc              '2019.01.21
                        End If                                  '2019.01.21
                    Else                                        '2019.01.21
                        Call Print_Sub_Proc                     '2019.01.21
                    End If                                      '2019.01.21
                    
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
    
'    Call Print_Sub_Proc
        
    If Print_Flg = 1 Then                       '2019.01.21
        If Option1(0).Value = True Then         '2019.01.21
            Call New_Print_Sub_A5_Proc          '2019.01.21
        Else                                    '2019.01.21
           Call New_Print_Sub_Proc              '2019.01.21
        End If                                  '2019.01.21
    Else                                        '2019.01.21
        Call Print_Sub_Proc                     '2019.01.21
    End If                                      '2019.01.21
        
        
        
        
        
        
    Exit Function           '2018.10.24
    
Error_Proc:
    
    Select Case Err.Number
        Case 380
            MsgBox "指定のプリンターは使用できません。"
        
        Case 482
        
            MsgBox "指定のプリンターは使用できません。"
        
        Case Else
            MsgBox Err.Description & " Error= " & Err.Number
            Unload Me
    End Select
End Function
                                    
                                    '画面初期状態を設定する
Private Sub Clear_Field()
Dim i As Integer
    
    For i = 0 To 4
        text(i).text = ""
    Next i
    text(ptxIriSuu).text = ""
    
    text(ptxSHIIRE_WORK_CENTER).text = ""

    text(ptxBikou2).text = ""

'    text(ptxGoukei).text = "0"
'    text(ptxwkMaiSuu).text = "0"



    lblST_TANABAN(0).Caption = ""

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.31
    Combo(1).Clear
    Combo(1).ListIndex = -1
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.31


    lblKEPPIN_CNT.Caption = ""      '2013.08.23
    lblKEPPIN_QTY.Caption = ""      '2013.08.23


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

Private Sub Combo1_LostFocus()

    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<   2016.09.30
    If WriteIni(App.EXEName, "LAST_PRINTER", App.EXEName, Combo1.text) Then
        Beep
        MsgBox ("INIファイルの書き込みに失敗しました。[" & App.EXEName & "]LAST_PRINTER=")
        Unload Me
    End If
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2016.09.30



End Sub

Private Sub Command_Click(Index As Integer)

Dim yn              As Integer
Dim RetBuf          As String
Dim sts             As Integer
Dim wkList_Box      As String
Dim wk_Kbn          As String * 1
Dim wk_Bikou        As String * 20
Dim wk_Maisuu       As Integer

'Dim wk_IRI_QTY      As String * 8      '2018.09.21
Dim wk_IRI_QTY      As String * 6       '2018.09.21

Dim wk_MAISU        As String * 3

Dim wk_AMARI        As Integer          '2017.11.20
Dim wk_keisan       As Integer          '2017.11.20


Dim wkGENSAN        As String * 22


Dim wkSHIIRE_WORK_CENTER As String * 8

Dim wkHin_Nai       As String * 15      '2018.09.21
'Dim wkHin_Nai       As String * 13     '2018.09.21


Dim wkKEPPIN_QTY    As String * 8       '2013.08.23


Dim Fsw         As Integer      '2016.12.27

Dim Fsw_NAI     As Integer      '2017.01.10

Dim mesg        As String       '2017.11.20

Select Case Index
        Case 0                              '確定
                                            
                                            '外部品番で読み込み
'            If Len(text(ptxHin_Gai).text) <> 0 Then
                
            '2010.11.25
            If Len(text(ptxHin_Gai).text) <> 0 And Len(text(ptxHin_Nai).text) = 0 Then
                
                Fsw = 0         '2017.01.10
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
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.01.10 桁数変更で読み替え
                        If HIN_GAI_LTRIM > 0 Then
                            If HIN_GAI_LTRIM < Len(Trim(text(ptxHin_Gai).text)) Then
                                If Fsw = 0 Then
                                    Fsw = 1
                                    text(ptxHin_Gai).text = Mid(text(ptxHin_Gai).text, HIN_GAI_LTRIM + 1, Len(text(ptxHin_Gai).text) - HIN_GAI_LTRIM)
                                    GoTo Item_Read
                                End If
                            End If
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.01.10 桁数変更で読み替え
                        
                        
                        
                        
                        MsgBox "入力したコードは登録されていません。"
                        Exit Sub
                    Case Else
                        
                        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.19
                        If sts > 3000 Or sts = 3 Then
                        
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                        
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2015.04.20
                            'sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "棚マスタ")
                            '    Beep
                            '    MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
                            'End If
                        
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
        
        
                    Do          '2015.04.20
                        If Not File_Open_Proc Then
                            Exit Do
                        End If
                    Loop
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
                
                Fsw_NAI = 0         '2017.01.10
                
Item_Read2:
                
                '2010.11.25
                If Len(text(ptxHin_Gai).text) = 0 And Len(text(ptxHin_Nai).text) = 0 Then              '2017.01.10
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
                            
                            
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.01.10 桁数変更で読み替え(内部品番)
                            If HIN_NAI_LTRIM > 0 Then
                                If HIN_NAI_LTRIM < Len(Trim(text(ptxHin_Nai).text)) Then
                                    If Fsw_NAI = 0 Then
                                        Fsw_NAI = 1
                                        text(ptxHin_Nai).text = Mid(text(ptxHin_Nai).text, HIN_NAI_LTRIM + 1, Len(text(ptxHin_Nai).text) - HIN_NAI_LTRIM)
                                        GoTo Item_Read2
                                    End If
                                End If
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.01.10 桁数変更で読み替え(内部品番)
                            
                            
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
    
            If RTrim(text(ptxHin_Gai).text) <> RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                Beep
                MsgBox "品番（外部）又は品番（内部）を入力後、Enterｷｰで入力を確定して下さい"
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
                'wk_Kbn = NAIGAI_NAI                                                        '2016.10.13　国内外表示削除
                'wkList_Box = wk_Kbn & " " & StrConv(ITEMREC.HIN_GAI, vbUnicode) + " "      '2016.10.13
                wkList_Box = StrConv(ITEMREC.HIN_GAI, vbUnicode) + " "                      '2016.10.13


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
                
                
                wkKEPPIN_QTY = lblKEPPIN_QTY.Caption            '2013.08.23
                wkList_Box = wkList_Box & wkKEPPIN_QTY          '2013.08.23
                
                
                List1.AddItem wkList_Box
            
                List3.AddItem text(ptxBikou2).text
            
            

            
            End If
                        
            If Item_Update_Proc() Then
                Unload Me
            End If
            
            
'>>>>>>>>>  枚数計算    2017.11.20
            
            wk_Maisuu = CInt(text(ptxGoukei).text) - CInt(text(ptxwkMaiSuu).text) + CInt(text(ptxMaiSuu).text)
            
            
            If Option1(0).Value Then
                lblSIZE.Caption = Option1(0).Caption
            Else
                lblSIZE.Caption = Option1(1).Caption
            End If
            
            If Option1(0).Value Then
                wk_Maisuu = CInt(ToRoundUp(CCur(wk_Maisuu / 6), 0))
                wk_keisan = CInt(text(ptxGoukei).text) - CInt(text(ptxwkMaiSuu).text) + CInt(text(ptxMaiSuu).text)
                
                wk_keisan = CInt(ToRoundDown(CCur(wk_keisan / 6), 0))
                wk_AMARI = (CInt(text(ptxGoukei).text) - CInt(text(ptxwkMaiSuu).text) + CInt(text(ptxMaiSuu).text) - wk_keisan * 6)
                
            Else
                wk_Maisuu = CInt(ToRoundUp(CCur(wk_Maisuu / 12), 0))
                wk_keisan = CInt(text(ptxGoukei).text) - CInt(text(ptxwkMaiSuu).text) + CInt(text(ptxMaiSuu).text)
                
                wk_keisan = CInt(ToRoundDown(CCur(wk_keisan / 12), 0))
                wk_AMARI = (CInt(text(ptxGoukei).text) - CInt(text(ptxwkMaiSuu).text) + CInt(text(ptxMaiSuu).text) - wk_keisan * 12)
            
            End If

            mesg = wk_Maisuu & "枚(端数" & wk_AMARI & ")"
            lblMAISUU.Caption = mesg
'>>>>>>>>>  枚数計算    2017.11.20
            
            wk_Maisuu = CInt(text(ptxGoukei).text) - CInt(text(ptxwkMaiSuu).text) + CInt(text(ptxMaiSuu).text)
            Call Clear_Field
            text(ptxGoukei).text = Format(wk_Maisuu, "#0")
            text(ptxHin_Nai).SetFocus
        
        
        
        
        Case 4                              'クリア 2018.12.40
            
                    List1.Clear
                    List3.Clear
                    text(ptxGoukei).text = "0"
                    text(ptxwkMaiSuu).text = "0"
        
        
        Case 8                              '印刷
            Beep
            yn = MsgBox("印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                
                sts = Print_Proc()
                Printer.EndDoc
                Call Clear_Field
                
                If CLEAR_BUTTON = 0 Then        '2018.12.04
                
                    List1.Clear
                    List3.Clear
            
                    text(ptxGoukei).text = "0"
                    text(ptxwkMaiSuu).text = "0"
            
            
            
                End If                          '2018.12.04
            End If
            
            text(ptxHin_Nai).SetFocus
            
        Case 11                             '終了
            If List1.ListCount = 0 Then
                Unload Me
            End If
            Beep
            yn = MsgBox("終了しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                Unload Me
            End If
            text(ptxHin_Nai).SetFocus
            
        Case Else
            Beep
    End Select
    
End Sub


Private Sub Form_DblClick()
'''    PrintForm            '2017.03.08
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
    
Dim sBuffer     As String * 255     '2016.12.28
Dim com         As String           '2016.12.28
    
    
    
  Dim LeftMargin As Long, TopMargin     As Long
    Dim RightMargin As Long, BottomMargin As Long
    Dim PhysHeight As Long, PhysWidth     As Long
    
    
    
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
            F1020501.Caption = "入庫現品票印刷（" + RTrim(JGYOBU_T(i).NAME) + ") " & Last_Update_Day
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)


                                'ﾜｰｸｽﾃｰｼｮﾝ番号取り込み  2016.12.28
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)


                                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.31
'                                '倉庫マスタＯＰＥＮ
'    If SOKO_Open(0) Then
'        Beep
'        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'        Unload Me
'    End If
'                                '品目マスタＯＰＥＮ
'    If ITEM_Open(0) Then
'        Beep
'        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'        Unload Me
'    End If
'
'    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.05.30
'                                'PNマスタＯＰＥＮ
'    If PN_M_Open(0) Then
'        Beep
'        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'        Unload Me
'    End If
'    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'                                '原産国マスタＯＰＥＮ
'    If GENSAN_Open(0) Then
'        Beep
'        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'        Unload Me
'    End If
    
    
    '2015.04.20
    Do
        If Not File_Open_Proc Then
            Exit Do
        End If
    Loop
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.31
    
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2016.09.30 INIﾌｧｲﾙを独立
                                'デフォルト用紙サイズ取り込み
'    If GetIni(App.EXEName, "DEF", "SYS", c) Then                   '2016.09.30
    If GetIni(App.EXEName, "DEF", App.EXEName, c) Then              '2016.09.30
        c = ""
    End If
    DEF = RTrim(c)
                                
                                
                                '仕入れ先更新可否   2010.10.07
'    If GetIni(App.EXEName, "SHIIRE_WORK_CENTER_F", "SYS", c) Then          '2016.09.30
    If GetIni(App.EXEName, "SHIIRE_WORK_CENTER_F", App.EXEName, c) Then     '2016.09.30
        SHIIRE_WORK_CENTER_F = True
    Else
    
        If Trim(c) = "1" Then
            SHIIRE_WORK_CENTER_F = False
        Else
            SHIIRE_WORK_CENTER_F = True
        End If
    End If
    text(ptxSHIIRE_WORK_CENTER).Locked = SHIIRE_WORK_CENTER_F
                                
                                
                                '前回使用　Printer  2016.09.30
    If GetIni(App.EXEName, "LAST_PRINTER", App.EXEName, c) Then
        c = ""
    End If
    Last_Printer = RTrim(c)



                                '除外桁数   2016.12.27
    If GetIni(App.EXEName, "HIN_GAI_LTRIM", App.EXEName, c) Then
        HIN_GAI_LTRIM = 0
    End If
    HIN_GAI_LTRIM = Val(Trim(c))

                                '除外桁数(内部)   2017.01.10
    If GetIni(App.EXEName, "HIN_NAI_LTRIM", App.EXEName, c) Then
        HIN_NAI_LTRIM = 0
    End If
    HIN_NAI_LTRIM = Val(Trim(c))



'>>>>>>>>>>>>>>>>>>>>   作業ログ情報    2016.12.27
    If GetIni(App.EXEName, "MENU_NO", App.EXEName, c) Then      '2016.01.15
        MENU_NO = ""
    Else
        MENU_NO = Trim(c)
    End If
                                
    If GetIni(App.EXEName, "RIRK_ID", App.EXEName, c) Then      '2016.01.15
        RIRK_ID = ""
    Else
        RIRK_ID = Trim(c)
    End If
                                
    If GetIni(App.EXEName, "MEMO", App.EXEName, c) Then         '2016.01.15
        MEMO = ""
    Else
        MEMO = Trim(c)
    End If
'>>>>>>>>>>>>>>>>>>>>   作業ログ情報    2016.12.27



'>>>>>>>>>>>>>>>>>>>>   クリアーボタン    2018.12.04
    If GetIni(App.EXEName, "CLEAR_BUTTON", App.EXEName, c) Then
        CLEAR_BUTTON = 0
    Else
        If Trim(c) = "1" Then
            CLEAR_BUTTON = 1
        Else
            CLEAR_BUTTON = 0
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>   クリアーボタン    2018.12.04





'>>>>>>>>>>>>>>>>>>>>   印刷制御    2019.01.21
    If GetIni(App.EXEName, "PRINT_FLG", App.EXEName, c) Then
        Print_Flg = 0
    Else
        If Trim(c) = "1" Then
            Print_Flg = 1
        Else
            Print_Flg = 0
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>   印刷制御    2019.01.21


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2016.09.30 INIﾌｧｲﾙを独立
                                
                                
                                
                                '印刷フォント設定
    With Code39Font
        .NAME = Label1.FontName
        .Size = Label1.FontSize
    End With
    Set Printer.Font = Code39Font
                                '印刷フォント設定
    With NormalFont
        .NAME = F1020501.FontName
        .Size = F1020501.FontSize
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
    List3.Clear
    
    text(ptxGoukei).text = "0"
    text(ptxwkMaiSuu).text = "0"

    
    Combo1.Clear
    
    
    If Trim(Last_Printer) = "" Then                                 '2016.09.30
        For Each Pri_Name In Printers
            If Pri_Name.DeviceName = Printer.DeviceName Then
                Combo1.AddItem Pri_Name.DeviceName
            End If
        Next
        For Each Pri_Name In Printers
'            If Pri_Name.DeviceName <> Printer.DriverName Then      '2016.10.07
            If Pri_Name.DeviceName <> Printer.DeviceName Then       '2016.10.07
                Combo1.AddItem Pri_Name.DeviceName
            End If
        Next
    Else                                                            '2016.09.30
        For Each Pri_Name In Printers                               '2016.09.30
            If Pri_Name.DeviceName = Last_Printer Then              '2016.09.30
                Combo1.AddItem Pri_Name.DeviceName                  '2016.09.30
            End If                                                  '2016.09.30
        Next                                                        '2016.09.30
        For Each Pri_Name In Printers                               '2016.09.30
            If Pri_Name.DeviceName <> Last_Printer Then             '2016.09.30
                Combo1.AddItem Pri_Name.DeviceName                  '2016.09.30
            End If                                                  '2016.09.30
        Next                                                        '2016.09.30
    End If                                                          '2016.09.30
    
    
    
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2017.03.03
    If GetIni(App.EXEName, "GENSAN", App.EXEName, c) Then         '2016.01.15
        GENSAN_KOKU_F = 0
    Else
        GENSAN_KOKU_F = Val(Trim(c))
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2017.03.03
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2018.12.04
    If CLEAR_BUTTON = 1 Then
        Command(4).Caption = "クリア"
        Command(4).Enabled = True
    Else
        Command(4).Caption = ""
        Command(4).Enabled = False
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2018.12.04
    
    
    
      With Printer
        'マージンをピクセル単位で取得しそれをmmに変換
        LeftMargin = .ScaleX(GetDeviceCaps(.hDC, PHYSICALOFFSETX), _
                                            vbPixels, vbMillimeters)
        TopMargin = .ScaleY(GetDeviceCaps(.hDC, PHYSICALOFFSETY), _
                                            vbPixels, vbMillimeters)
        PhysWidth = .ScaleX(GetDeviceCaps(.hDC, PHYSICALWIDTH), _
                                            vbPixels, vbMillimeters)
        PhysHeight = .ScaleY(GetDeviceCaps(.hDC, PHYSICALHEIGHT), _
                                            vbPixels, vbMillimeters)
        '用紙サイズから印刷可能領域を差引きマージンを求める
        RightMargin = PhysWidth - (.ScaleX(GetDeviceCaps(.hDC, HORZRES), _
                                    vbPixels, vbMillimeters) + LeftMargin)
        BottomMargin = PhysHeight - (.ScaleY(GetDeviceCaps(.hDC, VERTRES), _
                                    vbPixels, vbMillimeters) + TopMargin)
    End With
    
'    MsgBox "上　" & TopMargin
'    MsgBox "下　" & BottomMargin
    
    
'>>>    2019.02.27
    If Print_Flg = 0 Then
        Label(29).Visible = False
        text(ptxBikou2).Visible = False
    Else
        Frame1.Enabled = False
        Option1(0).Enabled = False
    End If
'>>>    2019.02.27
    
    Combo1.ListIndex = 0
    
    text(ptxHin_Nai).SetFocus
    
    End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer


    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.05.30
                                            'PNマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, PN_M_POS, PN_MREC, Len(PN_MREC), K0_PN_M, Len(K0_PN_M), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "PNマスタ", 0)
'2015.06.19            Beep
'2015.06.19            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        End If
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    
    
                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "棚マスタ", 0)
'2015.06.19            Beep
'2015.06.19            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ", 0)
'2015.06.19            Beep
'2015.06.19            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        End If
    End If
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "棚マスタ", 0)
'2015.06.19        Beep
'2015.06.19        MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
    End If

    End
End Sub

Private Sub List1_DblClick()

Dim ans     As Integer

    
Dim Num     As Integer                                                          '2016.10.13
Dim wk_Maisuu   As Integer  '2017.11.21
Dim wk_AMARI    As Integer  '2017.11.21
Dim wk_keisan   As Integer  '2017.11.21
Dim mesg        As String  '2017.11.21
    
Dim Lindex      As Integer
    
    
    ans = MsgBox("指定行を削除しますか？", vbYesNo + vbDefaultButton1, "確認入力")
    
    If ans = vbYes Then
        
        
        Num = Val(Mid(List1.List(List1.ListIndex), 47, 3))                      '2016.10.13
        text(ptxGoukei).text = Format(Val(text(ptxGoukei).text) - Num, "#0")    '2016.10.13
        
        
'>>>>>>>>>  枚数計算    2017.11.20
            wk_Maisuu = CInt(text(ptxGoukei).text)
            
            
            If Option1(0).Value Then
                lblSIZE.Caption = Option1(0).Caption
            Else
                lblSIZE.Caption = Option1(1).Caption
            End If
            
            If Option1(0).Value Then
                wk_Maisuu = CInt(ToRoundUp(CCur(wk_Maisuu / 6), 0))
                wk_keisan = CInt(text(ptxGoukei).text)
                
                wk_keisan = CInt(ToRoundDown(CCur(wk_keisan / 6), 0))
                wk_AMARI = (CInt(text(ptxGoukei).text) - wk_keisan * 6)
                
            Else
                wk_Maisuu = CInt(ToRoundUp(CCur(wk_Maisuu / 12), 0))
                wk_keisan = CInt(text(ptxGoukei).text)
                
                wk_keisan = CInt(ToRoundDown(CCur(wk_keisan / 12), 0))
                wk_AMARI = (CInt(text(ptxGoukei).text) - wk_keisan * 12)
            
            End If

            mesg = wk_Maisuu & "枚(端数" & wk_AMARI & ")"
            lblMAISUU.Caption = mesg
    
'>>>>>>>>>  枚数計算    2017.11.20
        
        
        
        
        
        
       Lindex = List1.ListIndex
        
        List1.RemoveItem List1.ListIndex
    
    
        List3.RemoveItem Lindex        '2019.02.26
    
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
    F1020501.Caption = "入庫現品票印刷（" + RTrim(JGYOBU_T(Index).NAME) + ") " & Last_Update_Day
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


Dim Fsw         As Integer      '2016.12.27

Dim Fsw_NAI     As Integer      '2017.01.10

Dim wkHIN_Code  As String       '2017.01.12



    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown
            Select Case Index
                Case 0
                        
                    Fsw = 0         '2016.12.27
Item_RE_READ:                       '2016.12.27
                    If Len(text(ptxHin_Gai).text) <> 0 Then
                                                
    
    
                        text(Index).text = RTrim(StrConv(text(Index).text, vbUpperCase))
                                                
                        If Fsw = 0 Then                         '2017.01.12
                            wkHIN_Code = text(Index).text       '2017.01.12
                        End If                                  '2017.01.12
                                                '外部品番で読み込み
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                        If Combo(0).text = NAIGAI1$ Then
                            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI$)
                        Else
                            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI$)
                        End If
                        'Call UniCode_Conv(K0_ITEM.HIN_GAI, RTrim(text(ptxHin_Gai).text))       '2017.01.12
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, wkHIN_Code)                          '2017.01.12
                        
Item_Read:
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                                
                                text(Index).text = wkHIN_Code                       '2017.01.12
                                
                                
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
                                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.08.23
                                If Trim(text(ptxHin_Gai).text) <> "" Then
KEPPIN_Read:
''2015.05.27 DEL                      Call UniCode_Conv(K0_KEPPIN.HIN_GAI, text(ptxHin_Gai).text)
''                                    sts = BTRV(BtOpGetEqual, KEPPIN_POS, KEPPINREC, Len(KEPPINREC), K0_KEPPIN, Len(K0_KEPPIN), 0)
''                                    Select Case sts
''                                        Case BtNoErr
''
''                                            lblKEPPIN_CNT = Val(StrConv(KEPPINREC.KEPPIN_CNT, vbUnicode))
''                                            lblKEPPIN_QTY = Val(StrConv(KEPPINREC.KEPPIN_QTY, vbUnicode))
''
''
''
''                                        Case BtErrKeyNotFound
''
''                                            lblKEPPIN_CNT = ""
''                                            lblKEPPIN_QTY = ""
''
''                                        Case Else
''                                            If sts > 3000 Or sts = 3 Then
''
''                                                Call File_Error(sts, BtOpGetEqual, "欠品データ", 0)
''
''                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2015.04.20
''                                                'sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
''                                                'If sts Then
''                                                '    Call File_Error(sts, BtOpReset, "")
''                                                '    Beep
''                                                '    MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
''                                                'End If
''
''                                                Do          '2015.04.20
''                                                    If Not File_Open_Proc Then
''                                                        Exit Do
''                                                    End If
''                                                Loop
''
''
''                                                GoTo KEPPIN_Read
''
''
''                                            End If
''
''
''
''                                            Call File_Error(sts, BtOpGetEqual, "欠品データ")
''                                            Beep
''                                            MsgBox "システム異常が発生しました。処理を中止して下さい。"
''                                            Unload Me
''
''                                    End Select
                                End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.08.23
                                
                                
                                
                                
'                                text(ptxMaiSuu).SetFocus
                                text(ptxIriSuu).SetFocus
                                Call Text_GotFocus(ptxIriSuu)
                  
'                                text(ptxHin_Nai).SetFocus   '2010.10.18
                  
                  
                            Case BtErrKeyNotFound
                                
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2016.12.27 桁数変更で読み替え
                                If HIN_GAI_LTRIM > 0 Then
                                    
                                    If HIN_GAI_LTRIM < Len(Trim(text(Index).text)) Then
                                    
                                        If Fsw = 0 Then
                                            Fsw = 1
                                            'text(Index).text = Mid(text(Index).text, HIN_GAI_LTRIM + 1, Len(text(Index).text) - HIN_GAI_LTRIM) '2017.01.12
                                            wkHIN_Code = Mid(text(Index).text, HIN_GAI_LTRIM + 1, Len(text(Index).text) - HIN_GAI_LTRIM)        '2017.01.12
                                            GoTo Item_RE_READ
                                        End If
                                    End If
                                End If
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2016.12.27 桁数変更で読み替え
                                
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> MT_2009.05.30
                                'MsgBox "入力したコードは登録されていません。"
                                'Text(0).SetFocus
                                'If PN_CHK(text(Index), "G", "FLABEL", 1) Then      '2017.01.12
                                If PN_CHK(wkHIN_Code, "G", "FLABEL", 1) Then        '2017.01.12
                                    text(Index).SetFocus
                                    Call Text_GotFocus(Index)
                                    Exit Sub
                                End If
                                
                                
                                text(Index).text = wkHIN_Code                       '2017.01.12
                                
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
                                
                                
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.19
                                If sts > 3000 Or sts = 3 Then
                                
                                    Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                
                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2015.04.20
                                    'sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                                    'If sts Then
                                    '    Call File_Error(sts, BtOpReset, "棚マスタ")
                                    '    Beep
                                    '    MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
                                    'End If
                                
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
                
                                    Do          '2015.04.20
                                        If Not File_Open_Proc() Then
                                            Exit Do
                                        End If
                                    Loop
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.31
                                
                                    GoTo Item_Read
                
                                
                                End If
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.19
                                
                                
                                
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                Beep
                                MsgBox "システム異常が発生しました。処理を中止して下さい。"
                                Unload Me
                        End Select
                        
                        '>>>>>>>>>> 2018.02.03
                        Call UniCode_Conv(K0_ITEM_CHG.N_JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                        Call UniCode_Conv(K0_ITEM_CHG.N_NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                        Call UniCode_Conv(K0_ITEM_CHG.N_HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
        
                        
                        Call UniCode_Conv(K0_ITEM_CHG.N_HIN_GAI, RTrim(text(ptxHin_Gai).text))
                        sts = BTRV(BtOpGetEqual, ITEM_CHG_POS, ITEM_CHG_REC, Len(ITEM_CHG_REC), K0_ITEM_CHG, Len(K0_ITEM_CHG), 0)
                        Select Case sts
                            Case BtNoErr
                                text(ptxBikou2).text = StrConv(ITEM_CHG_REC.O_HIN_GAI, vbUnicode)
                            
                            Case BtErrKeyNotFound
                                text(ptxBikou2).text = ""
                            Case Else
                                If sts > 3000 Or sts = 3 Then
                                    Call File_Error(sts, BtOpGetEqual, "品目読み替え", 0)
                                    Do
                                        If Not File_Open_Proc Then
                                            Exit Do
                                        End If
                                    Loop
                                    
                                    GoTo Item_Read2
                    
                                    
                                End If
                                            
                                Call File_Error(sts, BtOpGetEqual, "品目読み替え")
                                Beep
                                MsgBox "システム異常が発生しました。処理を中止して下さい。"
                                Unload Me
                        End Select
                    
                    
                    
                    
                    
                    
                    
                    
                    End If
                Case 2
'2017.01.10                    If Len(text(ptxHin_Gai).text) = 0 Then
                        If Len(text(ptxHin_Nai).text) <> 0 Then
                                                
                            text(Index).text = RTrim(StrConv(text(Index).text, vbUpperCase))
                            Fsw_NAI = 0                               '2017.01.10
                                                
                                                                            
                                                
                                                
                                                
Item_Read2:
                                                
                            If Fsw_NAI = 0 Then                         '2017.01.12
                                wkHIN_Code = text(Index).text           '2017.01.12
                            End If                                      '2017.01.12
                            
                                                
                                                
                                                
                                                '内部品番で読み込み
                            #If Center_chk Then
                                Call UniCode_Conv(K2_ITEM.JGYOBU, Last_JGYOBU)
                                If Combo(0).text = NAIGAI1$ Then
                                    Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI_NAI$)
                                Else
                                    Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI_GAI$)
                                End If
                                'Call UniCode_Conv(K2_ITEM.HIN_NAI, RTrim(text(ptxHin_Nai).text))   '2017.01.12
                                Call UniCode_Conv(K2_ITEM.HIN_NAI, wkHIN_Code)                      '2017.01.12
                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                            #Else
                                Call UniCode_Conv(K1_ITEM.JGYOBU, Last_JGYOBU)
                                If Combo(0).text = NAIGAI1$ Then
                                    Call UniCode_Conv(K1_ITEM.NAIGAI, NAIGAI_NAI$)
                                Else
                                    Call UniCode_Conv(K1_ITEM.NAIGAI, NAIGAI_GAI$)
                                End If
                                'Call UniCode_Conv(K1_ITEM.HIN_NAI, RTrim(text(ptxHin_Nai).text))       '2017.01.12
                                Call UniCode_Conv(K1_ITEM.HIN_NAI, wkHIN_Code)                          '2017.01.12
                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K1_ITEM, Len(K1_ITEM), 1)
                            #End If
                            Select Case sts
                                Case BtNoErr
                                    
                                    text(Index).text = wkHIN_Code                                       '2017.01.12
                                    
                                    
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
                                    
                                    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.08.23
                                    If Trim(text(ptxHin_Gai).text) <> "" Then
KEPPIN_Read2:
''2015.05.27 DEL                          Call UniCode_Conv(K0_KEPPIN.HIN_GAI, text(ptxHin_Gai).text)
''                                        sts = BTRV(BtOpGetEqual, KEPPIN_POS, KEPPINREC, Len(KEPPINREC), K0_KEPPIN, Len(K0_KEPPIN), 0)
''                                        Select Case sts
''                                            Case BtNoErr
''
''                                                lblKEPPIN_CNT = Val(StrConv(KEPPINREC.KEPPIN_CNT, vbUnicode))
''                                                lblKEPPIN_QTY = Val(StrConv(KEPPINREC.KEPPIN_QTY, vbUnicode))
''
''
''
''                                            Case BtErrKeyNotFound
''
''                                                lblKEPPIN_CNT = ""
''                                                lblKEPPIN_QTY = ""
''
''                                            Case Else
''                                                If sts > 3000 Or sts = 3 Then
''
''                                                    Call File_Error(sts, BtOpGetEqual, "欠品データ", 0)
''
''                                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2015.04.20
''                                                    'sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
''                                                    'If sts Then
''                                                    '    Call File_Error(sts, BtOpReset, "")
''                                                    '    Beep
''                                                    '    MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
''                                                    'End If
''
''                                                    Do      '2015.04.20
''                                                        If File_Open_Proc() Then
''                                                            Exit Do
''                                                        End If
''                                                    Loop
''
''                                                    GoTo KEPPIN_Read2
''
''
''                                                End If
''
''
''
''                                                Call File_Error(sts, BtOpGetEqual, "欠品データ")
''                                                Beep
''                                                MsgBox "システム異常が発生しました。処理を中止して下さい。"
''                                                Unload Me
''
''                                        End Select
                                    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.08.23
                                    
                                    
                                    
                                    text(ptxMaiSuu).SetFocus

'''                                    text(ptxHin_Nai).SetFocus   '2010.10.18

                                Case BtErrKeyNotFound
                                    
                                    
                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.01.10 桁数変更で読み替え(内部品番)
                                    If HIN_NAI_LTRIM > 0 Then
                                                                                
                                        If HIN_NAI_LTRIM < Len(Trim(text(Index).text)) Then
                                            If Fsw_NAI = 0 Then
                                                Fsw_NAI = 1
                                                'text(Index).text = Mid(text(Index).text, HIN_NAI_LTRIM + 1, Len(text(Index).text) - HIN_NAI_LTRIM)     '2017.01.12
                                                wkHIN_Code = Mid(text(Index).text, HIN_NAI_LTRIM + 1, Len(text(Index).text) - HIN_NAI_LTRIM)            '2017.01.12
                                                GoTo Item_Read2
                                            End If
                                        End If
                                    End If
                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.01.10 桁数変更で読み替え(内部品番)
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> MT_2009.05.30
                                    'MsgBox "入力したコードは登録されていません。"
                                    'Text(0).SetFocus
                                    
                                    If PN_CHK(text(Index), "N", "FLABEL", 1, 1) Then
                                        text(Index).SetFocus
                                        Call Text_GotFocus(Index)
                                        Exit Sub
                                    End If
                                    
                                    
                                    text(Index).text = wkHIN_Code                   '2017.01.12
                                    
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
                                    
                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2015.04.20
                                        'sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                                        'If sts Then
                                        '    Call File_Error(sts, BtOpReset, "棚マスタ")
                                        '    Beep
                                        '    MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
                                        'End If
                                    
                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.31
                                    
                                        '                            '倉庫マスタＯＰＥＮ
                                        'If SOKO_Open(0) Then
                                        '    Beep
                                        '    MsgBox "システム異常が発生しました。処理を中止して下さい。"
                                        '    Unload Me
                                        'End If
                                        '                            '品目マスタＯＰＥＮ
                                        'If ITEM_Open(0) Then
                                        '    Beep
                                        '    MsgBox "システム異常が発生しました。処理を中止して下さい。"
                                        '    Unload Me
                                        'End If
                                        '
                                        ''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.05.30
                                        '                            'PNマスタＯＰＥＮ
                                        'If PN_M_Open(0) Then
                                        '    Beep
                                        '    MsgBox "システム異常が発生しました。処理を中止して下さい。"
                                        '    Unload Me
                                        'End If
                                        ''<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                        '                            '原産国マスタＯＰＥＮ
                                        'If GENSAN_Open(0) Then
                                        '    Beep
                                        '    MsgBox "システム異常が発生しました。処理を中止して下さい。"
                                        '    Unload Me
                                        'End If
                                        '
                                        ''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.31
                                    
                                        
                                        Do      '2015.04.20
                                            If Not File_Open_Proc() Then
                                                Exit Do
                                            End If
                                        Loop
                                    
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
                        'End If                 2017.01.10
                    
                    
                    
                    
                
                        '>>>>>>>>>> 2018.02.03
                        Call UniCode_Conv(K0_ITEM_CHG.N_JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                        Call UniCode_Conv(K0_ITEM_CHG.N_NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                        Call UniCode_Conv(K0_ITEM_CHG.N_HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
        
                        
                        Call UniCode_Conv(K0_ITEM_CHG.N_HIN_GAI, RTrim(text(ptxHin_Gai).text))
                        sts = BTRV(BtOpGetEqual, ITEM_CHG_POS, ITEM_CHG_REC, Len(ITEM_CHG_REC), K0_ITEM_CHG, Len(K0_ITEM_CHG), 0)
                        Select Case sts
                            Case BtNoErr
                                text(ptxBikou2).text = StrConv(ITEM_CHG_REC.O_HIN_GAI, vbUnicode)
                            
                            Case BtErrKeyNotFound
                                text(ptxBikou2).text = ""
                            Case Else
                                If sts > 3000 Or sts = 3 Then
                                    Call File_Error(sts, BtOpGetEqual, "品目読み替え", 0)
                                    Do
                                        If Not File_Open_Proc Then
                                            Exit Do
                                        End If
                                    Loop
                                    
                                    GoTo Item_Read2
                    
                                    
                                End If
                                            
                                Call File_Error(sts, BtOpGetEqual, "品目読み替え")
                                Beep
                                MsgBox "システム異常が発生しました。処理を中止して下さい。"
                                Unload Me
                        End Select
                    
                    
                    
                    End If
            
            
            
            
            
            
            End Select
            
            
            
            
            If Index < 3 Then
                text(ptxIriSuu).SetFocus
            End If
            If Index = ptxIriSuu Then
                text(ptxMaiSuu).SetFocus
            End If
            If Index > 2 Then
                If Index < 7 Then
                    text(Index + 1).SetFocus
                End If
            
                If Index = 7 Then                   '2018.02.06
                    
                    If text(12).Visible Then        '2019.04.01
                    
                        text(12).SetFocus               '2018.02.06
                    
                    
                    Else
                        If text(11).Locked = False Then      '2019.04.01
                            text(11).SetFocus           '2019.04.01
                        
                        End If                          '2019.04.01
                    End If
                End If                              '2018.02.06
            
            
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
            .NAME = F1020501.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(20);                                 '2013.08.23      2014.02.18
                                                                
'        If Trim(Print_tbl(Gyo, 0).KEPPIN_QTY) = "" Then         '2013.08.23    2014.02.18
'            Printer.Print Tab(20);                              '2013.08.23    2014.02.18
'        Else                                                    '2013.08.23    2014.02.18
'            Printer.Print Tab(2);                               '2013.08.23    2014.02.18
'            Printer.Print "欠品";                               '2013.08.23    2014.02.18
'            Printer.Print "(";                                  '2013.08.23    2014.02.18
'            Printer.Print Trim(Print_tbl(Gyo, 0).KEPPIN_QTY);   '2013.08.23    2014.02.18
'            Printer.Print ")";                                  '2013.08.23    2014.02.18
'            Printer.Print Tab(20);                              '2013.08.23    2014.02.18
'        End If                                                  '2013.08.23    2014.02.18
        
        Printer.Print "入庫現品票";
        Printer.Print Tab(47);
        Printer.Print Trim(JGYOBU_NAME);

        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            
            Printer.Print Tab(80);                                 '2013.08.23      2014.02.18
'            If Trim(Print_tbl(Gyo, 1).KEPPIN_QTY) = "" Then         '2013.08.23    2014.02.18
'                Printer.Print Tab(80);                              '2013.08.23    2014.02.18
'            Else                                                    '2013.08.23    2014.02.18
'                Printer.Print Tab(62);                              '2013.08.23    2014.02.18
'                Printer.Print "欠品";                               '2013.08.23    2014.02.18
'                Printer.Print "(";                                  '2013.08.23    2014.02.18
'                Printer.Print Trim(Print_tbl(Gyo, 1).KEPPIN_QTY);   '2013.08.23    2014.02.18
'                Printer.Print ")";                                  '2013.08.23    2014.02.18
'                Printer.Print Tab(80);                              '2013.08.23    2014.02.18
'            End If                                                  '2013.08.23    2014.02.18
            Printer.Print "入庫現品票";
            Printer.Print Tab(104);
            Printer.Print Trim(JGYOBU_NAME)
        End If
'------------------------------------------------   2行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 6
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   3行目   ------------------
        Set Printer.Font = Code39Font
        Printer.Print Tab(2);
        Printer.Print "*" + Trim(Left(Print_tbl(Gyo, 0).HIN_GAI, 16)) + "*";    '2019/11/08 外部品番16桁対応
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(23);
            Printer.Print "*" + Trim(Left(Print_tbl(Gyo, 1).HIN_GAI, 16)) + "*"    '2019/11/08 外部品番16桁対応
        End If
'------------------------------------------------   4行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   5行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
'        Printer.Print "品番";         2019/11/08 外部品番16桁対応
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 18
        End With
        Set Printer.Font = NormalFont
        Printer.Print Left(Print_tbl(Gyo, 0).HIN_GAI, 16); '2019/11/08 外部品番16桁対応
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
'       Printer.Print "(" & Left(Print_tbl(Gyo, 0).HIN_NAI, 14) & ")";     '2018.09.21
'       Printer.Print "(" & Left(Print_tbl(Gyo, 0).HIN_NAI, 15) & ")";     '2018.09.21
        Printer.Print "  " & Left(Print_tbl(Gyo, 0).HIN_NAI, 16);          '2019/08/26 対内品番の()削除
        
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
            
            
'            Printer.Print "品番";     2019/11/08 外部品番16桁対応
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 18
            End With
            Set Printer.Font = NormalFont
            Printer.Print Left(Print_tbl(Gyo, 1).HIN_GAI, 16); '2019/11/08 外部品番16桁対応
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
'           Printer.Print "(" & Left(Print_tbl(Gyo, 1).HIN_NAI, 14) & ")"          '2018.09.21
'           Printer.Print "(" & Left(Print_tbl(Gyo, 1).HIN_NAI, 15) & ")"          '2018.09.21
            Printer.Print "  " & Left(Print_tbl(Gyo, 1).HIN_NAI, 16)               '2019/08/26 対内品番の()削除
            
        End If
'------------------------------------------------   6行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 4
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   7行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
        Printer.Print LeftB(Print_tbl(Gyo, 0).HIN_NAME, 80);
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
            Printer.Print LeftB(Print_tbl(Gyo, 1).HIN_NAME, 80) '2019/11/08 外部品番16桁対応
        End If
'------------------------------------------------   8行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 4
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   9行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print "　　入数" & ":";
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 18
        End With
        Set Printer.Font = NormalFont
        Printer.Print Format(Print_tbl(Gyo, 0).IRI_QTY, "#0");
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(30);
        Printer.Print "入荷日" & ":";
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 18
        End With
        Set Printer.Font = NormalFont
        Printer.Print text(ptxNyuka_YY).text & "/" & text(ptxNyuka_MM).text & "/" & text(ptxNyuka_DD).text;
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            Set Printer.Font = NormalFont
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 10
            End With
            Set Printer.Font = NormalFont
            
            Printer.Print "　　入数" & ":";
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 18
            End With
            Set Printer.Font = NormalFont
            Printer.Print Format(Print_tbl(Gyo, 1).IRI_QTY, "#0");
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 10
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(88);
            Printer.Print "入荷日" & ":";
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 18
            End With
            Set Printer.Font = NormalFont
            Printer.Print text(ptxNyuka_YY).text & "/" & text(ptxNyuka_MM).text & "/" & text(ptxNyuka_DD).text
        End If
'------------------------------------------------   10行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 4
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   11行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F1020501.FontName
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
                .NAME = F1020501.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 10
            End With
            Printer.Print "標準棚番" & ":" & Print_tbl(Gyo, 1).ST_SOKO & "-" & Print_tbl(Gyo, 1).ST_RETU & "-" & Print_tbl(Gyo, 1).ST_REN & "-" & Print_tbl(Gyo, 1).ST_DAN;
            Printer.Print Tab(88);
            Printer.Print "　備考" & ":" & RTrim(LeftB(Print_tbl(Gyo, 1).BIKOU, 40))
        End If
'------------------------------------------------   12行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        
        
        
        wkGENSAN = Left(Print_tbl(Gyo, 0).GENSAN, 13) & Right(Print_tbl(Gyo, 0).GENSAN, 2)
        
        
        
                
'>>>>>>>>>>>>>>>    2017.03.03
'        Printer.Print "　原産国" & ":" & Left(Print_tbl(Gyo, 0).GENSAN, 15);
        
        
        '2018.10.16
        If GENSAN_KOKU_F = 2 Then
        Else
        '2018.10.16
        
            If GENSAN_KOKU_F = 0 Then
                Printer.Print "　原産国" & ":" & wkGENSAN;
            Else
                If Print_tbl(Gyo, 0).GAI_BUHIN = "1" Or Print_tbl(Gyo, 0).GAI_BUHIN = "2" Or Print_tbl(Gyo, 0).GAI_BUHIN = "3" Then
                    Printer.Print "　原産国注意";
                Else
                    Printer.Print "　原産国" & ":" & wkGENSAN;
                End If
            End If
        End If
'>>>>>>>>>>>>>>>    2017.03.03
        
        
        Printer.Print Tab(30);
        Printer.Print "仕入先" & ":" & Print_tbl(Gyo, 0).SHIIRE_WORK_CENTER;
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 10
            End With
'            Printer.Print "　原産国" & ":" & Left(Print_tbl(Gyo, 1).GENSAN, 15);
            
        wkGENSAN = Left(Print_tbl(Gyo, 1).GENSAN, 13) & Right(Print_tbl(Gyo, 1).GENSAN, 2)
'        Printer.Print "　原産国" & ":" & Left(Print_tbl(Gyo, 0).GENSAN, 15);
        
        
'>>>>>>>>>>>>>>>    2017.03.03
'        Printer.Print "　原産国" & ":" & Left(Print_tbl(Gyo, 0).GENSAN, 15);
        '2018.10.16
        If GENSAN_KOKU_F = 2 Then
        Else
        '2018.10.16
            If GENSAN_KOKU_F = 0 Then
                Printer.Print "　原産国" & ":" & wkGENSAN;
            Else
                If Print_tbl(Gyo, 1).GAI_BUHIN = "1" Or Print_tbl(Gyo, 1).GAI_BUHIN = "2" Or Print_tbl(Gyo, 1).GAI_BUHIN = "3" Then
                    Printer.Print "　原産国注意";
                Else
                    Printer.Print "　原産国" & ":" & wkGENSAN;
                End If
            End If
        End If
'>>>>>>>>>>>>>>>    2017.03.03
            
            
            Printer.Print Tab(88);
            Printer.Print "仕入先" & ":" & Print_tbl(Gyo, 1).SHIIRE_WORK_CENTER;
        End If




'------------------------------------------------   13行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 8
        End With
        Set Printer.Font = NormalFont
        Printer.Print






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
                .NAME = F1020501.FontName
                .Size = 10
            End With
            Set Printer.Font = NormalFont
            Printer.Print
            
            



            If Max_Gyo <> 2 Then
            
                With NormalFont
                    .NAME = F1020501.FontName
                    .Size = 6
                End With
                Set Printer.Font = NormalFont
                Printer.Print
                Printer.Print
            Else
                With NormalFont
                    .NAME = F1020501.FontName
                    .Size = 4
                End With
                Set Printer.Font = NormalFont
                Printer.Print
                With NormalFont
                    .NAME = F1020501.FontName
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
                
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2015.04.20
                    'sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    'If sts Then
                    '    Call File_Error(sts, BtOpReset, "棚マスタ")
                    '    Beep
                    '    MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
                    'End If
                
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
                
                
                    Do '2015.04.20
                        If Not File_Open_Proc() Then
                            Exit Do
                        End If
                    Loop
                    
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

                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2015.04.20
                    'sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    'If sts Then
                    '    Call File_Error(sts, BtOpReset, "棚マスタ")
                    '    Beep
                    '    MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
                    'End If
                
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
                    
                    Do '2015.04.20
                        If Not File_Open_Proc() Then
                            Exit Do
                        End If
                    Loop

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.31
                
                    GoTo Item_Read

                
                End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.19
                
                
                
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                MsgBox "システム異常が発生しました。処理を中止して下さい。"
                Exit Function
        End Select
    Loop


    If ITEM_CHG_UPDATE_PROC() Then                  '2018.02.03
        Exit Function                               '2018.02.03
    End If                                          '2018.02.03


    Item_Update_Proc = False


End Function


Private Sub Text_LostFocus(Index As Integer)

    If Index = 0 Or Index = 2 Then
    
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
    
        
'        List2.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)        2013.01.28          '2014.02.18
        
        
'        If Trim(StrConv(GENSANREC.UPD_DATETIME, vbUnicode)) = "" Then                                               '2013.01.28        '2014.02.18
'            List2.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2013.01.28        '2014.02.18
'        Else                                                                                                        '2013.01.28        '2014.02.18
'            List2.AddItem StrConv(GENSANREC.UPD_DATETIME, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2013.01.28        '2014.02.18
'        End If
        
        
        If StrConv(GENSANREC.UPD_DATETIME, vbUnicode) > StrConv(GENSANREC.Ins_DateTime, vbUnicode) Then             '2014.02.18
            List2.AddItem StrConv(GENSANREC.UPD_DATETIME, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2014.02.18
        Else                                                                                                        '2014.02.18
            List2.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2014.02.18
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

'Private Sub File_Open_Proc()       2015.04.20 SUB-->Function
Private Function File_Open_Proc() As Integer
                                 
Dim c           As String * 128     '2013.8.23
                                
Dim sts         As Integer
                                
                                
    Call LOG_OUT(LOG_F, "File 再オープン処理 　開始")           '2015.03.26
    
    DoEvents
    
    
    File_Open_Proc = True
                                
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
    End If
                                
                                
                                
                                
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenRead) Then
        Beep
'        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'        Unload Me
        Exit Function
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Beep
'        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'        Unload Me
        Exit Function
    End If
    
    
                                '品目読み替えＯＰＥＮ   2018.02.03
    If ITEM_CHG_Open(BtOpenNomal) Then
        Beep
'        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'        Unload Me
        Exit Function
    End If
    
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.05.30
                                'PNマスタＯＰＥＮ
    If PN_M_Open(BtOpenRead) Then
        Beep
'        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'        Unload Me
        Exit Function
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                '原産国マスタＯＰＥＮ
    If GENSAN_Open(BtOpenNomal) Then
        Beep
'        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'        Unload Me
        Exit Function
    End If


                                '作業ログＯＰＥＮ   '2016.12.27
    If P_SAGYO_LOG_Open(BtOpenNomal) Then
        Unload Me
    End If




    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<    2012.01.31  -->　削除 2012.02.06
'                                'カントリーマスタ ＯＰＥＮ
'    If Country_Open(BtOpenRead) Then
'        Unload Me
'    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<    2012.01.31  -->　削除 2012.02.06


                                '欠品データＯＰＥＮ             2013.08.23
    
                                'ログファイル名取り込み
''2015.05.27 DEL    If GetIni("FILE", "KEPPIN", "SYS", c) Then
''                  Else
''                  If KEPPIN_Open(BtOpenNomal) Then
''                      Beep
''                      MsgBox "システム異常が発生しました。処理を中止して下さい。"
'''                     Unload Me
''                      Exit Function
''                  End If
''                  End If

    Call LOG_OUT(LOG_F, "File 再オープン処理 　正常終了")           '2015.03.26
    
    File_Open_Proc = False

End Function
' ------------------------------------------------------------------------
'       指定した精度の数値に四捨五入します。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に四捨五入された数値。
'
'       2012.03.25  frm より　移管
' ------------------------------------------------------------------------
Public Function ToHalfAdjust(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    If dValue > 0 Then
        ToHalfAdjust = Int(CDbl(dValue * dCoef + 0.5)) / dCoef
    Else
        ToHalfAdjust = Fix(CDbl(dValue * dCoef - 0.5)) / dCoef
    End If
End Function

' ------------------------------------------------------------------------
'       指定した精度の数値に切り上げします。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に切り上げられた数値。
'
'       2012.03.25  frm より　移管
'
' ------------------------------------------------------------------------
Public Function ToRoundUp(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    
        


    dCoef = (10 ^ iDigits)



    
    
    
    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
        Case Is > 0
            ToRoundUp = (Int(dValue * dCoef) + 1) / dCoef
        Case Is < 0
            ToRoundUp = (Fix(dValue * dCoef) - 1) / dCoef
        Case Else
            ToRoundUp = dValue
    End Select


'    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
'        Case Is > 0
'            ToRoundUp = (Int(dValue * dCoef + 0.9)) / dCoef
'        Case Is < 0
'            ToRoundUp = (Fix(dValue * dCoef - 0.9)) / dCoef
'        Case Else
'            ToRoundUp = dValue
'    End Select



End Function

' ------------------------------------------------------------------------
'       指定した精度の数値に切り捨てします。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に切り捨てられた数値。
'
'
'       2012.03.25  frm より　移管
'
' ------------------------------------------------------------------------
Public Function ToRoundDown(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
        Case Is > 0
            ToRoundDown = Int(dValue * dCoef) / dCoef
        Case Is < 0
            ToRoundDown = Fix(dValue * dCoef) / dCoef
        Case Else
            ToRoundDown = dValue
    End Select
End Function


Private Sub New_Print_Sub_Proc()
                                            
Dim Gyo         As Integer
Dim wk_IRI_QTY  As String * 5
                                            
Dim wkGENSAN    As String * 15
                                            
Dim wkSize      As Integer          '2019.02.26
                                            
'    Printer.NewPage
                                            
    On Error GoTo Err_Proc
                                            
    Printer.CurrentY = 0    '2019.02.22
                                            
    wkSize = 12
                                            
    For Gyo = 0 To 5
                                            
                                            
        If Len(Trim(Print_tbl(Gyo, 0).HIN_GAI)) = 0 Then
            Exit For
        End If


'------------------------------------------------   1行目   ------------------
'        Printer.Print Tab(20);                                 '2013.08.23      2014.02.18
                                                                
'        If Trim(Print_tbl(Gyo, 0).KEPPIN_QTY) = "" Then         '2013.08.23    2014.02.18
'            Printer.Print Tab(20);                              '2013.08.23    2014.02.18
'        Else                                                    '2013.08.23    2014.02.18
'            Printer.Print Tab(2);                               '2013.08.23    2014.02.18
'            Printer.Print "欠品";                               '2013.08.23    2014.02.18
'            Printer.Print "(";                                  '2013.08.23    2014.02.18
'            Printer.Print Trim(Print_tbl(Gyo, 0).KEPPIN_QTY);   '2013.08.23    2014.02.18
'            Printer.Print ")";                                  '2013.08.23    2014.02.18
'            Printer.Print Tab(20);                              '2013.08.23    2014.02.18
'        End If                                                  '2013.08.23    2014.02.18
        
        
'>>>>>>>>>>>>>>>>>>>>>>>>   2019.02.29
       
       
        wkSize = wkSize - Gyo
        If wkSize < 1 Then
        Else
             With NormalFont
                 .NAME = F1020501.FontName
                 .Size = wkSize
             End With
             Set Printer.Font = NormalFont
            
            
            
             If Gyo > 1 Then         '2019.02.26
                 Printer.Print       '2019.02.26
             End If                  '2019.02.26
        End If
'>>>>>>>>>>>>>>>>>>>>>>>>   2019.02.29
       
       
       
       
'>>>>>>>>>>>>>>>>>>>>>>>>   2019.02.22
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
'>>>>>>>>>>>>>>>>>>>>>>>>   2019.02.22


       
        Printer.Print Tab(20);                                 '2013.08.23      2014.02.18
        
        Printer.Print "入庫現品票";
        Printer.Print Tab(47);
        Printer.Print Trim(JGYOBU_NAME);

        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            
            Printer.Print Tab(80);                                 '2013.08.23      2014.02.18
'            If Trim(Print_tbl(Gyo, 1).KEPPIN_QTY) = "" Then         '2013.08.23    2014.02.18
'                Printer.Print Tab(80);                              '2013.08.23    2014.02.18
'            Else                                                    '2013.08.23    2014.02.18
'                Printer.Print Tab(62);                              '2013.08.23    2014.02.18
'                Printer.Print "欠品";                               '2013.08.23    2014.02.18
'                Printer.Print "(";                                  '2013.08.23    2014.02.18
'                Printer.Print Trim(Print_tbl(Gyo, 1).KEPPIN_QTY);   '2013.08.23    2014.02.18
'                Printer.Print ")";                                  '2013.08.23    2014.02.18
'                Printer.Print Tab(80);                              '2013.08.23    2014.02.18
'            End If                                                  '2013.08.23    2014.02.18
            Printer.Print "入庫現品票";
            Printer.Print Tab(104);
            Printer.Print Trim(JGYOBU_NAME)
        End If
'------------------------------------------------   2行目   ------------------
        If Gyo < 5 Then             '2019.02.26
        
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 6
            End With
            Set Printer.Font = NormalFont
            Printer.Print
        End If                      '2019.02.26
'------------------------------------------------   3行目   ------------------
        Set Printer.Font = Code39Font
        Printer.Print Tab(2);
        Printer.Print "*" + Trim(Left(Print_tbl(Gyo, 0).HIN_GAI, 16)) + "*";      '2019/11/08 外部品番16桁対応
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(23);
            Printer.Print "*" + Trim(Left(Print_tbl(Gyo, 1).HIN_GAI, 16)) + "*"  '2019/11/08 外部品番16桁対応
        End If
'------------------------------------------------   4行目   ------------------
        If Gyo < 5 Then             '2019.02.26
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 10
            End With
            Set Printer.Font = NormalFont
            Printer.Print
        End If
'------------------------------------------------   5行目   ------------------
       With NormalFont
            .NAME = F1020501.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
        
'        Printer.Print "品番";                          '2019/11/08 外部品番16桁対応
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 18
        End With
        Set Printer.Font = NormalFont
        Printer.Print Left(Print_tbl(Gyo, 0).HIN_GAI, 16); '2019/11/08 外部品番16桁対応
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
'       Printer.Print "(" & Left(Print_tbl(Gyo, 0).HIN_NAI, 14) & ")"; 2019/06/20 対内品番()削除
        Printer.Print Left(Print_tbl(Gyo, 0).HIN_NAI, 16);
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
            
            
'            Printer.Print "品番";        '2019/11/08 外部品番16桁対応
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 18
            End With
            Set Printer.Font = NormalFont
            Printer.Print Left(Print_tbl(Gyo, 1).HIN_GAI, 16); '2019/11/08 外部品番16桁対応
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
'           Printer.Print "(" & Left(Print_tbl(Gyo, 1).HIN_NAI, 14) & ")"  2019/06/20 対内品番()削除
            Printer.Print Left(Print_tbl(Gyo, 1).HIN_NAI, 16)
        End If
'------------------------------------------------   6行目   ------------------
        
        
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 4
        End With
        Set Printer.Font = NormalFont
        Printer.Print

'------------------------------------------------   7行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
        Printer.Print LeftB(Print_tbl(Gyo, 0).HIN_NAME, 80);  '2019/11/08 外部品番16桁対応
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
            Printer.Print LeftB(Print_tbl(Gyo, 1).HIN_NAME, 80) '2019/11/08 外部品番16桁対応
        End If
'------------------------------------------------   8行目   ------------------
        If Gyo < 5 Then             '2019.02.26
        
        
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 4
            End With
            Set Printer.Font = NormalFont
            Printer.Print

        End If
'------------------------------------------------   9行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print "　　入数" & ":";
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 18
        End With
        Set Printer.Font = NormalFont
        Printer.Print Format(Print_tbl(Gyo, 0).IRI_QTY, "#0");
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(30);
        Printer.Print "入荷日" & ":";
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 18
        End With
        Set Printer.Font = NormalFont
        Printer.Print text(ptxNyuka_YY).text & "/" & text(ptxNyuka_MM).text & "/" & text(ptxNyuka_DD).text;
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            Set Printer.Font = NormalFont
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 10
            End With
            Set Printer.Font = NormalFont
            
            Printer.Print "　　入数" & ":";
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 18
            End With
            Set Printer.Font = NormalFont
            Printer.Print Format(Print_tbl(Gyo, 1).IRI_QTY, "#0");
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 10
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(88);
            Printer.Print "入荷日" & ":";
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 18
            End With
            Set Printer.Font = NormalFont
            Printer.Print text(ptxNyuka_YY).text & "/" & text(ptxNyuka_MM).text & "/" & text(ptxNyuka_DD).text
        End If
'------------------------------------------------   10行目   ------------------
        If Gyo < 5 Then     '2019.02.26
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 4
            End With
            Set Printer.Font = NormalFont
            Printer.Print
        End If              '2019.02.26
'------------------------------------------------   11行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print "標準棚番" & ":" & Print_tbl(Gyo, 0).ST_SOKO & "-" & Print_tbl(Gyo, 0).ST_RETU & "-" & Print_tbl(Gyo, 0).ST_REN & "-" & Print_tbl(Gyo, 0).ST_DAN;
'        Printer.Print Tab(30);         '2019.02.27
        Printer.Print Tab(28);          '2019.02.27
'        Printer.Print "　備考" & ":" & RTrim(LeftB(Print_tbl(Gyo, 0).BIKOU, 40));  '2019.02.27
        Printer.Print "　備考" & ":" & RTrim(LeftB(Print_tbl(Gyo, 0).BIKOU, 40));   '2019.02.27 2019.04.05 20->40
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 10
            End With
            Printer.Print "標準棚番" & ":" & Print_tbl(Gyo, 1).ST_SOKO & "-" & Print_tbl(Gyo, 1).ST_RETU & "-" & Print_tbl(Gyo, 1).ST_REN & "-" & Print_tbl(Gyo, 1).ST_DAN;
'            Printer.Print Tab(88);                                                     '2019.02.27
'            Printer.Print "　備考" & ":" & RTrim(LeftB(Print_tbl(Gyo, 1).BIKOU, 40))   '2019.02.27
        
            Printer.Print Tab(86);                                                      '2019.02.27
            Printer.Print "　備考" & ":" & RTrim(LeftB(Print_tbl(Gyo, 1).BIKOU, 40))    '2019.02.27 2019.04.05 20->40
        
        End If
'------------------------------------------------   12行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        
        
        
        wkGENSAN = Left(Print_tbl(Gyo, 0).GENSAN, 13) & Right(Print_tbl(Gyo, 0).GENSAN, 2)
        
        
        
                
'>>>>>>>>>>>>>>>    2017.03.03
'        Printer.Print "　原産国" & ":" & Left(Print_tbl(Gyo, 0).GENSAN, 15);
        
        If GENSAN_KOKU_F = 2 Then
        Else
        
        
            If GENSAN_KOKU_F = 0 Then
                Printer.Print "　原産国" & ":" & wkGENSAN;
            Else
                If Print_tbl(Gyo, 0).GAI_BUHIN = "1" Or Print_tbl(Gyo, 0).GAI_BUHIN = "2" Or Print_tbl(Gyo, 0).GAI_BUHIN = "3" Then
                    Printer.Print "　原産国注意";
                Else
                    Printer.Print "　原産国" & ":" & wkGENSAN;
                End If
            End If


        End If
'>>>>>>>>>>>>>>>    2017.03.03
        
        
        
'>>>>>>>>>>>>>  2018.02.03

'        Printer.Print Tab(36);                                 '2019.02.27
        Printer.Print Tab(34);                                  '2019.02.27
        Printer.Print Left(Print_tbl(Gyo, 0).BIKOU2, 20);



'>>>>>>>>>>>>>  2018.02.03

'>>>>>>>>>>>>>>>    2017.03.03

        wkGENSAN = Left(Print_tbl(Gyo, 1).GENSAN, 13) & Right(Print_tbl(Gyo, 1).GENSAN, 2)


'        Printer.Print "　原産国" & ":" & Left(Print_tbl(Gyo, 0).GENSAN, 15);
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then        '2018.02.06
        Else                                                    '2018.02.06
            
            With NormalFont                                     '2018.02.06
                .NAME = F1020501.FontName                       '2018.02.06
                .Size = 14                                      '2018.02.06
            End With                                            '2018.02.06
            Set Printer.Font = NormalFont                       '2018.02.06
            Printer.Print Tab(43);                              '2018.02.06
            With NormalFont                                     '2018.02.06
                .NAME = F1020501.FontName                       '2018.02.06
                .Size = 10                                      '2018.02.06
            End With                                            '2018.02.06
            
                    
            If GENSAN_KOKU_F = 2 Then
            Else
            
                
                If GENSAN_KOKU_F = 0 Then
                    Printer.Print "　原産国" & ":" & wkGENSAN;
                Else
                    If Print_tbl(Gyo, 1).GAI_BUHIN = "1" Or Print_tbl(Gyo, 1).GAI_BUHIN = "2" Or Print_tbl(Gyo, 1).GAI_BUHIN = "3" Then
                        Printer.Print "　原産国注意";
                    Else
                        Printer.Print "　原産国" & ":" & wkGENSAN;
                    End If
                End If
        
            End If
        End If                                                  '2018.02.06
'>>>>>>>>>>>>>>>    2017.03.03
            
'>>>>>>>>>>>>>  2018.02.03

        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then        '2018.02.06
            Printer.Print                                       '2018.02.06
        Else                                                    '2018.02.06
'            Printer.Print Tab(94); '2019.02.27
            Printer.Print Tab(92);  '2019.02.27
            Printer.Print Left(Print_tbl(Gyo, 1).BIKOU2, 20)
        End If                                                  '2018.02.06
'>>>>>>>>>>>>>  2018.02.03
        
        
        
        Printer.Print Tab(30);
        Printer.Print "仕入先" & ":" & Print_tbl(Gyo, 0).SHIIRE_WORK_CENTER;
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 10
            End With
'            Printer.Print "　原産国" & ":" & Left(Print_tbl(Gyo, 1).GENSAN, 15);
            
        wkGENSAN = Left(Print_tbl(Gyo, 1).GENSAN, 13) & Right(Print_tbl(Gyo, 1).GENSAN, 2)
'        Printer.Print "　原産国" & ":" & Left(Print_tbl(Gyo, 0).GENSAN, 15);
        
        

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>2017.02.03
''>>>>>>>>>>>>>>>    2017.03.03
''        Printer.Print "　原産国" & ":" & Left(Print_tbl(Gyo, 0).GENSAN, 15);
'        If GENSAN_KOKU_F = 0 Then
'            Printer.Print "　原産国" & ":" & wkGENSAN;
'        Else
'            If Print_tbl(Gyo, 0).GAI_BUHIN = "1" Or Print_tbl(Gyo, 0).GAI_BUHIN = "2" Or Print_tbl(Gyo, 0).GAI_BUHIN = "3" Then
'                Printer.Print "　原産国注意";
'            Else
'                Printer.Print "　原産国" & ":" & wkGENSAN;
'            End If
'        End If
''>>>>>>>>>>>>>>>    2017.03.03
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>2017.02.03
            
            
            Printer.Print Tab(88);
            Printer.Print "仕入先" & ":" & Print_tbl(Gyo, 1).SHIIRE_WORK_CENTER;
        End If




'------------------------------------------------   13行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 8
        End With
        Set Printer.Font = NormalFont
        Printer.Print





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
            
            If Max_Gyo = 2 Then                 '2018.02.04
                With NormalFont                 '2018.02.04
                    .NAME = F1020501.FontName   '2018.02.04
                    .Size = 6                   '2018.02.04 2018.023 6--?2
                End With                        '2018.02.04
            Else                                '2018.02.04
                With NormalFont
                    .NAME = F1020501.FontName
                    .Size = 10
                End With
            End If                              '2018.02.23
            
            
            
            Set Printer.Font = NormalFont
            Printer.Print
'            End If                              '2018.02.23
            
            

'>>>>>>>>   2018.01.30
'
'            If Max_Gyo <> 2 Then
'
'                With NormalFont
'                    .NAME = F1020501.FontName
'                    .Size = 6
'                End With
'                Set Printer.Font = NormalFont
'                Printer.Print
'                Printer.Print
'            Else
'                With NormalFont
'                    .NAME = F1020501.FontName
'                    .Size = 4
'                End With
'                Set Printer.Font = NormalFont
'                Printer.Print
'                With NormalFont
'                    .NAME = F1020501.FontName
'                    .Size = 6
'                End With
'               Set Printer.Font = NormalFont
'                Printer.Print
'
'>>>>>>>>   2018.01.30
            
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


        'End If
    Next Gyo

    Exit Sub

Err_Proc:

    If Err.Number = 482 Then
        MsgBox "プリンターエラーが発生しました。"
    Else
        MsgBox "実行時エラー：" & Err.Number
    End If
                                            
                                            
                                            
End Sub

Private Sub New_Print_Sub_A5_Proc()
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
            .NAME = F1020501.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(20);                                 '2013.08.23      2014.02.18
                                                                
'        If Trim(Print_tbl(Gyo, 0).KEPPIN_QTY) = "" Then         '2013.08.23    2014.02.18
'            Printer.Print Tab(20);                              '2013.08.23    2014.02.18
'        Else                                                    '2013.08.23    2014.02.18
'            Printer.Print Tab(2);                               '2013.08.23    2014.02.18
'            Printer.Print "欠品";                               '2013.08.23    2014.02.18
'            Printer.Print "(";                                  '2013.08.23    2014.02.18
'            Printer.Print Trim(Print_tbl(Gyo, 0).KEPPIN_QTY);   '2013.08.23    2014.02.18
'            Printer.Print ")";                                  '2013.08.23    2014.02.18
'            Printer.Print Tab(20);                              '2013.08.23    2014.02.18
'        End If                                                  '2013.08.23    2014.02.18
        
        Printer.Print "入庫現品票";
        Printer.Print Tab(47);
        Printer.Print Trim(JGYOBU_NAME);

        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            
            Printer.Print Tab(80);                                 '2013.08.23      2014.02.18
'            If Trim(Print_tbl(Gyo, 1).KEPPIN_QTY) = "" Then         '2013.08.23    2014.02.18
'                Printer.Print Tab(80);                              '2013.08.23    2014.02.18
'            Else                                                    '2013.08.23    2014.02.18
'                Printer.Print Tab(62);                              '2013.08.23    2014.02.18
'                Printer.Print "欠品";                               '2013.08.23    2014.02.18
'                Printer.Print "(";                                  '2013.08.23    2014.02.18
'                Printer.Print Trim(Print_tbl(Gyo, 1).KEPPIN_QTY);   '2013.08.23    2014.02.18
'                Printer.Print ")";                                  '2013.08.23    2014.02.18
'                Printer.Print Tab(80);                              '2013.08.23    2014.02.18
'            End If                                                  '2013.08.23    2014.02.18
            Printer.Print "入庫現品票";
            Printer.Print Tab(104);
            Printer.Print Trim(JGYOBU_NAME)
        End If
'------------------------------------------------   2行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 6
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   3行目   ------------------
        Set Printer.Font = Code39Font
        Printer.Print Tab(2);
        Printer.Print "*" + Trim(Left(Print_tbl(Gyo, 0).HIN_GAI, 16)) + "*";   '2019/11/08 外部品番16桁対応
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(23);
            Printer.Print "*" + Trim(Left(Print_tbl(Gyo, 1).HIN_GAI, 16)) + "*" '2019/11/08 外部品番16桁対応
        End If
'------------------------------------------------   4行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 10                       '2018.02.26 10-6
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   5行目   ------------------
       With NormalFont
            .NAME = F1020501.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
'        Printer.Print "品番";           '2019/11/08 外部品番16桁対応
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 18                      '2018.02.26 18-16
        End With
        Set Printer.Font = NormalFont
        Printer.Print Left(Print_tbl(Gyo, 0).HIN_GAI, 16); '2019/11/08 外部品番16桁対応
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
'       Printer.Print "(" & Left(Print_tbl(Gyo, 0).HIN_NAI, 14) & ")"; 2019/06/20 対内品番の()削除
        Printer.Print Left(Print_tbl(Gyo, 0).HIN_NAI, 16);
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
            
            
'            Printer.Print "品番"; '2019/11/08 外部品番16桁対応
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 18                  '2018.02.26 18-16
            End With
            Set Printer.Font = NormalFont
            Printer.Print Left(Print_tbl(Gyo, 1).HIN_GAI, 16);   '2019/11/08 外部品番16桁対応
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
'           Printer.Print "(" & Left(Print_tbl(Gyo, 1).HIN_NAI, 14) & ")"  2019/06/20 対内品番の()削除
            Printer.Print Left(Print_tbl(Gyo, 1).HIN_NAI, 16)
            
        End If
'------------------------------------------------   6行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 2                    '2018.02.26 4-2
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   7行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
        Printer.Print LeftB(Print_tbl(Gyo, 0).HIN_NAME, 80); '2019/11/08 外部品番16桁対応
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
            Printer.Print LeftB(Print_tbl(Gyo, 1).HIN_NAME, 80) '2019/11/08 外部品番16桁対応
        End If
'------------------------------------------------   8行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 2                    '2018.02.26
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   9行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print "　　入数" & ":";
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 16                   '2018.02.26 18-16
        End With
        Set Printer.Font = NormalFont
        Printer.Print Format(Print_tbl(Gyo, 0).IRI_QTY, "#0");
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(30);
        Printer.Print "入荷日" & ":";
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 16                   '2018.02.26 18-16
        End With
        Set Printer.Font = NormalFont
        Printer.Print text(ptxNyuka_YY).text & "/" & text(ptxNyuka_MM).text & "/" & text(ptxNyuka_DD).text;
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            Set Printer.Font = NormalFont
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 10
            End With
            Set Printer.Font = NormalFont
            
            Printer.Print "　　入数" & ":";
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 18                  '2018.02.26 18-16
            End With
            Set Printer.Font = NormalFont
            Printer.Print Format(Print_tbl(Gyo, 1).IRI_QTY, "#0");
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 10
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(88);
            Printer.Print "入荷日" & ":";
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 18                  '2018.02.26 18-16
            End With
            Set Printer.Font = NormalFont
            Printer.Print text(ptxNyuka_YY).text & "/" & text(ptxNyuka_MM).text & "/" & text(ptxNyuka_DD).text
        End If
'------------------------------------------------   10行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 4                       '2018.02.26 4-2
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   11行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F1020501.FontName
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
                .NAME = F1020501.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 10
            End With
            Printer.Print "標準棚番" & ":" & Print_tbl(Gyo, 1).ST_SOKO & "-" & Print_tbl(Gyo, 1).ST_RETU & "-" & Print_tbl(Gyo, 1).ST_REN & "-" & Print_tbl(Gyo, 1).ST_DAN;
            Printer.Print Tab(88);
            Printer.Print "　備考" & ":" & RTrim(LeftB(Print_tbl(Gyo, 1).BIKOU, 40))
        End If
'------------------------------------------------   12行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        
        
        
        wkGENSAN = Left(Print_tbl(Gyo, 0).GENSAN, 13) & Right(Print_tbl(Gyo, 0).GENSAN, 2)
        
        
        
                
'>>>>>>>>>>>>>>>    2017.03.03
'        Printer.Print "　原産国" & ":" & Left(Print_tbl(Gyo, 0).GENSAN, 15);
        If GENSAN_KOKU_F = 0 Then
            Printer.Print "　原産国" & ":" & wkGENSAN;
        Else
            If Print_tbl(Gyo, 0).GAI_BUHIN = "1" Or Print_tbl(Gyo, 0).GAI_BUHIN = "2" Or Print_tbl(Gyo, 0).GAI_BUHIN = "3" Then
                Printer.Print "　原産国注意";
            Else
                Printer.Print "　原産国" & ":" & wkGENSAN;
            End If
        End If
'>>>>>>>>>>>>>>>    2017.03.03
        
        
'>>>>>>>>>>>>>  2018.02.03

        Printer.Print Tab(36);
        Printer.Print Left(Print_tbl(Gyo, 0).BIKOU2, 20);



'>>>>>>>>>>>>>  2018.02.03

'>>>>>>>>>>>>>>>    2017.03.03
'        Printer.Print "　原産国" & ":" & Left(Print_tbl(Gyo, 0).GENSAN, 15);
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then        '2018.02.06
        Else                                                    '2018.02.06
            
            With NormalFont                                     '2018.02.06
                .NAME = F1020501.FontName                       '2018.02.06
                .Size = 14                                      '2018.02.06
            End With                                            '2018.02.06
            Set Printer.Font = NormalFont                       '2018.02.06
            Printer.Print Tab(43);                              '2018.02.06
            With NormalFont                                     '2018.02.06
                .NAME = F1020501.FontName                       '2018.02.06
                .Size = 10                                      '2018.02.06
            End With                                            '2018.02.06
            
            
            
            
            
            
            If GENSAN_KOKU_F = 0 Then
                Printer.Print "　原産国" & ":" & wkGENSAN;
            Else
                If Print_tbl(Gyo, 0).GAI_BUHIN = "1" Or Print_tbl(Gyo, 0).GAI_BUHIN = "2" Or Print_tbl(Gyo, 0).GAI_BUHIN = "3" Then
                    Printer.Print "　原産国注意";
                Else
                    Printer.Print "　原産国" & ":" & wkGENSAN;
                End If
            End If
        End If                                                  '2018.02.06
'>>>>>>>>>>>>>>>    2017.03.03
            
'>>>>>>>>>>>>>  2018.02.03

        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then        '2018.02.06
            Printer.Print                                       '2018.02.06
        Else                                                    '2018.02.06
            Printer.Print Tab(94);
            Printer.Print Left(Print_tbl(Gyo, 0).BIKOU2, 20)
        End If                                                  '2018.02.06
'>>>>>>>>>>>>>  2018.02.03
        
        
        
        Printer.Print Tab(30);
        Printer.Print "仕入先" & ":" & Print_tbl(Gyo, 0).SHIIRE_WORK_CENTER;
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            With NormalFont
                .NAME = F1020501.FontName
                .Size = 10
            End With
'            Printer.Print "　原産国" & ":" & Left(Print_tbl(Gyo, 1).GENSAN, 15);
            
        wkGENSAN = Left(Print_tbl(Gyo, 1).GENSAN, 13) & Right(Print_tbl(Gyo, 1).GENSAN, 2)
'        Printer.Print "　原産国" & ":" & Left(Print_tbl(Gyo, 0).GENSAN, 15);
        
        

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>2017.02.03
''>>>>>>>>>>>>>>>    2017.03.03
''        Printer.Print "　原産国" & ":" & Left(Print_tbl(Gyo, 0).GENSAN, 15);
'        If GENSAN_KOKU_F = 0 Then
'            Printer.Print "　原産国" & ":" & wkGENSAN;
'        Else
'            If Print_tbl(Gyo, 0).GAI_BUHIN = "1" Or Print_tbl(Gyo, 0).GAI_BUHIN = "2" Or Print_tbl(Gyo, 0).GAI_BUHIN = "3" Then
'                Printer.Print "　原産国注意";
'            Else
'                Printer.Print "　原産国" & ":" & wkGENSAN;
'            End If
'        End If
''>>>>>>>>>>>>>>>    2017.03.03
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>2017.02.03
            
            
            Printer.Print Tab(88);
            Printer.Print "仕入先" & ":" & Print_tbl(Gyo, 1).SHIIRE_WORK_CENTER;
        End If




'------------------------------------------------   13行目   ------------------
        With NormalFont
            .NAME = F1020501.FontName
            .Size = 4                       '2018.02.26 8-4
        End With
        Set Printer.Font = NormalFont
        Printer.Print






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
            
            If Max_Gyo = 2 Then                 '2018.02.04
                With NormalFont                 '2018.02.04
                    .NAME = F1020501.FontName   '2018.02.04
                    .Size = 14                   '2018.02.04 2018.023 6--10
                End With                        '2018.02.04
            Else                                '2018.02.04
                With NormalFont
                    .NAME = F1020501.FontName
                    .Size = 10
                End With
            End If                              '2018.02.23
            
            
            
            Set Printer.Font = NormalFont
            Printer.Print
'            End If                              '2018.02.23
            
            

'>>>>>>>>   2018.01.30
'
'            If Max_Gyo <> 2 Then
'
'                With NormalFont
'                    .NAME = F1020501.FontName
'                    .Size = 6
'                End With
'                Set Printer.Font = NormalFont
'                Printer.Print
'                Printer.Print
'            Else
'                With NormalFont
'                    .NAME = F1020501.FontName
'                    .Size = 4
'                End With
'                Set Printer.Font = NormalFont
'                Printer.Print
'                With NormalFont
'                    .NAME = F1020501.FontName
'                    .Size = 6
'                End With
'               Set Printer.Font = NormalFont
'                Printer.Print
'
'>>>>>>>>   2018.01.30
            
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


        'End If
    Next Gyo

    Exit Sub

Err_Proc:

    If Err.Number = 482 Then
        MsgBox "プリンターエラーが発生しました。"
    Else
        MsgBox "実行時エラー：" & Err.Number
    End If

End Sub

Private Function ITEM_CHG_UPDATE_PROC() As Integer
' ------------------------------------------------------------------------
'       品目読み替え　更新
'
' ------------------------------------------------------------------------
Dim sts     As Integer
Dim com     As Integer
Dim wk_Kbn  As String * 1
Dim ans     As Integer

    ITEM_CHG_UPDATE_PROC = True
    
    
    If Trim(text(ptxBikou2).text) = "" Then
        ITEM_CHG_UPDATE_PROC = False
        Exit Function
    End If
    
    
Item_Read:
    
    Call UniCode_Conv(K0_ITEM_CHG.N_JGYOBU, Last_JGYOBU)
    If Combo(0).text = NAIGAI1$ Then
        Call UniCode_Conv(K0_ITEM_CHG.N_NAIGAI, NAIGAI_NAI$)
        wk_Kbn = NAIGAI_NAI
    Else
        Call UniCode_Conv(K0_ITEM_CHG.N_NAIGAI, NAIGAI_GAI$)
        wk_Kbn = NAIGAI_GAI
    End If
    Call UniCode_Conv(K0_ITEM_CHG.N_HIN_GAI, RTrim(text(ptxHin_Gai).text))
    sts = BTRV(BtOpGetEqual, ITEM_CHG_POS, ITEM_CHG_REC, Len(ITEM_CHG_REC), K0_ITEM_CHG, Len(K0_ITEM_CHG), 0)
    Select Case sts
        Case BtNoErr
            com = BtOpUpdate
        Case BtErrKeyNotFound
            com = BtOpInsert
        Case Else
            If sts > 3000 Or sts = 3 Then
                Call File_Error(sts, BtOpGetEqual, "品目読み替え", 0)
                Do
                    If Not File_Open_Proc Then
                        Exit Do
                    End If
                Loop
                
                GoTo Item_Read

                
            End If
                        
            Call File_Error(sts, BtOpGetEqual, "品目読み替え")
            Beep
            MsgBox "システム異常が発生しました。処理を中止して下さい。"
            Unload Me
    End Select
    
        
    If com = BtOpInsert Then
        Call UniCode_Conv(ITEM_CHG_REC.N_JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(ITEM_CHG_REC.N_NAIGAI, wk_Kbn)
        Call UniCode_Conv(ITEM_CHG_REC.N_HIN_GAI, RTrim(text(ptxHin_Gai).text))
    
    
        Call UniCode_Conv(ITEM_CHG_REC.HIN_NAME, text(ptxHin_Name).text)
    
    
    End If
    
    Call UniCode_Conv(ITEM_CHG_REC.O_HIN_GAI, RTrim(text(ptxBikou2).text))
    
Debug.Print StrConv(ITEM_CHG_REC.O_HIN_GAI, vbUnicode)
    
    Do
        sts = BTRV(com, ITEM_CHG_POS, ITEM_CHG_REC, Len(ITEM_CHG_REC), K0_ITEM_CHG, Len(K0_ITEM_CHG), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                ans = MsgBox("他端末でデータ使用中です。<ITEM_CHG.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    ITEM_CHG_UPDATE_PROC = False
                    Exit Function
                End If
            Case Else
                If sts > 3000 Or sts = 3 Then
                    Call File_Error(sts, BtOpUpdate, "品目読み替え", 0)
                    Do
                        If Not File_Open_Proc() Then
                            Exit Do
                        End If
                    Loop
               
                    GoTo Item_Read
               End If
                
                Call File_Error(sts, com, "品目読み替え")
                MsgBox "システム異常が発生しました。処理を中止して下さい。"
                Exit Function
        End Select
    Loop
    
    
    
    
    
    
    ITEM_CHG_UPDATE_PROC = False


End Function

