VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form PM000402 
   Caption         =   "パーツラベル発行"
   ClientHeight    =   11370
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   14715
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
   ScaleHeight     =   11370
   ScaleWidth      =   14715
   StartUpPosition =   2  '画面の中央
   Begin VB.ListBox List2 
      Height          =   540
      Left            =   5940
      Sorted          =   -1  'True
      TabIndex        =   71
      Top             =   5520
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   4
      ItemData        =   "PM000402.frx":0000
      Left            =   1800
      List            =   "PM000402.frx":0002
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   25
      Top             =   4740
      Width           =   2805
   End
   Begin VB.CheckBox Check1 
      Caption         =   "原産国印字する"
      Height          =   375
      Index           =   4
      Left            =   7470
      TabIndex        =   27
      Top             =   4740
      Width           =   2115
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Index           =   11
      Left            =   4845
      MaxLength       =   20
      TabIndex        =   26
      Top             =   4740
      Width           =   2490
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   1470
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   120
      Width           =   1725
   End
   Begin VB.Frame Frame1 
      Caption         =   "ﾗﾍﾞﾙ指示"
      Height          =   2895
      Left            =   5775
      TabIndex        =   63
      Top             =   6840
      Width           =   3615
      Begin VB.TextBox Text1 
         Alignment       =   1  '右揃え
         Height          =   375
         IMEMode         =   2  'ｵﾌ
         Index           =   14
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   2  'ｵﾌ
         Index           =   18
         Left            =   2940
         MaxLength       =   1
         TabIndex        =   8
         Top             =   2280
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   2  'ｵﾌ
         Index           =   17
         Left            =   1320
         MaxLength       =   11
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   2  'ｵﾌ
         Index           =   15
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   2  'ｵﾌ
         Index           =   16
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '右揃え
         Height          =   375
         IMEMode         =   2  'ｵﾌ
         Index           =   13
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "数量"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   70
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "日付"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   67
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "ｵｰﾀﾞｰ��"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   66
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "ｱｲﾃﾑ��"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   65
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "枚数"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   64
         Top             =   480
         Width           =   615
      End
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4575
      Index           =   0
      Left            =   1755
      TabIndex        =   28
      Top             =   5160
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   8070
      _Version        =   393217
      TextRTF         =   $"PM000402.frx":0004
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Index           =   4
      Left            =   5520
      MaxLength       =   5
      TabIndex        =   14
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Index           =   3
      Left            =   1800
      MaxLength       =   30
      TabIndex        =   13
      Top             =   2040
      Width           =   3735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "枚数ラベル"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   6360
      TabIndex        =   21
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "適用機種ラベル"
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   20
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "プラ"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   19
      Top             =   3480
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "紙"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   18
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Index           =   7
      Left            =   5400
      MaxLength       =   20
      TabIndex        =   17
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Index           =   6
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   16
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Index           =   5
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   15
      Top             =   2520
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   3
      Left            =   1800
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   12
      Top             =   1560
      Width           =   5325
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   2
      Left            =   1800
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   11
      Top             =   1080
      Width           =   5325
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Index           =   10
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   24
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Index           =   9
      Left            =   4440
      MaxLength       =   10
      TabIndex        =   23
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Index           =   8
      Left            =   1785
      MaxLength       =   10
      TabIndex        =   22
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Index           =   12
      Left            =   9600
      MaxLength       =   25
      TabIndex        =   29
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   10800
      MaxLength       =   30
      TabIndex        =   10
      Top             =   600
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   1
      Left            =   1470
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   600
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Index           =   1
      Left            =   5640
      MaxLength       =   40
      TabIndex        =   9
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   2400
      MaxLength       =   20
      TabIndex        =   2
      Top             =   600
      Width           =   2535
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
      TabIndex        =   44
      Top             =   9960
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
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   9960
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
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   9960
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
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   9960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "外装"
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
      TabIndex        =   40
      Top             =   9960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "JAN"
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
      TabIndex        =   39
      Top             =   9960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ｱｲﾃﾑ"
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
      TabIndex        =   38
      Top             =   9960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ﾗﾍﾞﾙ"
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
      TabIndex        =   37
      Top             =   9960
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
      TabIndex        =   36
      Top             =   9960
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
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   9960
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
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   9960
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
      TabIndex        =   33
      Top             =   9960
      Width           =   855
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4575
      Index           =   2
      Left            =   9600
      TabIndex        =   31
      Top             =   2640
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   8070
      _Version        =   393217
      TextRTF         =   $"PM000402.frx":00C2
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1575
      Index           =   3
      Left            =   9600
      TabIndex        =   32
      Top             =   7680
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2778
      _Version        =   393217
      TextRTF         =   $"PM000402.frx":0180
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   735
      Index           =   1
      Left            =   9600
      TabIndex        =   30
      Top             =   1680
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"PM000402.frx":023E
   End
   Begin VB.Label Label 
      Caption         =   "単価空白発行:99999999"
      Height          =   255
      Index           =   19
      Left            =   6720
      TabIndex        =   74
      Top             =   4320
      Width           =   2655
   End
   Begin VB.Label lblUpd_DateTime 
      Height          =   255
      Left            =   11610
      TabIndex        =   73
      Top             =   10200
      Width           =   2535
   End
   Begin VB.Label lblUpd_Tanto 
      Height          =   255
      Left            =   11610
      TabIndex        =   72
      Top             =   9780
      Width           =   2535
   End
   Begin VB.Label Label 
      Caption         =   "原産国"
      Height          =   255
      Index           =   18
      Left            =   735
      TabIndex        =   69
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "事業部"
      Height          =   255
      Index           =   17
      Left            =   525
      TabIndex        =   68
      Top             =   240
      Width           =   795
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
      TabIndex        =   62
      Top             =   10320
      Width           =   180
   End
   Begin VB.Label Label 
      Caption         =   "備考"
      Height          =   255
      Index           =   16
      Left            =   1200
      TabIndex        =   61
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "作業指示"
      Height          =   255
      Index           =   15
      Left            =   720
      TabIndex        =   60
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "適用機種備考"
      Height          =   255
      Index           =   14
      Left            =   9600
      TabIndex        =   59
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "棚番(2)"
      Height          =   255
      Index           =   13
      Left            =   4440
      TabIndex        =   58
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "棚番(1)"
      Height          =   255
      Index           =   12
      Left            =   840
      TabIndex        =   57
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "入り数"
      Height          =   255
      Index           =   11
      Left            =   840
      TabIndex        =   56
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "事業部名"
      Height          =   255
      Index           =   10
      Left            =   960
      TabIndex        =   55
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "会社名"
      Height          =   255
      Index           =   9
      Left            =   960
      TabIndex        =   54
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "価格(3)"
      Height          =   255
      Index           =   8
      Left            =   6000
      TabIndex        =   53
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "価格(2)"
      Height          =   255
      Index           =   7
      Left            =   3360
      TabIndex        =   52
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "価格(1)"
      Height          =   255
      Index           =   6
      Left            =   720
      TabIndex        =   51
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "機種(3)"
      Height          =   255
      Index           =   5
      Left            =   8640
      TabIndex        =   50
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "機種(2)"
      Height          =   255
      Index           =   4
      Left            =   8640
      TabIndex        =   49
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "機種(1)"
      Height          =   255
      Index           =   3
      Left            =   8640
      TabIndex        =   48
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "PART　NAME"
      Height          =   255
      Index           =   1
      Left            =   9480
      TabIndex        =   47
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "品目コード"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   46
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "品名"
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   45
      Top             =   720
      Width           =   495
   End
End
Attribute VB_Name = "PM000402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'テキスト用添字
Private Const ptxHIN_GAI% = 0               '品番
Private Const ptxHIN_NAME% = 1              '品名
Private Const ptxL_HIN_NAME_E% = 2          '品名E
Private Const ptxL_BIKOU% = 3               '備考
Private Const ptxL_BIKOU3% = 4              '備考３
Private Const ptxL_IRI_QTY% = 5             '入り数
Private Const ptxL_TANA1% = 6               '棚番(1)
Private Const ptxL_TANA2% = 7               '棚番(2)
Private Const ptxL_URIKIN1% = 8             '価格(1)
Private Const ptxL_URIKIN2% = 9             '価格(2)
Private Const ptxL_URIKIN3% = 10            '価格(3)

Private Const ptxGENSANKOKU% = 11           '原産国 2008.06.12



Private Const ptxL_KISHU1% = 12             '機種(1)
'Private Const ptxL_KISHU2% = 12             '機種(2)




Private Const ptxL_MAISU% = 13              'ﾗﾍﾞﾙ枚数

Private Const ptxL_QTY% = 14                '数量   2008.10.03


Private Const ptxL_ORDERNO% = 15            'ｵｰﾀﾞｰ��
Private Const ptxL_ITEMNO% = 16             'ｱｲﾃﾑ��
Private Const ptxL_PRI_DATE% = 17           '印刷日付

Private Const ptxL_MARK% = 18               '再梱包ﾏｰｸ  2007.11.08

'リッチテキスト用添字
Private Const prchL_SAGYO_SHIJI% = 0        '作業指示
Private Const prchL_KISHU2% = 1             '機種(2)
Private Const prchL_KISHU3% = 2             '機種(3)
Private Const prchL_KISHU_BIKOU% = 3        '適用機種備考


'コンボ用添字
Private Const pcmbJGYOBU% = 0               '事業部     '2008.06.12


Private Const pcmbNAIGAI% = 1               '国内外
Private Const pcmbL_KAISHA% = 2             '会社名
Private Const pcmbL_JGYOBU% = 3             '事業部名
Private Const pcmbGENSAN% = 4               '原産国



'チェック用添字
Private Const pchkL_PAPER% = 0              '紙
Private Const pchkL_PLASTIC% = 1            'ﾌﾟﾗｽﾁｯｸ
Private Const pchkL_LABEL% = 2              '適用機種ﾗﾍﾞﾙ
Private Const pchkL_MAISU% = 3              '枚数ﾗﾍﾞﾙ

Private Const pchkGENSANKOKU% = 4           '原産国印字有無


'ｺﾏﾝﾄﾞﾎﾞﾀﾝ特殊処理
Private Const pcmdLabel% = 4                'ﾗﾍﾞﾙ印刷指示
Private Const pcmdItem% = 5                 'ｱｲﾃﾑ印刷指示
Private Const pcmdJan% = 6                  'JAN印刷指示
Private Const pcmdGAISO% = 7                '外装印刷指示


Private GENSANKOKU_FLG  As String * 1       '原産国　印字有無


Private INIT_FLG        As Boolean



Private KAISYA_CHK_F    As Boolean          '会社／事業部のエラーﾁｪｯｸ有無 2008.09.26

Private PRINT_CHECK_F   As Boolean          '印刷制御有無   2008.09.26



Private GENSANKOKU_CHECK_TBL _
                        As Variant          '原産国ﾁｪｯｸ有無（事業部） 2009.03.28



Private TANKA_SPACE_F   As String           '2010.02.02

Private KISHU_CHECK     As Integer          '2012.09.03 機種ﾁｪｯｸ


Private KAISYA_RESTRICT_F   As Boolean      '2013.03.29

Private GAI_BUHIN_CHECK As Boolean          '2016.01.29


'Private RETURN_F        As Boolean          '2013.03.29


'Private Const Last_Update_Day$ = "(PM00040 2016.02.01 10:45)"
Private Const Last_Update_Day$ = "(PM00040 2016.02.01 14:45)"

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PM000402.MousePointer = vbHourglass

    Call Ctrl_Lock(PM000402)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM000402)


    PM000402.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
Dim com     As Integer
Dim ans     As Integer
Dim sts     As Integer

Dim i       As Integer
Dim j       As Integer
Dim k       As Integer
    
    
    
    Error_Check_Proc = True
    
    
    
    Select Case Mode
        
        Case ptxHIN_GAI      '品番
            
            If Trim(Text1(ptxHIN_GAI).Text) = "" Then
                MsgBox "入力した項目はエラーです。"
                Text1(ptxHIN_GAI).SetFocus
                Exit Function
            End If
            
        
        
            If Last_JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode) And _
                Right(Combo1(pcmbNAIGAI), 1) = StrConv(ITEMREC.NAIGAI, vbUnicode) And _
                Trim(Text1(ptxHIN_GAI).Text) = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
            Else
                Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI), 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
Start_PRoc:
            
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    
                        Call Item_Disp_Proc(Last_JGYOBU & Right(Combo1(pcmbNAIGAI), 1) & Text1(ptxHIN_GAI).Text)
                    
                    Case BtErrKeyNotFound
                    
                    
                    
                        For i = 0 To UBound(JGYOBU_T)
                            For j = 0 To Combo1(pcmbNAIGAI).ListCount - 1
                                Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU_T(i).CODE)
                                Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI).List(j), 1))
                                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
        
                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                Select Case sts
                                    Case BtNoErr
        
                                        
                                        
                                        For k = 0 To Combo1(pcmbJGYOBU).ListCount - 1
                                        
                                            
                                            If Right(Combo1(pcmbJGYOBU).List(k), 1) = JGYOBU_T(i).CODE Then
                                            
                                                Combo1(pcmbJGYOBU).ListIndex = k
                                                
                                                Last_JGYOBU = JGYOBU_T(i).CODE
                                                Exit For
                                            
                                            End If
                                        
                                        Next k
                                    
                                    
                                        For k = 0 To Combo1(pcmbNAIGAI).ListCount - 1
                                        
                                            
                                            If Right(Combo1(pcmbNAIGAI).List(k), 1) = StrConv(ITEMREC.NAIGAI, vbUnicode) Then
                                            
                                                Combo1(pcmbNAIGAI).ListIndex = k
                                                Exit For
                                            
                                            End If
                                        
                                        Next k
                                        
                                        Call Item_Disp_Proc(Last_JGYOBU & Right(Combo1(pcmbNAIGAI), 1) & Text1(ptxHIN_GAI).Text)
                                        Exit For
        
                                    Case BtErrKeyNotFound
                                        Exit For
        
                                    Case Else
                                        
                                        
                                        
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.19 ﾃｽﾄﾛｼﾞｯｸ
                                        If sts > 3000 Or sts = 3 Then
                                        
                                            
                            '                MsgBox "てすと"
                                            
                                            
                                            Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                            sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                            If sts Then
                                                Call File_Error(sts, BtOpReset, "")
                                            End If
                                        
                                        
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.23
                            '                                            '品目マスタＯＰＥＮ
                            '                If ITEM_Open(BtOpenNomal) Then
                            '                    Unload Me
                            '                End If
                            '
                            '                                            '品目マスタＯＰＥＮ
                            '                If L_ITEM_Open(BtOpenNomal) Then
                            '                    Unload Me
                            '                End If
                            '
                            '                                            '原産国マスタＯＰＥＮ
                            '                If GENSAN_Open(BtOpenNomal) Then
                            '                    Beep
                            '                    MsgBox "システム異常が発生しました。処理を中止して下さい。"
                            '                    Unload Me
                            '                End If
                            '                                            'コードマスタＯＰＥＮ
                            '                If P_CODE_Open(BtOpenNomal) Then
                            '                    Unload Me
                            '                End If
                            '                                            '受払先マスタ（仕入先）ＯＰＥＮ
                            '                If P_UKEHARAI_Open(BtOpenNomal) Then
                            '                    Unload Me
                            '                End If
                            '
                            '                Call P_CODE_TBL_Proc
                            '
                            '
                            '                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.06.01
                            '                                            'PNマスタＯＰＥＮ
                            '                If PN_M_Open(0) Then
                            '                    Beep
                            '                    MsgBox "システム異常が発生しました。処理を中止して下さい。"
                            '                    Unload Me
                            '                End If
                            '                '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                            '
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.23
                                        
                                            '2012.01.23
                                            Call File_Open_Proc
                                        
                                            GoTo Start_PRoc
                                        
                                        End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.19 ﾃｽﾄﾛｼﾞｯｸ
                                        
                                        
                                        
                                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                        
                                        
                                        Exit Function
                                End Select
        
        
                            Next j
                    
                    
                            If sts = BtNoErr Then
                    
                            
                                Exit For
                            
                            End If
                    
                    
                        Next i
                    
                    
                        
                        If i > UBound(JGYOBU_T) Then
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> MT_2009.06.01
                            'MsgBox "入力したコードは、未登録です。"
                            'Exit Function
                                
                            If PN_CHK(Text1(ptxHIN_GAI), "G", "PLABEL", 1) Then
                                ''MsgBox "入力したコードは、未登録です。"
                                
                                Exit Function
                            End If
                            
                            Call Item_Disp_Proc(Last_JGYOBU & Right(Combo1(pcmbNAIGAI), 1) & Text1(ptxHIN_GAI).Text)
                            
                        End If
                    Case Else
                        
                        
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.19 ﾃｽﾄﾛｼﾞｯｸ
                        If sts > 3000 Or sts = 3 Then
                        
                            
            '                MsgBox "てすと"
                            
                            
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                            sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            If sts Then
                                Call File_Error(sts, BtOpReset, "")
                            End If
                        
                        
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.23
            '                                            '品目マスタＯＰＥＮ
            '                If ITEM_Open(BtOpenNomal) Then
            '                    Unload Me
            '                End If
            '
            '                                            '品目マスタＯＰＥＮ
            '                If L_ITEM_Open(BtOpenNomal) Then
            '                    Unload Me
            '                End If
            '
            '                                            '原産国マスタＯＰＥＮ
            '                If GENSAN_Open(BtOpenNomal) Then
            '                    Beep
            '                    MsgBox "システム異常が発生しました。処理を中止して下さい。"
            '                    Unload Me
            '                End If
            '                                            'コードマスタＯＰＥＮ
            '                If P_CODE_Open(BtOpenNomal) Then
            '                    Unload Me
            '                End If
            '                                            '受払先マスタ（仕入先）ＯＰＥＮ
            '                If P_UKEHARAI_Open(BtOpenNomal) Then
            '                    Unload Me
            '                End If
            '
            '                Call P_CODE_TBL_Proc
            '
            '
            '                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.06.01
            '                                            'PNマスタＯＰＥＮ
            '                If PN_M_Open(0) Then
            '                    Beep
            '                    MsgBox "システム異常が発生しました。処理を中止して下さい。"
            '                    Unload Me
            '                End If
            '                '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.23
                        
                        
                            '2012.01.23
                            Call File_Open_Proc
                        
                        
                            GoTo Start_PRoc
                        
                        End If
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.19 ﾃｽﾄﾛｼﾞｯｸ
                        
                        
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                End Select
            End If
        
            
        
        Case ptxL_IRI_QTY          '入り数
        
            If Trim(Text1(ptxL_IRI_QTY).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxL_IRI_QTY).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(ptxL_IRI_QTY).SetFocus
                    Exit Function
                Else
                
                    Text1(ptxL_IRI_QTY).Text = Format(CLng(Text1(ptxL_IRI_QTY).Text), "#0")
                
                End If
            End If
        
        Case ptxL_URIKIN1          '価格(1)
        
            If Trim(Text1(ptxL_URIKIN1).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxL_URIKIN1).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(ptxL_URIKIN1).SetFocus
                    Exit Function
                Else
                
                    Text1(ptxL_URIKIN1).Text = Format(CLng(Text1(ptxL_URIKIN1).Text), "#0")
                
                End If
            End If
        
        Case ptxL_URIKIN2          '価格(2)
        
            If Trim(Text1(ptxL_URIKIN2).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxL_URIKIN2).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(ptxL_URIKIN2).SetFocus
                    Exit Function
                Else
                
                    Text1(ptxL_URIKIN2).Text = Format(CLng(Text1(ptxL_URIKIN2).Text), "#0")
                
                End If
            End If
        
        Case ptxL_URIKIN3          '価格(3)
        
            If Trim(Text1(ptxL_URIKIN3).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxL_URIKIN3).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(ptxL_URIKIN3).SetFocus
                    Exit Function
                Else
                
                    Text1(ptxL_URIKIN3).Text = Format(CLng(Text1(ptxL_URIKIN3).Text), "#0")
                
                End If
            End If
        
        
        
    End Select
        
    Error_Check_Proc = False


End Function

Private Function Item_Disp_Proc(CODE As String) As Integer
'----------------------------------------------------------------------------
'                   画面表示
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim i       As Integer

Dim L_CODE  As String

Dim L_KISHU2    As String * 52

Dim L_KISHU_BIKOU   As String * 450     '2012.01.13
Dim L_KISHU3        As String * 150     '2012.01.13
Dim L_SAGYO_SHIJI   As String * 450     '2012.01.13
    
    Item_Disp_Proc = True

Start_PRoc:
    
    '品目ﾏｽﾀ読み込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo1(pcmbJGYOBU).Text, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            
            'ﾚｺｰﾄﾞ内容の表示
                                            '品目ｺｰﾄﾞ
            Text1(ptxHIN_GAI).Text = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                            '品名
            Text1(ptxHIN_NAME).Text = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                                            '品名E
            Text1(ptxL_HIN_NAME_E).Text = Trim(StrConv(ITEMREC.L_HIN_NAME_E, vbUnicode))
                                            '会社名
            
            
            '>>>>>>>>>>>>>>>>>>>    ﾃﾞﾌｫﾙﾄ会社名    2013.04.02
            If KAISYA_DEF_VALUE <> "" Then
                If Trim(StrConv(ITEMREC.L_KAISHA_CODE, vbUnicode)) = "" Then
                    Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, KAISYA_DEF_VALUE)
                End If
            End If
            '>>>>>>>>>>>>>>>>>>>    ﾃﾞﾌｫﾙﾄ会社名    2013.04.02
            
            
            
            If Trim(StrConv(ITEMREC.L_KAISHA_CODE, vbUnicode)) = "" Then
                Combo1(pcmbL_KAISHA).ListIndex = -1
            Else
                
                
                For i = 0 To Combo1(pcmbL_KAISHA).ListCount - 1
                    
                    L_CODE = Left(Right(Combo1(pcmbL_KAISHA).List(i), 4), 2)
                    If Trim(L_CODE) = "" Then
                        L_CODE = Right(Combo1(pcmbL_KAISHA).List(i), 2)
                    End If
                    
                    
                    If StrConv(ITEMREC.L_KAISHA_CODE, vbUnicode) = L_CODE Then
                        Combo1(pcmbL_KAISHA).ListIndex = i
                        Exit For
                        
                    End If
                
                
                Next i
            End If
                                            

            '>>>>>>>>>>>>>>>>>>>    ﾃﾞﾌｫﾙﾄ事業部    2013.04.02
            If JIGYOBU_DEF_VALUE <> "" Then
                If Trim(StrConv(ITEMREC.L_JGYOBU_CODE, vbUnicode)) = "" Then
                    Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, JIGYOBU_DEF_VALUE)
                End If
            End If
            '>>>>>>>>>>>>>>>>>>>    ﾃﾞﾌｫﾙﾄ事業部    2013.04.02
                                            
                                            '事業部
            If Trim(StrConv(ITEMREC.L_JGYOBU_CODE, vbUnicode)) = "" Then
                Combo1(pcmbL_JGYOBU).ListIndex = -1
            Else
                For i = 0 To Combo1(pcmbL_JGYOBU).ListCount - 1
                    L_CODE = Left(Right(Combo1(pcmbL_JGYOBU).List(i), 4), 2)
                    If Trim(L_CODE) = "" Then
                        L_CODE = Right(Combo1(pcmbL_JGYOBU).List(i), 2)
                    End If
                    
                    
                    If StrConv(ITEMREC.L_JGYOBU_CODE, vbUnicode) = L_CODE Then
                        Combo1(pcmbL_JGYOBU).ListIndex = i
                        Exit For
                        
                    End If
                
                
                Next i
            End If
                                            '備考
            Text1(ptxL_BIKOU).Text = Trim(StrConv(ITEMREC.L_BIKOU, vbUnicode))
                                            '備考3
            Text1(ptxL_BIKOU3).Text = Trim(StrConv(ITEMREC.L_BIKOU3, vbUnicode))
                                            '入り数
            If Not IsNumeric(Trim(StrConv(ITEMREC.L_IRI_QTY, vbUnicode))) Then
                Text1(ptxL_IRI_QTY).Text = ""
            Else
                Text1(ptxL_IRI_QTY).Text = CLng(StrConv(ITEMREC.L_IRI_QTY, vbUnicode))
            End If
                                            '棚番(1)
            'Text1(ptxL_TANA1).Text = Trim(StrConv(ITEMREC.L_TANA1, vbUnicode))         '2013.04.25
            Text1(ptxL_TANA1).Text = RTrim(StrConv(ITEMREC.L_TANA1, vbUnicode))         '2013.04.25
                                            '棚番(2)
            'Text1(ptxL_TANA2).Text = Trim(StrConv(ITEMREC.L_TANA2, vbUnicode))         '2013.04.25
            Text1(ptxL_TANA2).Text = RTrim(StrConv(ITEMREC.L_TANA2, vbUnicode))         '2013.04.25
                                            
                                            
                                            '紙
'            If StrConv(ITEMREC.L_PAPER, vbUnicode) = L_PAPER_OFF Then
'                Check1(pchkL_PAPER).Value = vbUnchecked
'            Else
'                Check1(pchkL_PAPER).Value = vbChecked
'            End If
                                            
                                            
            If StrConv(ITEMREC.L_PAPER, vbUnicode) = L_PAPER_ON Then
                Check1(pchkL_PAPER).Value = vbChecked
            Else
                Check1(pchkL_PAPER).Value = vbUnchecked
            End If
                                            
                                            'プラ
'            If StrConv(ITEMREC.L_PLASTIC, vbUnicode) = L_PLASTIC_OFF Or StrConv(ITEMREC.L_PLASTIC, vbUnicode) <= " " Then
'                Check1(pchkL_PLASTIC).Value = vbUnchecked
'            Else
'                Check1(pchkL_PLASTIC).Value = vbChecked
'            End If
                                            
                                            
            If StrConv(ITEMREC.L_PLASTIC, vbUnicode) = L_PLASTIC_ON Then
                Check1(pchkL_PLASTIC).Value = vbChecked
            Else
                Check1(pchkL_PLASTIC).Value = vbUnchecked
            End If
                                            
                                            
                                            '適用機種ラベル
'            If StrConv(ITEMREC.L_LABEL, vbUnicode) = L_LABEL_OFF Or StrConv(ITEMREC.L_LABEL, vbUnicode) <= " " Then
'                Check1(pchkL_LABEL).Value = vbUnchecked
'            Else
'                Check1(pchkL_LABEL).Value = vbChecked
'            End If
                                            
                                            
            If StrConv(ITEMREC.L_LABEL, vbUnicode) = L_LABEL_ON Then
                Check1(pchkL_LABEL).Value = vbChecked
            Else
                Check1(pchkL_LABEL).Value = vbUnchecked

            End If
                                            
                                            '枚数ラベル
'            If StrConv(ITEMREC.L_MAISU, vbUnicode) = L_MAISU_OFF Or StrConv(ITEMREC.L_MAISU, vbUnicode) <= " " Then
'                Check1(pchkL_MAISU).Value = vbUnchecked
'            Else
'                Check1(pchkL_MAISU).Value = vbChecked
'            End If
                                            
            If StrConv(ITEMREC.L_MAISU, vbUnicode) = L_MAISU_ON Then
                Check1(pchkL_MAISU).Value = vbChecked
            Else
                Check1(pchkL_MAISU).Value = vbUnchecked
            End If
                                            
                                            
                                            '価格(1)
            If Not IsNumeric(Trim(StrConv(ITEMREC.L_URIKIN1, vbUnicode))) Then
                Text1(ptxL_URIKIN1).Text = ""
            Else
                Text1(ptxL_URIKIN1).Text = Format(CDbl(StrConv(ITEMREC.L_URIKIN1, vbUnicode)), "#0")
            End If
                                            '価格(2)
            If Not IsNumeric(Trim(StrConv(ITEMREC.L_URIKIN2, vbUnicode))) Then
                Text1(ptxL_URIKIN2).Text = ""
            Else
                Text1(ptxL_URIKIN2).Text = Format(CDbl(StrConv(ITEMREC.L_URIKIN2, vbUnicode)), "#0")
            End If
                                            '価格(3)
            If Not IsNumeric(Trim(StrConv(ITEMREC.L_URIKIN3, vbUnicode))) Then
                Text1(ptxL_URIKIN3).Text = ""
            Else
                Text1(ptxL_URIKIN3).Text = Format(CDbl(StrConv(ITEMREC.L_URIKIN3, vbUnicode)), "#0")
            End If



            '原産国     2008.06.12
            Text1(ptxGENSANKOKU).Text = Trim(StrConv(ITEMREC.GENSANKOKU, vbUnicode))

            
            
            
            '引数追加   2012.01.23
            If GENSANKOKU_SET_Proc(Trim(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode))) Then
                Exit Function
            End If

'2010.11.11 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                                                        '原産国印字有無
            If StrConv(ITEMREC.PRT_GENSANKOKU, vbUnicode) = "1" Then
                Check1(pchkGENSANKOKU).Value = vbChecked
            ElseIf StrConv(ITEMREC.PRT_GENSANKOKU, vbUnicode) = "0" Then
                Check1(pchkGENSANKOKU).Value = vbUnchecked
            Else
                    '品目ﾏｽﾀ項目が未設定⇒ini定義値により初期表示
                If GENSANKOKU_FLG = "1" Then
                    Check1(pchkGENSANKOKU).Value = vbChecked
                Else
                    Check1(pchkGENSANKOKU).Value = vbUnchecked
                End If
            End If
'2010.11.11 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<




'>>>>>>>>>>>>>>>>>>>>   2012.01.13
                                            '作業指示
'            RichTextBox1(prchL_SAGYO_SHIJI).Text = IIf(Len(RTrim(StrConv(ITEMREC.L_SAGYO_SHIJI, vbUnicode))) = 450, "", Trim(StrConv(ITEMREC.L_SAGYO_SHIJI, vbUnicode)))
            
            On Error Resume Next
            
            L_SAGYO_SHIJI = RTrim(StrConv(ITEMREC.L_SAGYO_SHIJI, vbUnicode))
            
            RichTextBox1(prchL_SAGYO_SHIJI).Text = L_SAGYO_SHIJI
            If Err.Number = 7 Then
                RichTextBox1(prchL_SAGYO_SHIJI).Text = ""
            End If
            On Error GoTo 0
'>>>>>>>>>>>>>>>>>>>>   2012.01.13
                                            
                                            
                                            
                                            
                                            '機種(1)
            Text1(ptxL_KISHU1).Text = Trim(StrConv(ITEMREC.L_KISHU1, vbUnicode))
                                            '機種(2)
'            Text1(ptxL_KISHU2).Text = Trim(StrConv(ITEMREC.L_KISHU2, vbUnicode))
            ' 2006.02.06 KUBOTA IIFでメモリ不足エラーを回避
'            RichTextBox1(prchL_KISHU2).Text = IIf(Len(RTrim(StrConv(ITEMREC.L_KISHU2, vbUnicode))) = 52, "", RTrim(StrConv(ITEMREC.L_KISHU2, vbUnicode)))
            
            L_KISHU2 = StrConv(ITEMREC.L_KISHU2, vbUnicode)
'            RichTextBox1(prchL_KISHU2).Text = IIf(Len(RTrim(StrConv(ITEMREC.L_KISHU2, vbUnicode))) = 52, "", RTrim(StrConv(ITEMREC.L_KISHU2, vbUnicode)))
'>>>>>>>>>>>>>>>>>>>>   2011.12.15
            On Error Resume Next
            RichTextBox1(prchL_KISHU2).Text = L_KISHU2
            If Err.Number = 7 Then
                RichTextBox1(prchL_KISHU2).Text = ""
            End If
            On Error GoTo 0
'>>>>>>>>>>>>>>>>>>>>   2011.12.15
                                            
                                            '機種(3)
'            RichTextBox1(prchL_KISHU3).Text = Trim(StrConv(ITEMREC.L_KISHU3, vbUnicode))
'>>>>>>>>>>>>>>>>>>>>   2012.01.13
'            RichTextBox1(prchL_KISHU3).Text = IIf(Len(RTrim(StrConv(ITEMREC.L_KISHU_BIKOU, vbUnicode))) = 450, "", Trim(StrConv(ITEMREC.L_KISHU_BIKOU, vbUnicode)))
'            RichTextBox1(prchL_KISHU_BIKOU).Text = IIf(Len(RTrim(StrConv(ITEMREC.L_KISHU3, vbUnicode))) = 150, "", RTrim(StrConv(ITEMREC.L_KISHU3, vbUnicode)))
            On Error Resume Next
            
            L_KISHU_BIKOU = RTrim(StrConv(ITEMREC.L_KISHU_BIKOU, vbUnicode))
            
            RichTextBox1(prchL_KISHU3).Text = L_KISHU_BIKOU
            If Err.Number = 7 Then
                RichTextBox1(prchL_KISHU3).Text = ""
            End If
            On Error GoTo 0
                                            '適用機種備考
'            RichTextBox1(prchL_KISHU_BIKOU).Text = Trim(StrConv(ITEMREC.L_KISHU_BIKOU, vbUnicode))
'            RichTextBox1(prchL_KISHU_BIKOU).Text = IIf(Len(RTrim(StrConv(ITEMREC.L_KISHU3, vbUnicode))) = 150, "", RTrim(StrConv(ITEMREC.L_KISHU3, vbUnicode)))
            On Error Resume Next
            
            L_KISHU3 = RTrim(StrConv(ITEMREC.L_KISHU3, vbUnicode))
            RichTextBox1(prchL_KISHU_BIKOU).Text = L_KISHU3
            If Err.Number = 7 Then
                RichTextBox1(prchL_KISHU_BIKOU).Text = ""
            End If
            On Error GoTo 0
'>>>>>>>>>>>>>>>>>>>>   2012.01.13
            
            
            
            
            '印刷日付
            Text1(ptxL_PRI_DATE).Text = Format(Now, "YYYY/mm/DD")
        
        
        
            lblUpd_Tanto.Caption = StrConv(ITEMREC.UPD_TANTO, vbUnicode)
            lblUpd_DateTime.Caption = StrConv(ITEMREC.UPD_DATETIME, vbUnicode)
        
        
        Case BtErrKeyNotFound
        
            MsgBox "他端末で変更されています。前画面に戻ります。"
            PM000402.Visible = False
            INIT_FLG = False
            
            Exit Function
                    
        Case Else
            
            
            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.19 ﾃｽﾄﾛｼﾞｯｸ
            If sts > 3000 Or sts = 3 Then
            
                
'                MsgBox "てすと"
                
                
                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                If sts Then
                    Call File_Error(sts, BtOpReset, "")
                End If
            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.23
'                                            '品目マスタＯＰＥＮ
'                If ITEM_Open(BtOpenNomal) Then
'                    Unload Me
'                End If
'
'                                            '品目マスタＯＰＥＮ
'                If L_ITEM_Open(BtOpenNomal) Then
'                    Unload Me
'                End If
'
'                                            '原産国マスタＯＰＥＮ
'                If GENSAN_Open(BtOpenNomal) Then
'                    Beep
'                    MsgBox "システム異常が発生しました。処理を中止して下さい。"
'                    Unload Me
'                End If
'                                            'コードマスタＯＰＥＮ
'                If P_CODE_Open(BtOpenNomal) Then
'                    Unload Me
'                End If
'                                            '受払先マスタ（仕入先）ＯＰＥＮ
'                If P_UKEHARAI_Open(BtOpenNomal) Then
'                    Unload Me
'                End If
'
'                Call P_CODE_TBL_Proc
'
'
'                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.06.01
'                                            'PNマスタＯＰＥＮ
'                If PN_M_Open(0) Then
'                    Beep
'                    MsgBox "システム異常が発生しました。処理を中止して下さい。"
'                    Unload Me
'               End If
'                '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.23
            
            
                '2012.02.13
                Call File_Open_Proc

            
                GoTo Start_PRoc
            
            End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.01.19 ﾃｽﾄﾛｼﾞｯｸ
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
             
            
            Exit Function
        
    End Select

    Item_Disp_Proc = False

End Function

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   品目マスタ出力
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim com     As Integer
Dim i       As Integer

Dim L_CODE  As String

    Update_Proc = True
    
    '品目マスタ　読み込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                Exit Function
        
        End Select
    
    
    Loop
    '--------------------------------------------------------レコード内容編集
    
    If com = BtOpInsert Then
        Call UniCode_Conv(ITEMREC.JGYOBU, Last_JGYOBU)                              '事業部
        Call UniCode_Conv(ITEMREC.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))        '国内外
        Call UniCode_Conv(ITEMREC.HIN_GAI, Text1(ptxHIN_GAI).Text)                  '品目ｺｰﾄﾞ
        
        Call UniCode_Conv(ITEMREC.ST_SET_DT, "")                                    '標準棚番設定日付
        Call UniCode_Conv(ITEMREC.ST_SOKO, "")                                      '標準入庫　倉庫
        Call UniCode_Conv(ITEMREC.ST_RETU, "")                                      '標準入庫　列
        Call UniCode_Conv(ITEMREC.ST_REN, "")                                       '標準入庫　連
        Call UniCode_Conv(ITEMREC.ST_DAN, "")                                       '標準入庫　段
        Call UniCode_Conv(ITEMREC.BEF_SOKO, "")                                     '前回入庫　倉庫
        Call UniCode_Conv(ITEMREC.BEF_RETU, "")                                     '前回入庫　列
        Call UniCode_Conv(ITEMREC.BEF_REN, "")                                      '前回入庫　連
        Call UniCode_Conv(ITEMREC.BEF_DAN, "")                                      '前回入庫　段
        Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")                                  '最終入庫日
        Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")                                  '最終出庫日
        Call UniCode_Conv(ITEMREC.HIN_NAI, "")                                      '品番（内）
        Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")                                   'ﾎｽﾄ倉庫
        Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")                                   'ﾎｽﾄ棚番
        Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")                              '補充点
        Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")                            '月平均出荷数
        Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                                  'ｻﾝﾌﾟﾙ数
        Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                                  'ｻﾝﾌﾟﾙ数
        Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")                                  '最終入荷日付
        Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")                                  '最終照合日付
        Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "00000000")                         '照合時在庫数
        Call UniCode_Conv(ITEMREC.BIKOU, "")                                        '印刷備考
        Call UniCode_Conv(ITEMREC.IRI_QTY, "")                                      '印刷入り数
        Call UniCode_Conv(ITEMREC.JAN_CODE, "")                                     'JANｺｰﾄﾞ
        Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")                                   '品番読み替えｺｰﾄﾞ
        Call UniCode_Conv(ITEMREC.GOODS_KBN, "1")                                   '商品化有無
        Call UniCode_Conv(ITEMREC.PACKING_NO, "")                                   '個装箱��
        Call UniCode_Conv(ITEMREC.RANK, "")                                         '現在ﾗﾝｸ
        Call UniCode_Conv(ITEMREC.NEW_RANK, "")                                     '新ﾗﾝｸ
        Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")                                  'ｸﾞﾘｯｸｽ棚番1
        Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")                                  'ｸﾞﾘｯｸｽ棚番2
        Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")                                  'ｸﾞﾘｯｸｽ棚番3
    
        Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")                                 '業務管理　 仕入区分
        Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, "")                                 '           販売区分
        Call UniCode_Conv(ITEMREC.G_SYUSHI, "")                                     '           収支単位
        Call UniCode_Conv(ITEMREC.G_KUMITATE, "")                                   '           組立製品
        Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")                                  '           標準粗利売価単価　9(8)V99
        Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")                               '           標準粗利売価設定日
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")                                  '           標準粗利原価単価  9(8)V99
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")                               '           標準粗利原価設定日
        
        For i = 0 To 2                                                              '仕入先情報
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).CODE, "")                     '           仕入先コード
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA, "")                    '           単価
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, "")                 '           単価設定日
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LOT, "")                      '           単価設定日
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LEAD_TIME, "")                '           リードタイム
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, "")            '           最終発注日
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, "")           '           最終発注数
        
        Next i
    
        Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")                              '           前月在庫金額
        Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")                                 '           資材区分
        Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_ON)                        '           ﾗﾍﾞﾙ貼り付け
        Call UniCode_Conv(ITEMREC.S_TANTO, "")                                      '収単／担当者
        
        Call UniCode_Conv(ITEMREC.FILLER, "")                                       'Filler
    
    End If
    
    Call UniCode_Conv(ITEMREC.HIN_NAME, Text1(ptxHIN_NAME).Text)                    '品名
        
    Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, Text1(ptxL_HIN_NAME_E).Text)            '品名E
                                                                                        
                                                                                    '会社名
'    Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, Left(Right(Combo1(pcmbL_KAISHA).Text, 4), 2))
                                                                                    '事業部名
'    Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, Left(Right(Combo1(pcmbL_JGYOBU).Text, 4), 2))
    
    
     L_CODE = Left(Right(Combo1(pcmbL_KAISHA).Text, 4), 2)
     If Trim(L_CODE) = "" Then
         L_CODE = Right(Combo1(pcmbL_KAISHA).Text, 2)
     End If
     Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, L_CODE)
    
     L_CODE = Left(Right(Combo1(pcmbL_JGYOBU).Text, 4), 2)
     If Trim(L_CODE) = "" Then
         L_CODE = Right(Combo1(pcmbL_JGYOBU).Text, 2)
     End If
     Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, L_CODE)
    
    
    
    
    Call UniCode_Conv(ITEMREC.L_BIKOU, Text1(ptxL_BIKOU).Text)                      '備考
    Call UniCode_Conv(ITEMREC.L_BIKOU3, Text1(ptxL_BIKOU3).Text)                    '備考3
    
    If Trim(Text1(ptxL_IRI_QTY).Text) = "" Then                                     '入り数
        Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")
    Else
        Call UniCode_Conv(ITEMREC.L_IRI_QTY, Format(CLng((Text1(ptxL_IRI_QTY).Text)), "00000000"))
    End If
    
    Call UniCode_Conv(ITEMREC.L_TANA1, Text1(ptxL_TANA1).Text)                      '棚番(1)
    Call UniCode_Conv(ITEMREC.L_TANA2, Text1(ptxL_TANA2).Text)                      '棚番(2)
    
    If Check1(pchkL_PAPER).Value = vbChecked Then                                   '紙
        Call UniCode_Conv(ITEMREC.L_PAPER, L_PAPER_ON)
    Else
        Call UniCode_Conv(ITEMREC.L_PAPER, L_PAPER_OFF)
    End If
    
    If Check1(pchkL_PLASTIC).Value = vbChecked Then                                 'プラスチック
        Call UniCode_Conv(ITEMREC.L_PLASTIC, L_PLASTIC_ON)
    Else
        Call UniCode_Conv(ITEMREC.L_PLASTIC, L_PLASTIC_OFF)
    End If
    
    If Check1(pchkL_LABEL).Value = vbChecked Then                                   '適用機種ラベル
        Call UniCode_Conv(ITEMREC.L_LABEL, L_LABEL_ON)
    Else
        Call UniCode_Conv(ITEMREC.L_LABEL, L_LABEL_OFF)
    End If
    
    If Check1(pchkL_MAISU).Value = vbChecked Then                                   '枚数ラベル
        Call UniCode_Conv(ITEMREC.L_MAISU, L_MAISU_ON)
    Else
        Call UniCode_Conv(ITEMREC.L_MAISU, L_MAISU_OFF)
    End If
    
    If Trim(Text1(ptxL_URIKIN1).Text) = "" Then                                     '価格(1)
        Call UniCode_Conv(ITEMREC.L_URIKIN1, "")
    Else
        Call UniCode_Conv(ITEMREC.L_URIKIN1, Format(CDbl((Text1(ptxL_URIKIN1).Text)), "0000000000"))
    End If
    
    If Trim(Text1(ptxL_URIKIN2).Text) = "" Then                                     '価格(2)
        Call UniCode_Conv(ITEMREC.L_URIKIN2, "")
    Else
        Call UniCode_Conv(ITEMREC.L_URIKIN2, Format(CDbl((Text1(ptxL_URIKIN2).Text)), "0000000000"))
    End If
    
    If Trim(Text1(ptxL_URIKIN3).Text) = "" Then                                     '価格(3)
        Call UniCode_Conv(ITEMREC.L_URIKIN3, "")
    Else
        Call UniCode_Conv(ITEMREC.L_URIKIN3, Format(CDbl((Text1(ptxL_URIKIN3).Text)), "0000000000"))
    End If
    
    '原産国 2008.06.12
    Call UniCode_Conv(ITEMREC.GENSANKOKU, Text1(ptxGENSANKOKU).Text)


    '原産国印字有無 2010.11.11
    Call UniCode_Conv(ITEMREC.PRT_GENSANKOKU, CStr(Check1(pchkGENSANKOKU).Value))


    Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, RichTextBox1(prchL_SAGYO_SHIJI).Text)         '作業指示
    Call UniCode_Conv(ITEMREC.L_KISHU1, Text1(ptxL_KISHU1).Text)                    '機種(1)
    Call UniCode_Conv(ITEMREC.xL_KISHU2, "")                                        '旧機種(2)
    Call UniCode_Conv(ITEMREC.L_KISHU2, RichTextBox1(prchL_KISHU2).Text)            '機種(2)
    Call UniCode_Conv(ITEMREC.L_KISHU3, RichTextBox1(prchL_KISHU_BIKOU).Text)       '機種(3)
    Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, RichTextBox1(prchL_KISHU3).Text)       '適用機種
    
    Call UniCode_Conv(ITEMREC.UPD_TANTO, "PM004")                                   '更新担当者ｺｰﾄﾞ 2012.01.13
                                                                                    '更新日時
    Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    Do
        
        DoEvents
        
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Update_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Exit Function
        End Select
    
    Loop
    
    
    Update_Proc = False


End Function

Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   品目マスタ削除
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer

    Delete_Proc = True
    
    '品目マスタ　読み込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)

    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                Delete_Proc = False
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Delete_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                Exit Function
        
        End Select
    
    
    Loop
    
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpDelete, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Delete_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, BtOpDelete, "品目マスタ")
                Exit Function
        End Select
    Loop


    Delete_Proc = False


End Function


Private Sub Check1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    
Dim i   As Integer
    
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
    
    
    Select Case Index
    
        Case pcmbJGYOBU
    
            
            For i = 0 To UBound(JGYOBU_T)
                If Right(Combo1(pcmbJGYOBU).Text, 1) = JGYOBU_T(i).CODE Then
                
                    
                    Last_JGYOBU = JGYOBU_T(i).CODE
                    Exit For
                
                End If
            Next i
    
    
    End Select
    
    
    Call Tab_Ctrl(Shift)        '移動

End Sub


Private Sub Combo1_LostFocus(Index As Integer)
Dim i   As Integer
    
    
    
    Select Case Index
    
        Case pcmbJGYOBU
    
            For i = 0 To UBound(JGYOBU_T)
                If Right(Combo1(pcmbJGYOBU).Text, 1) = JGYOBU_T(i).CODE Then
                
                    
                    Last_JGYOBU = JGYOBU_T(i).CODE
                    Exit For
                
                End If
            Next i
    
    
    
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans         As Integer
Dim i           As Integer

Dim Parts       As String   '品番
Dim ID          As Long     '指示��

Dim PartsLabel  As Integer  '品番ﾗﾍﾞﾙ 0:なし 以外：枚数
Dim KisyuLabel  As Integer  '機種ﾗﾍﾞﾙ 0:なし
Dim JanLabel    As Integer  'JANﾗﾍﾞﾙ 0:なし
Dim GLabel      As Integer  '外装ﾗﾍﾞﾙ 0:なし
Dim ItemLabel   As Integer  'ｱｲﾃﾑﾗﾍﾞﾙ枚数

Dim OrderNo     As String
Dim ItemNo      As String

Dim Pri_Date    As String

Dim objAccess       As Access.Application
Dim strAccessPath   As String

Dim com         As Integer
Dim sts         As Integer

Dim L_Print_Flg As Boolean
    
Dim check_flg   As Boolean
    
    
Dim check_flg1  As Boolean      '2008.09.26
    
    
Dim L_CODE      As String
    
Dim FileNo      As Long         '2008.05.30
    
    
Dim L_QTY       As Long         '2008.10.03
    
Dim yn          As Integer      '2012.01.13
    
Dim KISHU1      As String       '2012.09.03
Dim KISHU2      As String       '2012.09.03

    
    

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  エラーチェック 2016.01.29
Dim GEN_NG_F        As Integer      '原産国空白
Dim GEN_AT_GAI_F    As Integer      '原産国注意(海外)
Dim GEN_AT_PLU_F    As Integer      '原産国注意(複数)
Dim TANKA_SP_F      As Integer      '単価空白(単価２,単価３)
Dim KISHU_NG_F      As Integer      '機種空白
Dim KAISYA_NG_F     As Integer      '会社／事業部空白

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  エラーチェック 2016.01.29
                    
    
    
    
    
    
    
    
    Select Case Index
        Case P_CMD_Upd                      '更新
            
            
            For i = ptxHIN_GAI To ptxL_KISHU1
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
            
            
            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.01.13
                
            
'           Len-->LenB  2012.01.23
            If LenB(RTrim(RichTextBox1(prchL_SAGYO_SHIJI).Text)) > 450 Then
                
                '"作業指示が桁数オーバーしています(最大450文字) 継続しますか？"-->"作業指示が桁数オーバーしています(最大450文字) オーバーした文字は切り捨てられます。"    2012.01.23
                yn = MsgBox("作業指示が桁数オーバーしています(最大450文字)、オーバーした文字は切り捨てられます。", vbYesNo, "確認入力")
                If yn = vbNo Then
                    RichTextBox1(prchL_SAGYO_SHIJI).SetFocus
                    Exit Sub
                End If
            End If
        
        
'           Len-->LenB  2012.01.23
            If LenB(RTrim(RichTextBox1(prchL_KISHU2).Text)) > 52 Then
                '"機種(2)が桁数オーバーしています(最大52文字) 継続しますか？"-->"機種(2)が桁数オーバーしています(最大52文字)オーバした文字は切り捨てられます。"    2012.01.23
                yn = MsgBox("機種(2)が桁数オーバーしています(最大52文字)、オーバした文字は切り捨てられます。", vbYesNo, "確認入力")
                If yn = vbNo Then
                    Exit Sub
                End If
                RichTextBox1(prchL_KISHU2).SetFocus
            End If
        
        
'           Len-->LenB  2012.01.23
            If LenB(RTrim(RichTextBox1(prchL_KISHU3).Text)) > 450 Then
                '"機種(3)が桁数オーバーしています(最大450文字)継続しますか？"-->"機種(3)が桁数オーバーしています(最大450文字)、オーバした文字は切り捨てられます。"    2012.01.23
                yn = MsgBox("機種(3)が桁数オーバーしています(最大450文字)、オーバした文字は切り捨てられます。", vbYesNo, "確認入力")
                If yn = vbNo Then
                    RichTextBox1(prchL_KISHU3).SetFocus
                    Exit Sub
                End If
            End If
        
        
    
'           Len-->LenB  2012.01.23
'            If LenB(RTrim(RichTextBox1(prchL_KISHU_BIKOU).Text)) > 150 Then
'           strconv追加 2015.12.28
            If LenB(StrConv(RTrim(RichTextBox1(prchL_KISHU_BIKOU).Text), vbFromUnicode)) > 150 Then
                '"適用機種備考が桁数オーバーしています(最大150文字)継続しますか？"-->"適用機種備考が桁数オーバーしています(最大150文字)、オーバした文字は切り捨てられます。"    2012.01.23
                yn = MsgBox("適用機種備考が桁数オーバーしています(最大150文字)、オーバした文字は切り捨てられます。", vbYesNo, "確認入力")
                If yn = vbNo Then
                    RichTextBox1(prchL_KISHU_BIKOU).SetFocus
                    Exit Sub
                End If
            End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.01.13
            
            
            
            
            
            
            
            Beep
            ans = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            Else
                Exit Sub
            End If
'            PM000402.Visible = False
'            INIT_FLG = False
                    
            Call Clear_Proc
        
        Case P_CMD_DEL                      '削除
            ans = MsgBox("削除しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Delete_Proc() Then
                    Unload Me
                End If
            Else
                Exit Sub
            End If
'            PM000402.Visible = False
'            INIT_FLG = False
 
            Call Clear_Proc
 
 
 
 '       Case P_CMD_DSP                      '検索/表示
 '       Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
 '       Case P_CMD_PRT                      '印刷
        
        Case pcmdLabel, pcmdItem, pcmdJan, pcmdGAISO
            
            If Error_Check_Proc(ptxHIN_GAI) Then        '2014.07.18
                Exit Sub                                '2014.07.18
            End If                                      '2014.07.18

            
            If Not IsNumeric(Text1(ptxL_MAISU).Text) Then
        
                MsgBox "入力した項目はエラーです。"
                Text1(ptxL_MAISU).SetFocus
                Exit Sub
        
            Else
                If CInt(Text1(ptxL_MAISU).Text) <= 0 Then
                
                    MsgBox "入力した項目はエラーです。"
                    Text1(ptxL_MAISU).SetFocus
                    Exit Sub
                
                End If
            
            End If
            
            If Trim(Text1(ptxL_PRI_DATE).Text) = "" Then
            Else
                If Not IsDate(Text1(ptxL_PRI_DATE).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(ptxL_MAISU).SetFocus
                    Exit Sub
                End If
            End If
        
            L_Print_Flg = True
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  エラーチェック 2016.01.29
            GEN_NG_F = 0        '原産国空白
            GEN_AT_GAI_F = 0    '原産国注意(海外)
            GEN_AT_PLU_F = 0    '原産国注意(複数)
            TANKA_SP_F = 0      '単価空白(単価２,単価３)
            KISHU_NG_F = 0      '機種空白
            KAISYA_NG_F = 0     '会社／事業部空白

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>原産国空白チェック>>>>>>>>>>>>>>>>>>>>>>>>>>
            For i = 0 To UBound(GENSANKOKU_CHECK_TBL)
                If Last_JGYOBU = GENSANKOKU_CHECK_TBL(i) Then
                    Exit For
                End If
            Next i
            If i > UBound(GENSANKOKU_CHECK_TBL) Then
                GEN_NG_F = 9
            Else
                If Check1(pchkGENSANKOKU).Value = vbUnchecked Then
                    GEN_NG_F = 9
                Else
                    '原産国、空白か？
                    If Text1(ptxGENSANKOKU).Enabled Then
                        If Trim(Text1(ptxGENSANKOKU).Text) = "" Then
                            GEN_NG_F = 1
                        End If
                    Else
                        If Combo1(pcmbGENSAN).Enabled Then
                            If Trim(Left(Combo1(pcmbGENSAN).Text, 20)) = "" Then
                                GEN_NG_F = 1
                            End If
                        End If
                    End If
                End If
            End If

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>海外供給区分チェック>>>>>>>>>>>>>>>>>>>>>>>>>>
            If GAI_BUHIN_CHECK Then
                If StrConv(ITEMREC.GAI_BUHIN, vbUnicode) = "1" Or _
                    StrConv(ITEMREC.GAI_BUHIN, vbUnicode) = "2" Or _
                    StrConv(ITEMREC.GAI_BUHIN, vbUnicode) = "3" Then
                    GEN_AT_GAI_F = 1
                End If
            End If
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>原産国海外向けチェック＆原産国複数チェック>>>>
            If Text1(ptxGENSANKOKU).Enabled Then
                GEN_AT_PLU_F = 1
            Else
                GEN_AT_PLU_F = Combo1(pcmbGENSAN).ListCount
            End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>単価チェック>>>>>>>>>>>>>>>>>>>>>>>>>>
            If TANKA_SPACE_F = "1" Then
                If Not IsNumeric(Text1(ptxL_URIKIN2).Text) Or _
                     Not IsNumeric(Text1(ptxL_URIKIN3).Text) Then
                    TANKA_SP_F = 1
                End If
            End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>代表機種チェック>>>>>>>>>>>>>>>>>>>>>>>>>>
            KISHU1 = ""
            KISHU2 = ""
            If KISHU_CHECK = 1 Then
                For i = 1 To Len(Trim(Text1(ptxL_KISHU1).Text))
                    If Mid(Text1(ptxL_KISHU1).Text, i, 1) <= " " Then
                    Else
                        KISHU1 = KISHU1 & Mid(Text1(ptxL_KISHU1).Text, i, 1)
                    End If
                Next i
                
                For i = 1 To Len(Trim(RichTextBox1(prchL_KISHU2).Text))
                    If Mid(RichTextBox1(prchL_KISHU2).Text, i, 1) <= " " Then
                    Else
                        KISHU2 = KISHU2 & Mid(RichTextBox1(prchL_KISHU2).Text, i, 1)
                    End If
                Next i
                
                If Trim(KISHU1) = "" And Trim(KISHU2) = "" Then
                    KISHU_NG_F = 1
                End If
            End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>会社名／事業部名チェック>>>>>>>>>>>>>>>>>>>>>>>>>>
            If KAISYA_RESTRICT_F Then
                KAISYA_NG_F = 9
            Else
                If KAISYA_CHK_F Then
                    If Trim(Combo1(pcmbL_KAISHA).Text) = "" Or _
                        Trim(Combo1(pcmbL_JGYOBU).Text) = "" Then
                        KAISYA_NG_F = 1
                    End If
                End If
            End If

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>メッセージ作成>>>>>>>>>>>>>>>>>>>>>>>>>>
            yn = Mesg_Set_Proc(GEN_NG_F, GEN_AT_GAI_F, GEN_AT_PLU_F, TANKA_SP_F, KISHU_NG_F, KAISYA_NG_F, KISHU1, KISHU2)
            If yn = vbCancel Then
                Exit Sub
            End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  エラーチェック 2016.01.29



            check_flg1 = False                              '2008.09.26
            If Trim(Combo1(pcmbL_KAISHA).Text) = "" Then    '2008.09.26
            Else                                            '2008.09.26
                check_flg1 = True                           '2008.09.26
            End If                                          '2008.09.26
'            check_flg1 = False                              '2008.09.26 ← 2013.01.04 削除
            If Trim(Combo1(pcmbL_JGYOBU).Text) = "" Then    '2008.09.26
            Else                                            '2008.09.26
                check_flg1 = True                           '2008.09.26
            End If                                          '2008.09.26
        
        
        
            '>>>>>>>>>>>>>>>>>>>    2013.04.02  小野 or条件
            If KAISYA_RESTRICT_F Then
                If Trim(Combo1(pcmbL_KAISHA).Text) = "" Or _
                    Trim(Combo1(pcmbL_JGYOBU).Text) = "" Then
                    check_flg1 = False
                End If
            End If
            '>>>>>>>>>>>>>>>>>>>    2013.04.02
        
        
        
            If Not check_flg1 Then       '2008.09.26
            
                If KAISYA_CHK_F Then
            
'                    MsgBox "会社名もしくは事業部が空白の為、印刷できません"
'                    Text1(ptxHIN_GAI).SetFocus
'
'                    Exit Sub


                    '2013.04.02 確認メッセージ変更
                    'ans = MsgBox("会社名/事業部 が指定されていません。(ＯＫ＝発行、ｷｬﾝｾﾙ=発行しない)", vbOKCancel + vbQuestion + vbDefaultButton2, "確認入力")
                    'ans = MsgBox("会社名/事業部 が指定されていません。(ＯＫ＝発行、ｷｬﾝｾﾙ=発行しない)" & vbCrLf & vbCrLf & "会 社　名：" & Left(Trim(Combo1(pcmbL_KAISHA).Text), 30) & vbCrLf & "事業部名：" & Left(Trim(Combo1(pcmbL_JGYOBU).Text), 30), vbOKCancel + vbQuestion + vbDefaultButton2, "確認入力")   '2013.09.03 DELETE
'2016.01.29                    ans = MsgBox("会社/事業部が指定されていません。ラベル印刷しますか？", vbYesNo + vbQuestion + vbDefaultButton2, "確認入力")                  '2013.09.03
                    '2013.04.02 確認メッセージ変更
                    'If ans = vbCancel Then         '2013.09.03
ans = vbYes     '2016.01.29
                    If ans = vbNo Then              '2013.09.03
                        Text1(ptxHIN_GAI).SetFocus
                        Exit Sub
                    End If






                End If
            
            End If
        
        
        
            If KAISYA_CHK_F Then        '2008.09.26
            
            
            
                If Not IsNumeric(Text1(ptxL_URIKIN2).Text) Or _
                     Not IsNumeric(Text1(ptxL_URIKIN3).Text) Then
                
            '↓2010.02.08
                    
                    
                    
                    If TANKA_SPACE_F = "1" Then
                    
                        'ans = MsgBox("単価未登録です。(ＯＫ＝強制発行、ｷｬﾝｾﾙ=発行しない)", vbOKCancel + vbQuestion + vbDefaultButton2, "確認入力")     '2013.09.03
'2016.01.29                        ans = MsgBox("単価未設定です。ラベル印刷しますか？", vbYesNo + vbQuestion, "確認入力")                                          '2013.09.03
                        'If ans = vbCancel Then         '2013.09.03
ans = vbYes '2016.01.29
                        If ans = vbNo Then              '2013.09.03
                            Text1(ptxHIN_GAI).SetFocus
                            Exit Sub
                        End If
                    Else

'2016.01.29                        MsgBox "単価未登録の為、発行できません"
'2016.01.29                        Text1(ptxHIN_GAI).SetFocus
'2016.01.29                        Exit Sub
            
                    End If
            '↑2010.02.08
                
                End If
            
            
            
            
                check_flg = True
            
            
            Else
                check_flg = False
                If Not IsNumeric(Text1(ptxL_URIKIN1).Text) Then
                Else
                    If CDbl(Text1(ptxL_URIKIN1).Text) <> 0 Then
                        check_flg = True
                    End If
                End If
                
                If Not IsNumeric(Text1(ptxL_URIKIN2).Text) Then
                Else
                    If CDbl(Text1(ptxL_URIKIN2).Text) <> 0 Then
                        check_flg = True
                    End If
                End If
                If Not IsNumeric(Text1(ptxL_URIKIN3).Text) Then
                Else
                    If CDbl(Text1(ptxL_URIKIN3).Text) <> 0 Then
                        check_flg = True
                    End If
                End If
            End If
            
            
            If PRINT_CHECK_F Then       '2008.09.26
            
            
                '↓2008.05.30
                Do
                    On Error Resume Next
    
                    FileNo = FreeFile
    
                    Open LabelPrint_F For Input As FileNo
    
                    Select Case Err.Number
                        Case 0
    
    
                            Close #FileNo
    
                            ans = MsgBox("ラベル発行中です", vbAbortRetryIgnore + vbQuestion, "確認入力")
    
                            Select Case ans
                            
                                Case vbAbort    '中止
    
                                    Exit Sub
                            
                                Case vbIgnore   '無視
                            
                                    Exit Do
                            
                            End Select
    
    
    
    
                        Case 53
                            Exit Do
    
    
                        Case Else
    
                            Unload Me
    
    
                    End Select
    
                    On Error GoTo 0
    
                Loop
                
'''''''''''''''2011.04.17
''                Open LabelPrint_F For Output As FileNo
''                Close #FileNo
'''''''''''''''2011.04.17
            
            End If
            '↑2008.05.30
            
            
            
            
            
            If Not check_flg Then
'2016.01.29                ans = MsgBox("単価未設定です。ラベル印刷しますか？", vbYesNo + vbQuestion, "確認入力")
ans = vbYes '2016.01.29
                If ans = vbYes Then
                Else
                    L_Print_Flg = False
                End If
            End If
            
            '2009.03.28
            For i = 0 To UBound(GENSANKOKU_CHECK_TBL)
            
            
                If Last_JGYOBU = GENSANKOKU_CHECK_TBL(i) Then
                    Exit For
                End If
            
            Next i
            '2009.03.28
            If i > UBound(GENSANKOKU_CHECK_TBL) Then
            Else
                
                If Trim(Text1(ptxGENSANKOKU).Text) = "" And Trim(Combo1(pcmbGENSAN).Text) = "" Then
                    

                    'ans = MsgBox("原産国が空白です。(ＯＫ＝印刷中止、ｷｬﾝｾﾙ=継続)", vbOKCancel + vbQuestion, "確認入力")                '2013.09.03
'2016.01.29                    ans = MsgBox("原産国が空白です。ラベル印刷しますか？", vbYesNo + vbQuestion + vbDefaultButton2, "確認入力")         '2013.09.03
                    'If ans = vbCancel Then '2013.09.03
ans = vbYes '2016.01.29
                    If ans = vbYes Then      '2013.09.03
                    Else
                        L_Print_Flg = False
                    
                        Exit Sub            '2013.01.17
                    End If
                End If
            End If
                
                
                
            '2012.09.03     代表機種ﾁｪｯｸ
            If KISHU_CHECK Then
                KISHU1 = ""
                KISHU2 = ""
                
                'For i = 1 To Len(Trim(Text1(prchL_KISHU2).Text))
                For i = 1 To Len(Trim(Text1(ptxL_KISHU1).Text))
                    If Mid(Text1(ptxL_KISHU1).Text, i, 1) <= " " Then
                    Else
                        KISHU1 = KISHU1 & Mid(Text1(ptxL_KISHU1).Text, i, 1)
                    End If
                Next i
                
                For i = 1 To Len(Trim(RichTextBox1(prchL_KISHU2).Text))
                    If Mid(RichTextBox1(prchL_KISHU2).Text, i, 1) <= " " Then
                    Else
                        KISHU2 = KISHU2 & Mid(RichTextBox1(prchL_KISHU2).Text, i, 1)
                    End If
                Next i
                
                If Trim(KISHU1) = "" And Trim(KISHU2) = "" Then
                    'ans = MsgBox("代表機種が空白です。(ＯＫ＝印刷中止、ｷｬﾝｾﾙ=継続)", vbOKCancel + vbQuestion, "確認入力")          '2013.09.03
'2016.01.29                    ans = MsgBox("代表機種が空白です。ラベル印刷しますか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認入力")    '2013.09.03
                    'If ans = vbCancel Then     '2013.09.03
ans = vbYes '2016.01.29
                    If ans = vbYes Then          '2013.09.03
                    Else
                        L_Print_Flg = False
                        
                        Exit Sub            '2013.01.17
                    
                    End If
                End If
            End If
            
            '2012.09.03     代表機種ﾁｪｯｸ
            
                
                
                
            If L_Print_Flg Then
                                
                                
                                
 '''''''''''''''2011.04.17
                If PRINT_CHECK_F Then       '2008.09.26
                                
                                
                    Open LabelPrint_F For Output As FileNo
                    Close #FileNo
                End If
'''''''''''''''2011.04.17
                               
                                
                                
                                
'-----------------  ﾚｺｰﾄﾞの中身入れ替え
                Call UniCode_Conv(ITEMREC.HIN_NAME, Text1(ptxHIN_NAME).Text)                    '品名

                Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, Text1(ptxL_HIN_NAME_E).Text)            '品名E


                                                                                                '会社名
                L_CODE = Left(Right(Combo1(pcmbL_KAISHA).Text, 4), 2)
                If Trim(L_CODE) = "" Then
                    L_CODE = Right(Combo1(pcmbL_KAISHA).Text, 2)
                End If
                Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, L_CODE)

                L_CODE = Left(Right(Combo1(pcmbL_JGYOBU).Text, 4), 2)
                If Trim(L_CODE) = "" Then
                    L_CODE = Right(Combo1(pcmbL_JGYOBU).Text, 2)
                End If
                Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, L_CODE)
' 2013.01.05 Upd >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                '空白時の設定値 セット
                If KAISYA_DEF_VALUE <> "" Then
                    If Trim(StrConv(ITEMREC.L_KAISHA_CODE, vbUnicode)) = "" Then
                        Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, KAISYA_DEF_VALUE)
                    End If
                End If

                If JIGYOBU_DEF_VALUE <> "" Then
                    If Trim(StrConv(ITEMREC.L_JGYOBU_CODE, vbUnicode)) = "" Then
                        Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, JIGYOBU_DEF_VALUE)
                    End If
                End If


' 2013.03.29 DEL
'
'                '会社名表示制限対象なら｢会社名｣｢事業部名｣を空白にする
'                If KAISYA_RESTRICT <> "" Then
'                    If KAISYA_RESTRICT <> StrConv(ITEMREC.L_KAISHA_CODE, vbUnicode) Then
'                        Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")
'                        Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")
'                    End If
'                End If
' 2013.03.29 DEL

' 2013.04.02 INS
            If KAISYA_RESTRICT_F Then
                If Trim(Combo1(pcmbL_KAISHA).Text) = "" Or _
                    Trim(Combo1(pcmbL_JGYOBU).Text) = "" Then
                    Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")
                    Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")
                End If
            End If
' 2013.04.02 INS

' 2013.01.05 Upd <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

                Call UniCode_Conv(ITEMREC.L_BIKOU, Text1(ptxL_BIKOU).Text)                      '備考
                Call UniCode_Conv(ITEMREC.L_BIKOU3, Text1(ptxL_BIKOU3).Text)                    '備考3

                If Trim(Text1(ptxL_IRI_QTY).Text) = "" Then                                     '入り数
                    Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")
                Else
                    Call UniCode_Conv(ITEMREC.L_IRI_QTY, Format(CLng((Text1(ptxL_IRI_QTY).Text)), "00000000"))
                End If

                Call UniCode_Conv(ITEMREC.L_TANA1, Text1(ptxL_TANA1).Text)                      '棚番(1)

                '2008.10.29 棚番(1)に標準棚番をセット
                Call UniCode_Conv(ITEMREC.L_TANA1, StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_DAN, vbUnicode))

                '2008.10.29


                Call UniCode_Conv(ITEMREC.L_TANA2, Text1(ptxL_TANA2).Text)                      '棚番(2)
                
                If Check1(pchkL_PAPER).Value = vbChecked Then                                   '紙
                    Call UniCode_Conv(ITEMREC.L_PAPER, L_PAPER_ON)
                Else
                    Call UniCode_Conv(ITEMREC.L_PAPER, L_PAPER_OFF)
                End If
                
                If Check1(pchkL_PLASTIC).Value = vbChecked Then                                 'プラスチック
                    Call UniCode_Conv(ITEMREC.L_PLASTIC, L_PLASTIC_ON)
                Else
                    Call UniCode_Conv(ITEMREC.L_PLASTIC, L_PLASTIC_OFF)
                End If
                
                If Check1(pchkL_LABEL).Value = vbChecked Then                                   '適用機種ラベル
                    Call UniCode_Conv(ITEMREC.L_LABEL, L_LABEL_ON)
                Else
                    Call UniCode_Conv(ITEMREC.L_LABEL, L_LABEL_OFF)
                End If
                
                If Check1(pchkL_MAISU).Value = vbChecked Then                                   '枚数ラベル
                    Call UniCode_Conv(ITEMREC.L_MAISU, L_MAISU_ON)
                Else
                    Call UniCode_Conv(ITEMREC.L_MAISU, L_MAISU_OFF)
                End If
                
                If Trim(Text1(ptxL_URIKIN1).Text) = "" Then                                     '価格(1)
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, "")
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, Format(CDbl((Text1(ptxL_URIKIN1).Text)), "0000000000"))
                End If
                
                If Trim(Text1(ptxL_URIKIN2).Text) = "" Then                                     '価格(2)
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, "")
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, Format(CDbl((Text1(ptxL_URIKIN2).Text)), "0000000000"))
                End If
                
                If Trim(Text1(ptxL_URIKIN3).Text) = "" Then                                     '価格(3)
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, "")
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, Format(CDbl((Text1(ptxL_URIKIN3).Text)), "0000000000"))
                End If
                
                
                '原産国 2008.06.12
                If Check1(pchkGENSANKOKU).Value = vbChecked Then
                    
                    
                    If Text1(ptxGENSANKOKU).Enabled Then
                        
                        Call UniCode_Conv(ITEMREC.GENSANKOKU, Text1(ptxGENSANKOKU).Text)
                    Else
                                
                        If Combo1(pcmbGENSAN).Enabled Then
                            Call UniCode_Conv(ITEMREC.GENSANKOKU, Trim(Left(Combo1(pcmbGENSAN).Text, 20)))
                        End If
                    End If
                Else
                    Call UniCode_Conv(ITEMREC.GENSANKOKU, "")
                End If
                
                
                
                Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, RichTextBox1(prchL_SAGYO_SHIJI).Text)  '作業指示
                Call UniCode_Conv(ITEMREC.L_KISHU1, Text1(ptxL_KISHU1).Text)                    '機種(1)
                Call UniCode_Conv(ITEMREC.xL_KISHU2, "")                                        '旧機種(2)
                Call UniCode_Conv(ITEMREC.L_KISHU2, RichTextBox1(prchL_KISHU2).Text)            '機種(2)
'                Call UniCode_Conv(ITEMREC.L_KISHU3, RichTextBox1(prchL_KISHU3).Text)           '機種(3)
'                Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, RichTextBox1(prchL_KISHU_BIKOU).Text) '適用機種

                Call UniCode_Conv(ITEMREC.L_KISHU3, RichTextBox1(prchL_KISHU_BIKOU).Text)       '機種(3)
                Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, RichTextBox1(prchL_KISHU3).Text)       '適用機種


'-----------------  ﾚｺｰﾄﾞの中身入れ替え
                                
                                
                PartsLabel = 0
                KisyuLabel = 0
                JanLabel = 0
                GLabel = 0
                ItemLabel = 0

                Parts = Text1(ptxHIN_GAI).Text     '品番
    
                    
                Select Case Index
                    Case pcmdLabel
                        If Check1(pchkL_LABEL).Value = vbChecked Then
                        
                            KisyuLabel = CInt(Text1(ptxL_MAISU).Text)
                        Else
                            PartsLabel = CInt(Text1(ptxL_MAISU).Text)
                        
                        
                        End If
                    Case pcmdItem
                    
                        ItemLabel = CInt(Text1(ptxL_MAISU).Text)
                                            
                    
                    Case pcmdJan
                        JanLabel = CInt(Text1(ptxL_MAISU).Text)
                    Case pcmdGAISO
                        GLabel = CInt(Text1(ptxL_MAISU).Text)
                End Select
                OrderNo = Text1(ptxL_ORDERNO).Text
                ItemNo = Text1(ptxL_ITEMNO).Text
                Pri_Date = Text1(ptxL_PRI_DATE).Text
                
                On Error Resume Next
                Set objAccess = GetObject(, "Access.Application")
                If Err().Number <> 0 Then
                    
                    MsgBox "この端末では商品ラベル発行は行えません。"
'                        MsgBox "GetObject(Access.Application)" & Err().Number & " " & Err().Description
                Else
'                        MsgBox Err.Number
                        
                        
                    '2012.01.23
                    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                        "パーツラベル印刷　ラベル印刷開始", Me.hwnd, 0)
                        
                        
                        
                    strAccessPath = App.Path
                    If Right(strAccessPath, 1) <> "\" Then
                        strAccessPath = strAccessPath & "\"
                    End If
                    
                    strAccessPath = strAccessPath & "litem.mdb"
                    Set objAccess = GetObject(strAccessPath)

                
                
                    
                    com = BtOpGetFirst
                    Do
                    
                    
                    
                        sts = BTRV(com, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            
                                sts = BTRV(BtOpDelete, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
                                
                                Select Case sts
                                
                                    Case BtNoErr
                                    Case Else
                                        Call File_Error(sts, com, "ﾗﾍﾞﾙ用品目ﾏｽﾀ")
                                        Exit Sub
                                End Select
                            
                            Case BtErrEOF
                                Exit Do
                            Case Else
                                Call File_Error(sts, com, "ﾗﾍﾞﾙ用品目ﾏｽﾀ")
                                Exit Sub
                        End Select
                        
                        com = BtOpGetNext
                    
                    
                    Loop
                        
                    '再梱包ﾏｰｸ追加  2007.11.08
                    Call UniCode_Conv(ITEMREC.L_MARK, Text1(ptxL_MARK).Text)
                        
                        
                    sts = BTRV(BtOpInsert, L_ITEM_POS, ITEMREC, Len(ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                
                    
                
                
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Exit Sub
                        
                
                    End Select
                            
                    If IsNumeric(Text1(ptxL_QTY).Text) Then     '2008.10.03
                        L_QTY = CLng(Text1(ptxL_QTY).Text)      '2008.10.03
                    Else                                        '2008.10.03
                        L_QTY = "1"                             '2008.10.03
                    End If                                      '2008.10.03
                            
                            
                    ID = 0
'                    objAccess.Run "NewPosPrintLabel", _
'                                        Trim(Parts), _
'                                        PartsLabel, _
'                                        KisyuLabel, _
'                                        JanLabel, _
'                                        GLabel, _
'                                        ID, _
'                                        ItemLabel, _
'                                        Trim(OrderNo), _
'                                        Trim(ItemNo), _
'                                        Pri_Date

                    '2008.10.03 引数追加(L_QTY)
                    objAccess.Run "NewPosPrintLabel", _
                                        Trim(Parts), _
                                        PartsLabel, _
                                        KisyuLabel, _
                                        JanLabel, _
                                        GLabel, _
                                        ID, _
                                        ItemLabel, _
                                        Trim(OrderNo), _
                                        Trim(ItemNo), _
                                        Pri_Date, _
                                        L_QTY
                
                
                
                    '2012.01.23
                    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                        "パーツラベル印刷　ラベル印刷終了", Me.hwnd, 0)
                
                    'If RETURN_F Then                                            '2013.03.29
                    '    Me.Visible = False
                    'End If                                                      '2013.03.29
                
                    Text1(ptxL_MAISU).SetFocus      '2013.01.17
                
                End If
                
                
                
                
                
                Set objAccess = Nothing
            
            
            
            
            
            End If
            
            
            
            
            '2008.12.19
            Text1(ptxL_QTY).Text = "1"

            '2011.04.17
            Text1(ptxL_MAISU).SetFocus
        
            'PM000402.Visible = False
            'INIT_FLG = False
        
        
        
        
        
        
        
        
        Case P_CMD_End                      '終了
    
            Unload Me
    End Select

End Sub

Private Sub Form_Activate()
    
'Dim i       As Integer
'Dim CODE    As String
    
'    If INIT_FLG Then
'        Exit Sub
'    End If

'    If JGYOBU_T(i).CODE = Last_JGYOBU Then
'        PM000402.Caption = "商品化システム　品目マスタメンテナンス（商品ラベル項目）（" + RTrim(JGYOBU_T(i).NAME) + ")"
'        LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
'        LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'    End If



'    Select Case G_SCREEN_FLG
'        Case G_SCREEN_INS       '新規
'
'            Combo1(pcmbNAIGAI).BackColor = G_INPUT_OK
'            Combo1(pcmbNAIGAI).TabStop = True
'            Combo1(pcmbNAIGAI).Locked = False
'
'
'            Text1(ptxHIN_GAI).BackColor = G_INPUT_OK
'            Text1(ptxHIN_GAI).TabStop = True
'            Text1(ptxHIN_GAI).Locked = False
'
'            Text1(ptxHIN_NAME).BackColor = G_INPUT_OK
'            Text1(ptxHIN_NAME).TabStop = True
'            Text1(ptxHIN_NAME).Locked = False
'
'
'            For i = ptxHIN_GAI To ptxL_ITEMNO
'                Text1(i).Text = ""
'            Next i
'
'            For i = prchL_SAGYO_SHIJI To prchL_KISHU_BIKOU
'                RichTextBox1(i).Text = ""
'            Next i
'
'
'            For i = pcmbNAIGAI To pcmbL_JGYOBU
'
'                Combo1(i).ListIndex = -1
'            Next i
'
'
'
'
'            Combo1(pcmbNAIGAI).SetFocus
'            Combo1(pcmbNAIGAI).ListIndex = 0
'
'
'
'
'        Case G_SCREEN_UPD       '更新
'
'            Combo1(pcmbNAIGAI).BackColor = G_INPUT_NG
'            Combo1(pcmbNAIGAI).TabStop = False
'            Combo1(pcmbNAIGAI).Locked = True
'
'
'
'            Text1(ptxHIN_GAI).BackColor = G_INPUT_NG
'            Text1(ptxHIN_GAI).TabStop = False
'            Text1(ptxHIN_GAI).Locked = True
'
'            Text1(ptxHIN_NAME).BackColor = G_INPUT_OK
'            Text1(ptxHIN_NAME).TabStop = True
'            Text1(ptxHIN_NAME).Locked = False
'
'
'            CODE = PM000401.txSEL_KEY.Text
'
'            If Item_Disp_Proc(CODE) Then
'                Exit Sub
'            End If
'
'            For i = ptxL_MAISU To ptxL_ITEMNO
'                Text1(i).Text = ""
'            Next i
'
'            '========================================================= 2007/03/19 =====
'''            Text1(ptxL_HIN_NAME_E).SetFocus
'            Text1(ptxL_MAISU).SetFocus
'            '==========================================================================
'
'    End Select
'
'
'    INIT_FLG = True
'
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
            Command1(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

Dim com     As Integer
Dim sts     As Integer




Dim c       As String * 128
Dim i       As Integer


Dim wkJGYOBU    As String * 1       '2013.01.15
Dim wkNAIGAI    As String * 1       '2013.01.15


    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If

                                
                                
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "パーツラベル発行処理", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
                                
                                
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                
                                
                                
    PRINT_CHECK_F = True        '2008.09.26
                                'ラベル印刷用コントロールＦ獲得2008.05.30
    If GetIni("FILE", "labelprint", "SYS", c) Then
'        Beep
'        MsgBox "ラベル印刷用コントロールＦの獲得に失敗しました。処理を中止して下さい。"
'        End
        PRINT_CHECK_F = False   '2008.09.26
    Else
        LabelPrint_F = RTrim(c)
    End If
'    LabelPrint_F = RTrim(c)
                                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> P_SYS.INI--> PM00040.INI 2013.09.03
                                
                                '原産国印字有無 2008.06.12
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2016.01.29  PM00040.INI --> P_SYS.INI[PLABEL]
'''    If GetIni(App.EXEName, "GENSANKOKU_DEF_F", "P_SYS", c) Then
''    If GetIni(App.EXEName, "GENSANKOKU_DEF_F", App.EXEName, c) Then
''        GENSANKOKU_FLG = "0"
''    Else
''        GENSANKOKU_FLG = RTrim(c)
''    End If
                                
    If GetIni("PLABEL", "GENSANKOKU_DEF_F", "P_SYS", c) Then
        GENSANKOKU_FLG = "0"
    Else
        GENSANKOKU_FLG = RTrim(c)
    End If
                                
                                
                                
                                '会社事業部エラーﾁｪｯｸ有無 2008.09.26
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2016.01.29  PM00040.INI --> P_SYS.INI[PLABEL]
'''    If GetIni(App.EXEName, "KAISYA_CHECK", "P_SYS", c) Then
''    If GetIni(App.EXEName, "KAISYA_CHECK", App.EXEName, c) Then
''        KAISYA_CHK_F = False
''    Else
''
''        If Trim(c) = "1" Then
''            KAISYA_CHK_F = True
''        Else
''            KAISYA_CHK_F = False
''        End If
''    End If
                                
    If GetIni("PLABEL", "KAISYA_CHECK", "P_SYS", c) Then
        KAISYA_CHK_F = False
    Else

        If Trim(c) = "1" Then
            KAISYA_CHK_F = True
        Else
            KAISYA_CHK_F = False
        End If
    End If
                                
                                '原産国空白ﾁｪｯｸ 2009.03.28
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2016.01.29  PM00040.INI --> P_SYS.INI[PLABEL]
'''    If GetIni(App.EXEName, "GENSANKOKU_CHECK", "P_SYS", c) Then
''    If GetIni(App.EXEName, "GENSANKOKU_CHECK", App.EXEName, c) Then
''        ReDim GENSANKOKU_CHECK_TBL(0 To 0)
''        GENSANKOKU_CHECK_TBL(0) = "*"
''    Else
''        GENSANKOKU_CHECK_TBL = Split(Trim(c), ",", -1)
''    End If
                                
    If GetIni("PLABEL", "GENSANKOKU_CHECK", "P_SYS", c) Then
        ReDim GENSANKOKU_CHECK_TBL(0 To 0)
        GENSANKOKU_CHECK_TBL(0) = "*"
    Else
        GENSANKOKU_CHECK_TBL = Split(Trim(c), ",", -1)
    End If
                                
                                
                                
                                '単価空白ﾁｪｯｸ 2010.02.02
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2016.01.29  PM00040.INI --> P_SYS.INI[PLABEL]
'''    If GetIni(App.EXEName, "TANKA_SPACE_F", "P_SYS", c) Then
''    If GetIni(App.EXEName, "TANKA_SPACE_F", App.EXEName, c) Then
''        TANKA_SPACE_F = "0"
''    Else
''        If Trim(c) = "1" Then
''            TANKA_SPACE_F = "1"
''        Else
''            TANKA_SPACE_F = "0"
''        End If
''    End If
                                
                                
    If GetIni("PLABEL", "TANKA_SPACE_F", "P_SYS", c) Then
        TANKA_SPACE_F = "0"
    Else
        If Trim(c) = "1" Then
            TANKA_SPACE_F = "1"
        Else
            TANKA_SPACE_F = "0"
        End If
    End If
                                
                                '代表機種ﾁｪｯｸ   2012.09.03
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2016.01.29  PM00040.INI --> P_SYS.INI[PLABEL]
'''    If GetIni(App.EXEName, "KISHU_CHECK", "P_SYS", c) Then
''    If GetIni(App.EXEName, "KISHU_CHECK", App.EXEName, c) Then
''        KISHU_CHECK = False
''    Else
''        If Trim(c) = "1" Then
''            KISHU_CHECK = True
''        Else
''            KISHU_CHECK = False
''        End If
''    End If
                               
    If GetIni("PLABEL", "KISHU_CHECK", "P_SYS", c) Then
        KISHU_CHECK = 0
    Else
        If Trim(c) = "1" Then
            KISHU_CHECK = 1
        Else
            KISHU_CHECK = 0
        End If
    End If
                                '海外供給区分ﾁｪｯｸ   2016.01.29
    If GetIni("PLABEL", "GAI_BUHIN_CHK", "P_SYS", c) Then
        GAI_BUHIN_CHECK = 0
    Else
        If Trim(c) = "1" Then
            GAI_BUHIN_CHECK = 1
        Else
            GAI_BUHIN_CHECK = 0
        End If
    End If
                                
                                
                                
                                
                                
                                
                                '会社名非表示設定   2013.03.29
'    If GetIni(App.EXEName, "KAISYA_RESTRICT_F", "P_SYS", c) Then
    If GetIni(App.EXEName, "KAISYA_RESTRICT_F", App.EXEName, c) Then
        KAISYA_RESTRICT_F = False
    Else
        If Trim(c) = "1" Then
            KAISYA_RESTRICT_F = True
        Else
            KAISYA_RESTRICT_F = False
        End If
    End If
                                
                                
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
        
    Combo1(pcmbJGYOBU).Clear
    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If

        Combo1(pcmbJGYOBU).AddItem RTrim(JGYOBU_T(i).NAME) & "                             " & JGYOBU_T(i).CODE

        
    Next i
        
        
    For i = 0 To Combo1(pcmbJGYOBU).ListCount - 1
    
        
        If Right(Combo1(pcmbJGYOBU).List(i), 1) = Last_JGYOBU Then
        
            Combo1(pcmbJGYOBU).ListIndex = i
            Exit For
        
        End If
    
    Next i
        
        
        
        
'    For i = 0 To UBound(JGYOBU_T)
'        If JGYOBU_T(i).CODE = " " Then
'            Unload SubMenu(i)
'            Exit For
'        End If
'
'        Load SubMenu(i + 1)
'        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)
'
'        If JGYOBU_T(i).CODE = Last_JGYOBU Then
'            PM000402.Caption = "商品化システム　品目マスタメンテナンス（商品ラベル項目）（" + RTrim(JGYOBU_T(i).NAME) + ")"
'            SubMenu(i).Checked = True
'            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
'            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'        Else
'            SubMenu(i).Checked = False
'        End If
'    Next i
'
'    Unload SubMenu(i)
                                
                                
    PM000402.Caption = PM000402.Caption & " " & Last_Update_Day
                                
                                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.23
'                                '品目マスタＯＰＥＮ
'    If ITEM_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'
'                                '品目マスタＯＰＥＮ
'    If L_ITEM_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'
'                                '原産国マスタＯＰＥＮ
'    If GENSAN_Open(BtOpenNomal) Then
'        Beep
'        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'        Unload Me
'    End If
'                                'コードマスタＯＰＥＮ
'    If P_CODE_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'                                '受払先マスタ（仕入先）ＯＰＥＮ
'    If P_UKEHARAI_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'
'    Call P_CODE_TBL_Proc
'
'
'    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.06.01
'                                'PNマスタＯＰＥＮ
'    If PN_M_Open(0) Then
'        Beep
'        MsgBox "システム異常が発生しました。処理を中止して下さい。"
'        Unload Me
'    End If
'    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.23
    

    '2012.01.13
    Call File_Open_Proc


    Combo1(pcmbNAIGAI).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo1(pcmbNAIGAI).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    Combo1(pcmbNAIGAI).ListIndex = 0

    '会社名のセット
    If Code_Set_Proc(pcmbL_KAISHA, P_KBN07_CD) Then
        Unload Me
    End If

    '事業部名のセット
    If Code_Set_Proc(pcmbL_JGYOBU, P_KBN07_CD) Then
        Unload Me
    End If

    Text1(ptxL_QTY).Text = "1"              '2008.10.03


    INIT_FLG = False



    '会社ｺｰﾄﾞ 空白時の設定値取込み          2013.01.04
    If GetIni(App.EXEName, "KAISYA_DEF_VALUE", "P_SYS", c) Then
        KAISYA_DEF_VALUE = ""
    Else
        KAISYA_DEF_VALUE = Trim(c)
    End If

    '事業部ｺｰﾄﾞ 空白時の設定値取込み        2013.01.04
    If GetIni(App.EXEName, "JIGYOBU_DEF_VALUE", "P_SYS", c) Then
        JIGYOBU_DEF_VALUE = ""
    Else
        JIGYOBU_DEF_VALUE = Trim(c)
    End If

    '会社名表示制限コード(会社ｺｰﾄﾞ)         2013.01.05
    If GetIni(App.EXEName, "KAISYA_RESTRICT", "P_SYS", c) Then
        KAISYA_RESTRICT = ""
    Else
        KAISYA_RESTRICT = Trim(c)
    End If


    'ｺﾏﾝﾄﾞﾗｲﾝ引数（外部品番）有⇒自動表示   2013.01.05
    c = Command
    'RETURN_F = False                    '2013.03.31
    If Trim(c) <> "" Then
        Show
        DoEvents

        
        'RETURN_F = True                 '2013.03.31
        wkJGYOBU = Mid(c, 1, 1)
        wkNAIGAI = Mid(c, 2, 1)
        
        For i = 0 To Combo1(pcmbJGYOBU).ListCount
            If Right(Combo1(pcmbJGYOBU).List(i), 1) = wkJGYOBU Then
                Combo1(pcmbJGYOBU).ListIndex = i
                Exit For
            End If
        Next i
        Last_JGYOBU = wkJGYOBU


        For i = 0 To Combo1(pcmbNAIGAI).ListCount
            If Right(Combo1(pcmbNAIGAI).List(i), 1) = wkNAIGAI Then
                Combo1(pcmbNAIGAI).ListIndex = i
                Exit For
            End If
        Next i


        Text1(ptxHIN_GAI).Text = Mid(StrConv(Trim(c), vbUpperCase), 3, Len(Trim(c)) - 2)

        If Error_Check_Proc(ptxHIN_GAI) Then     'エラーチェック
            Exit Sub
        End If

        Text1(ptxL_MAISU).SetFocus
    End If


End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer



    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.06.01
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
    
                                            
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            
                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "コードマスタ")
        End If
    End If
                                            '受払先マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "受払先マスタ")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PM000402 = Nothing

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
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
        
    If Index = ptxHIN_GAI Then
        Text1(ptxHIN_GAI).Text = StrConv(RTrim(Text1(ptxHIN_GAI).Text), vbUpperCase)
    End If
        
        
    If Error_Check_Proc(Index) Then     'エラーチェック
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Function Code_Set_Proc(Index As Integer, KBN As String) As Integer
'----------------------------------------------------------------------------
'                   コードマスタをコンボにセットする。
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim Key_Len     As Integer
Dim OPTION1     As Integer
Dim OPTION2     As Integer

Dim wkOption    As String



Dim i           As Integer
    
    Code_Set_Proc = True
    
    Combo1(Index).Clear
    
    For i = 0 To UBound(P_KBN_TBL)
    
        If KBN = P_KBN_TBL(i).KBN_CD Then
            Key_Len = P_KBN_TBL(i).KBN_Len
            Exit For
        End If
    
    Next i
    
    If i > UBound(P_KBN_TBL) Then
        Exit Function
    End If
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, KBN)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")

    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> KBN Then
                    
                    Exit Do
                
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "コードマスタ")
                Exit Function
        
        End Select

        wkOption = ""
        If P_KBN_TBL(i).KBN_OP1 Then
            wkOption = Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
        End If
        If P_KBN_TBL(i).KBN_OP2 Then
            wkOption = wkOption & Trim(StrConv(P_CODEREC.OPTION2, vbUnicode))
        End If
        
        
        
        Combo1(Index).AddItem StrConv(P_CODEREC.C_NAME, vbUnicode) & "                                        " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), Key_Len) & wkOption
        
        
        com = BtOpGetNext
    
    Loop

    Code_Set_Proc = False
    



End Function



Private Sub Clear_Proc()
    
    
Dim i   As Integer
    
    
    For i = ptxHIN_GAI To ptxL_MARK
        Text1(i).Text = ""
    Next i

    For i = prchL_SAGYO_SHIJI To prchL_KISHU_BIKOU
        RichTextBox1(i).Text = ""
    Next i


    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.23
'    For i = pcmbL_KAISHA To pcmbL_JGYOBU
    For i = pcmbL_KAISHA To pcmbGENSAN
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2012.01.23

        Combo1(i).ListIndex = -1
    Next i

    Text1(ptxL_QTY).Text = "1"

    '2008.12.19
    Text1(ptxL_MAISU).Text = "1"

    
    Call UniCode_Conv(ITEMREC.HIN_GAI, "")


    Check1(pchkGENSANKOKU).Value = vbUnchecked


    Text1(ptxHIN_GAI).SetFocus

End Sub

Private Sub Text1_LostFocus(Index As Integer)

    If Index = ptxHIN_GAI Then
        Text1(ptxHIN_GAI).Text = StrConv(RTrim(Text1(ptxHIN_GAI).Text), vbUpperCase)
    End If

End Sub
Private Function GENSANKOKU_SET_Proc(TORI_GENSANKOKU As String) As Integer
'----------------------------------------------------------------------------
'                   原産国マスタをコンボにセットする。
'----------------------------------------------------------------------------

Dim sts     As Integer
Dim com     As Integer
Dim i       As Integer

Dim wkTORI_GENSANKOKU   As String   '2013.03.27

    GENSANKOKU_SET_Proc = True
    
    
    
    'NULL 空白変換  2013.03.27
    wkTORI_GENSANKOKU = ""
    For i = 1 To Len(TORI_GENSANKOKU)
                
        If Mid(TORI_GENSANKOKU, i, 1) < " " Then
            wkTORI_GENSANKOKU = wkTORI_GENSANKOKU & " "
        Else
            wkTORI_GENSANKOKU = wkTORI_GENSANKOKU & Mid(TORI_GENSANKOKU, i, 1)
        End If
    
    Next i
    
    
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    TORI_GENSANKOKUの有無チェック＆書き込み   2012.01.23
    
    If Trim(wkTORI_GENSANKOKU) = "" Then
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
                'Call UniCode_Conv(GENSANREC.GENSANKOKU, TORI_GENSANKOKU)       '2013.03.31
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
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    TORI_GENSANKOKUの有無チェック＆書き込み   2012.01.23
    
    Combo1(pcmbGENSAN).Clear
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
                Call File_Error(sts, com, "原産国マスタ")
                Exit Function
        End Select
    
        
        'List2.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)        2013.02.19
'        If Trim(StrConv(GENSANREC.UPD_DATETIME, vbUnicode)) = "" Then                                               '2013.02.19    2014.02.18
'            List2.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2013.02.19    2014.02.18
'        Else                                                                                                        '2013.02.19    2014.02.18
'            List2.AddItem StrConv(GENSANREC.UPD_DATETIME, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2013.02.19    2014.02.18
'        End If                                                                                                      '2013.02.19    2014.02.18
        
        
        
        If StrConv(GENSANREC.UPD_DATETIME, vbUnicode) > StrConv(GENSANREC.Ins_DateTime, vbUnicode) Then             '2014.02.18
            List2.AddItem StrConv(GENSANREC.UPD_DATETIME, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2014.02.18
        Else                                                                                                        '2014.02.18
            List2.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2014.02.18
        End If
        
        
        com = BtOpGetNext
    Loop
        
    Combo1(pcmbGENSAN).Enabled = False
    Text1(ptxGENSANKOKU).Enabled = False
        
    If List2.ListCount > 0 Then
        Combo1(pcmbGENSAN).Enabled = True
        For i = List2.ListCount - 1 To 0 Step -1
        
            Combo1(pcmbGENSAN).AddItem Right(List2.List(i), 20)
        
        Next i
    
        Combo1(pcmbGENSAN).ListIndex = 0
    Else
        Text1(ptxGENSANKOKU).Enabled = True
    End If
    
    GENSANKOKU_SET_Proc = False


End Function



Public Sub File_Open_Proc()
'----------------------------------------------------------------------------
'               ファイル　ＯＰＥＮ処理
'           2012.01.23
'----------------------------------------------------------------------------
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '品目マスタＯＰＥＮ
    If L_ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '原産国マスタＯＰＥＮ
    If GENSAN_Open(BtOpenNomal) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenRead) Then
        Unload Me
    End If
                                '受払先マスタ（仕入先）ＯＰＥＮ
    If P_UKEHARAI_Open(BtOpenRead) Then
        Unload Me
    End If
                                
    Call P_CODE_TBL_Proc
                                
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.06.01
                                'PNマスタＯＰＥＮ
    If PN_M_Open(0) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<    2012.01.23  --> 削除 2012.02.26
                                'カントリーマスタ ＯＰＥＮ
'    If Country_Open(BtOpenRead) Then
'        Unload Me
'    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<    2012.01.23  --> 削除 2012.02.26


End Sub

Private Function Mesg_Set_Proc(GEN_NG_F As Integer, GEN_AT_GAI_F As Integer, GEN_AT_PLU_F As Integer, TANKA_SP_F As Integer, KISHU_NG_F As Integer, KAISYA_NG_F As Integer, KISHU1 As String, KISHU2 As String) As Integer
'----------------------------------------------------------------------------
'               エラーメッセージ作成
'           2016.01.29
'----------------------------------------------------------------------------
Dim Mesg        As String
Dim i           As Integer
    
    
Dim GENSANKOKU  As String * 20
    
Dim KAISYA_NAME As String * 20
Dim JGYOBU_NAME As String * 20
    
    
Dim TANKA       As String * 20
    
Dim KISHU       As String * 20
    
Dim Tanka2      As String * 9
Dim Tanka3      As String * 9
    
        Mesg = "パーツラベルの内容を確認して下さい" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
        Mesg = Mesg & "○品番   " & StrConv(ITEMREC.HIN_GAI, vbUnicode) & Chr(13) & Chr(10)
        Mesg = Mesg & "○品名   " & StrConv(ITEMREC.HIN_NAME, vbUnicode) & Chr(13) & Chr(10)
        Mesg = Mesg & "○品名Ｅ " & StrConv(ITEMREC.L_HIN_NAME_E, vbUnicode) & Chr(13) & Chr(10)

    
        GENSANKOKU = ""
        If Text1(ptxGENSANKOKU).Enabled Then
            GENSANKOKU = Trim(Text1(ptxGENSANKOKU).Text)
        Else
            If Combo1(pcmbGENSAN).Enabled Then
                GENSANKOKU = Trim(Left(Combo1(pcmbGENSAN).Text, 20))
            End If
        End If




        Select Case GEN_NG_F
            Case 0
                If GEN_AT_PLU_F < 2 Then
                    If GEN_AT_GAI_F = 0 Then
                        Mesg = Mesg & "○原産国 " & GENSANKOKU & Chr(13) & Chr(10)
                    Else
                    
                        Mesg = Mesg & "△原産国 " & GENSANKOKU & "　←原産国注意（海外向け）" & Chr(13) & Chr(10)
                
                    End If
                Else
                        Mesg = Mesg & "△原産国 "
                        For i = 0 To Combo1(pcmbGENSAN).ListCount - 1
                            GENSANKOKU = Combo1(pcmbGENSAN).List(i)
                            
                            If i = 0 Then
                                Mesg = Mesg & GENSANKOKU
                            Else
                                Mesg = Mesg & "　　　　　   " & GENSANKOKU
                            End If
                            If i = Combo1(pcmbGENSAN).ListIndex Then
                                Mesg = Mesg & "　←原産国注意（複数）" & Chr(13) & Chr(10)
                            Else
                                Mesg = Mesg & Chr(13) & Chr(10)
                            End If
                        Next i
                End If
            
            Case 1
                Mesg = Mesg & "×原産国 " & GENSANKOKU & "　←空白です" & Chr(13) & Chr(10)
            Case 9
                Mesg = Mesg & "○原産国 " & Chr(13) & Chr(10)
        End Select
    
    
        Mesg = Mesg & "○供給区分海外 " & "     " & StrConv(ITEMREC.GAI_BUHIN, vbUnicode) & Chr(13) & Chr(10)
    
        Tanka2 = Text1(ptxL_URIKIN2).Text
        Tanka3 = Text1(ptxL_URIKIN3).Text
        TANKA = Tanka2 & " " & Tanka3
        If TANKA_SP_F = 1 Then
            Mesg = Mesg & "×単価　 " & TANKA & "   　  ←空白です" & Chr(13) & Chr(10)
        Else
            Mesg = Mesg & "○単価　 " & TANKA & Chr(13) & Chr(10)
        End If
    
        KISHU = KISHU1
        If KISHU_NG_F = 1 Then
            Mesg = Mesg & "×代表機種　 " & KISHU & " ←空白です" & Chr(13) & Chr(10)
        Else
            Mesg = Mesg & "○代表機種　 " & KISHU & Chr(13) & Chr(10)
        End If
    
        If KAISYA_NG_F = 9 Then
        Else
            KAISYA_NAME = Combo1(pcmbL_KAISHA).Text
            JGYOBU_NAME = Combo1(pcmbL_JGYOBU).Text
            If KAISYA_NG_F = 1 Then
                If Trim(KAISYA_NAME) = "" Then
                    Mesg = Mesg & "×会社名 " & KAISYA_NAME & " " & "　　 ←空白です" & Chr(13) & Chr(10)
                Else
                    Mesg = Mesg & "○会社名 " & KAISYA_NAME & Chr(13) & Chr(10)
                End If
                If Trim(JGYOBU_NAME) = "" Then
                    Mesg = Mesg & "×事業部名 " & JGYOBU_NAME & " " & "　←空白です" & Chr(13) & Chr(10)
                Else
                    Mesg = Mesg & "○事業部名 " & JGYOBU_NAME & Chr(13) & Chr(10)
                End If
            Else
                    Mesg = Mesg & "○会社名 " & KAISYA_NAME & Chr(13) & Chr(10)
                    Mesg = Mesg & "○事業部名 " & JGYOBU_NAME & Chr(13) & Chr(10)
            End If
        End If
    
    
        Mesg = Mesg & Chr(13) & Chr(10)
    
    
        Mesg = Mesg & "　　　　【ＯＫ】パーツラベルを印刷" & Chr(13) & Chr(10)
        Mesg = Mesg & " 【キャンセル】印刷中止" & Chr(13) & Chr(10)
    
    
    
        Mesg_Set_Proc = MsgBox(Mesg, vbOKCancel + vbDefaultButton2 + vbExclamation, "パーツラベル項目確認")


End Function
