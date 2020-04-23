VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form PM000502 
   Caption         =   "商品化システム　構成マスタメンテナンス"
   ClientHeight    =   11010
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   14670
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
   ScaleHeight     =   11010
   ScaleWidth      =   14670
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Index           =   4
      Left            =   6120
      MaxLength       =   40
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1335
      Index           =   0
      Left            =   10320
      TabIndex        =   6
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2355
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"PM000502.frx":0000
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Index           =   3
      Left            =   3720
      MaxLength       =   40
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   3
      Text            =   "1234567890"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ﾘﾅﾝﾊﾞｰ"
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   8400
      TabIndex        =   41
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "全ﾁｪｯｸ"
      Height          =   375
      Index           =   1
      Left            =   6360
      TabIndex        =   40
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "一括削除"
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   39
      Top             =   360
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   0
      Left            =   1440
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   2805
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Index           =   1
      Left            =   4080
      MaxLength       =   40
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "戻 る"
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
      TabIndex        =   38
      Top             =   10440
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
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   10440
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
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   10440
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   10440
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
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   10440
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
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   10440
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
      Index           =   5
      Left            =   4800
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   10440
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
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   10440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "削 除"
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
      TabIndex        =   31
      Top             =   10440
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
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   10440
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
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   10440
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
      TabIndex        =   28
      Top             =   10440
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   8535
      Index           =   0
      Left            =   0
      TabIndex        =   44
      Top             =   1800
      Width           =   14655
      Begin VB.OptionButton Option1 
         Caption         =   "同梱／構成"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   840
         TabIndex        =   19
         Top             =   5040
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "外装資材"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   13
         Top             =   3000
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "個装資材"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Height          =   2055
         Index           =   1
         Left            =   120
         TabIndex        =   49
         Top             =   3000
         Width           =   10095
         Begin VB.TextBox txtG_KEY 
            Height          =   360
            Left            =   3000
            TabIndex        =   61
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  '右揃え
            Height          =   375
            IMEMode         =   3  'ｵﾌ固定
            Index           =   12
            Left            =   8400
            MaxLength       =   6
            TabIndex        =   17
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   375
            IMEMode         =   3  'ｵﾌ固定
            Index           =   9
            Left            =   120
            MaxLength       =   3
            TabIndex        =   14
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000F&
            Height          =   375
            IMEMode         =   3  'ｵﾌ固定
            Index           =   11
            Left            =   3360
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   720
            Width           =   4935
         End
         Begin VB.TextBox Text1 
            Height          =   375
            IMEMode         =   3  'ｵﾌ固定
            Index           =   10
            Left            =   720
            MaxLength       =   20
            TabIndex        =   15
            Top             =   720
            Width           =   2535
         End
         Begin VB.ListBox List1 
            Height          =   780
            Index           =   1
            ItemData        =   "PM000502.frx":00BE
            Left            =   120
            List            =   "PM000502.frx":00C0
            TabIndex        =   18
            Top             =   1200
            Width           =   9735
         End
         Begin VB.Label Label 
            Alignment       =   1  '右揃え
            Caption         =   "員数"
            Height          =   255
            Index           =   2
            Left            =   8400
            TabIndex        =   52
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label 
            Alignment       =   2  '中央揃え
            Caption         =   "品　　　名"
            Height          =   255
            Index           =   5
            Left            =   3360
            TabIndex        =   51
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label Label 
            Alignment       =   2  '中央揃え
            Caption         =   "品　　　番"
            Height          =   255
            Index           =   6
            Left            =   720
            TabIndex        =   50
            Top             =   480
            Width           =   2535
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3375
         Index           =   2
         Left            =   120
         TabIndex        =   45
         Top             =   5040
         Width           =   14415
         Begin VB.TextBox txtD_KEY 
            Height          =   360
            Left            =   4080
            TabIndex        =   62
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox Combo1 
            Height          =   360
            Index           =   1
            Left            =   720
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   21
            Top             =   600
            Width           =   1125
         End
         Begin VB.TextBox Text1 
            Height          =   375
            IMEMode         =   3  'ｵﾌ固定
            Index           =   17
            Left            =   9360
            MaxLength       =   40
            TabIndex        =   26
            Top             =   600
            Width           =   4935
         End
         Begin VB.ListBox List1 
            Height          =   1980
            Index           =   2
            ItemData        =   "PM000502.frx":00C2
            Left            =   120
            List            =   "PM000502.frx":00C4
            TabIndex        =   27
            Top             =   1080
            Width           =   14175
         End
         Begin VB.TextBox Text1 
            Height          =   375
            IMEMode         =   3  'ｵﾌ固定
            Index           =   14
            Left            =   1920
            MaxLength       =   20
            TabIndex        =   22
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000F&
            Height          =   375
            IMEMode         =   3  'ｵﾌ固定
            Index           =   15
            Left            =   4560
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox Text1 
            Height          =   375
            IMEMode         =   3  'ｵﾌ固定
            Index           =   13
            Left            =   120
            MaxLength       =   3
            TabIndex        =   20
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  '右揃え
            Height          =   375
            IMEMode         =   3  'ｵﾌ固定
            Index           =   16
            Left            =   8400
            MaxLength       =   6
            TabIndex        =   24
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label 
            Alignment       =   2  '中央揃え
            Caption         =   "種別"
            Height          =   255
            Index           =   12
            Left            =   840
            TabIndex        =   59
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label 
            Alignment       =   2  '中央揃え
            Caption         =   "備　　　考"
            Height          =   255
            Index           =   11
            Left            =   9720
            TabIndex        =   58
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label 
            Alignment       =   2  '中央揃え
            Caption         =   "品　　　番"
            Height          =   255
            Index           =   7
            Left            =   2040
            TabIndex        =   48
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label 
            Alignment       =   2  '中央揃え
            Caption         =   "品　　　名"
            Height          =   255
            Index           =   8
            Left            =   4560
            TabIndex        =   47
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label Label 
            Alignment       =   1  '右揃え
            Caption         =   "員数"
            Height          =   255
            Index           =   10
            Left            =   8400
            TabIndex        =   46
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2415
         Index           =   3
         Left            =   120
         TabIndex        =   53
         Top             =   600
         Width           =   10095
         Begin VB.TextBox txtK_KEY 
            Height          =   360
            Left            =   3000
            TabIndex        =   60
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Height          =   375
            IMEMode         =   3  'ｵﾌ固定
            Index           =   6
            Left            =   720
            MaxLength       =   20
            TabIndex        =   9
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000F&
            Height          =   375
            IMEMode         =   3  'ｵﾌ固定
            Index           =   7
            Left            =   3360
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   600
            Width           =   4935
         End
         Begin VB.TextBox Text1 
            Height          =   375
            IMEMode         =   3  'ｵﾌ固定
            Index           =   5
            Left            =   120
            MaxLength       =   3
            TabIndex        =   8
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  '右揃え
            Height          =   375
            IMEMode         =   3  'ｵﾌ固定
            Index           =   8
            Left            =   8400
            MaxLength       =   6
            TabIndex        =   11
            Top             =   600
            Width           =   855
         End
         Begin VB.ListBox List1 
            Height          =   1260
            Index           =   0
            ItemData        =   "PM000502.frx":00C6
            Left            =   120
            List            =   "PM000502.frx":00C8
            TabIndex        =   12
            Top             =   1080
            Width           =   9735
         End
         Begin VB.Label Label 
            Alignment       =   2  '中央揃え
            Caption         =   "品　　　番"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   56
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label 
            Alignment       =   2  '中央揃え
            Caption         =   "品　　　名"
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   55
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label Label 
            Alignment       =   1  '右揃え
            Caption         =   "員数"
            Height          =   255
            Index           =   4
            Left            =   8400
            TabIndex        =   54
            Top             =   360
            Width           =   855
         End
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
         Left            =   0
         TabIndex        =   57
         Top             =   8400
         Width           =   180
      End
   End
   Begin VB.Label Label 
      Caption         =   "内職ｸﾗｽ"
      Height          =   255
      Index           =   15
      Left            =   5160
      TabIndex        =   67
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "付加ｸﾗｽ"
      Height          =   255
      Index           =   14
      Left            =   2880
      TabIndex        =   66
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "備　考"
      Height          =   375
      Index           =   1
      Left            =   9240
      TabIndex        =   65
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "(30文字×4行)"
      Height          =   375
      Index           =   0
      Left            =   8520
      TabIndex        =   64
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label 
      Caption         =   "基本ｸﾗｽ"
      Height          =   255
      Index           =   13
      Left            =   480
      TabIndex        =   63
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "仕向け先"
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   43
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "品番"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   42
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "PM000502"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'テキスト用添字
Private Const ptxHIN_GAI% = 0               '品番
Private Const ptxHIN_NAME% = 1              '品名
Private Const ptxCLASS_CODE% = 2            '基本ｸﾗｽ
Private Const ptxF_CLASS_CODE% = 3          '付加ｸﾗｽ呼び名
Private Const ptxCLASS_NAME% = 4          '付加ｸﾗｽ呼び名    内職！  高沢修正


'2019.05.28 以下の項目順が全て１個ずれている！
Private Const ptxK_SEQNO% = 4 + 1             '個装資材　追番
Private Const ptxK_HIN_GAI% = 5 + 1            '個装資材　品番
Private Const ptxK_HIN_NAME% = 6 + 1           '個装資材　品名
Private Const ptxK_KO_QTY% = 7 + 1             '個装資材　員数

Private Const ptxG_SEQNO% = 8 + 1              '個装資材　追番
Private Const ptxG_HIN_GAI% = 9 + 1            '個装資材　品番
Private Const ptxG_HIN_NAME% = 10 + 1          '個装資材　品名
Private Const ptxG_KO_QTY% = 11 + 1            '個装資材　員数

Private Const ptxD_SEQNO% = 12 + 1             '同梱／構成　追番
Private Const ptxD_HIN_GAI% = 13 + 1           '同梱／構成　品番
Private Const ptxD_HIN_NAME% = 14 + 1          '同梱／構成　品名
Private Const ptxD_KO_QTY% = 15 + 1            '同梱／構成　員数
Private Const ptxD_BIKOU% = 16 + 1             '同梱／構成　員数


'コンボ用添字
Private Const pcmbSHIMUKE% = 0              '仕向け先
Private Const pcmbD_SYUBETSU% = 1           '種別

'ﾘｽﾄﾎﾞｯｸｽ用添字
Private Const plstK_ITEM% = 0               '個装資材
Private Const plstG_ITEM% = 1               '外装資材
Private Const plstD_ITEM% = 2               '同梱／構成

'ﾗｼﾞｵﾎﾞﾀﾝ用添字
Private Const poptK_ITEM% = 0               '個装資材
Private Const poptG_ITEM% = 1               '外装資材
Private Const poptD_ITEM% = 2               '同梱／構成

'(特殊)ｺﾏﾝﾄﾞﾎﾞﾀﾝ用添字
Private Const pcmbALLDEL% = 0               '一括削除
Private Const pcmbALLCHK% = 1               '一括ﾁｪｯｸ
Private Const pcmbRENUM% = 2                '追番ﾘﾅﾝﾊﾞｰ

'リッチテキスト用添字
Private Const prchBIKOU% = 0                '備考

Private INIT_FLG    As Boolean

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PM000502.MousePointer = vbHourglass

    Call Ctrl_Lock(PM000502)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM000502)


    PM000502.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
Dim com     As Integer
Dim ans     As Integer
Dim sts     As Integer

Dim i       As Integer
    
    Error_Check_Proc = True
    
    
    Select Case Mode
        
        Case ptxHIN_GAI      '品番
            
            
            
            
            
            
        
            If G_SCREEN_FLG = G_SCREEN_INS And _
                Not Text1(ptxHIN_GAI).Locked Then
                
                
                If Trim(Text1(ptxHIN_GAI).Text) = "" Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(ptxHIN_GAI).SetFocus
                    Exit Function
                End If
                
                
                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        Text1(ptxHIN_NAME).Text = ""
                        MsgBox "入力した項目はエラーです。"
                        Text1(ptxHIN_GAI).SetFocus
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        PM000502.Visible = False
                        INIT_FLG = False
                        Exit Function
                End Select
                Text1(ptxHIN_NAME).Text = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                
                
                
                
                '新規時は重複チェック
                
                Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
                Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                            
                Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
            
                Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_HEAD)
                Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
            
                sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                Select Case sts
                    Case BtNoErr
                        
                         
                        ans = MsgBox("入力したコードは、登録済です。更新処理として継続しますか？", vbYesNo, "確認入力")
                        If ans = vbNo Then
                            Text1(ptxHIN_GAI).SetFocus
                            Exit Function
                        End If
                
                
                        Call Item_Disp_Proc(Right(Combo1(pcmbSHIMUKE), 4) & Text1(ptxHIN_GAI).Text)
                    
                    Case BtErrKeyNotFound
                    Case Else
                        Call File_Error(sts, BtOpGetGreater, "構成マスタ")
                        Exit Function
                End Select
            
            
                Combo1(pcmbSHIMUKE).BackColor = G_INPUT_NG
                Combo1(pcmbSHIMUKE).Locked = True
                Combo1(pcmbSHIMUKE).TabStop = False
            
            
                Text1(ptxHIN_GAI).BackColor = G_INPUT_NG
                Text1(ptxHIN_GAI).Locked = True
                Text1(ptxHIN_GAI).TabStop = False
            
            
            End If
        
        Case ptxCLASS_CODE          '基本ｸﾗｽ
                
            'ｸﾗｽﾏｽﾀ読み込み
            If Trim(Text1(Mode).Text) <> "" Then
            
                Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
                Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxCLASS_CODE).Text)
                    
                sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                Select Case sts
                    Case BtNoErr
                        Text1(ptxCLASS_NAME).Text = StrConv(P_CLASSREC.CLASS_NAME, vbUnicode)
                    Case BtErrKeyNotFound
                        Text1(ptxCLASS_NAME).Text = ""
                        MsgBox "入力した項目はエラーです。"
                        Text1(ptxCLASS_CODE).SetFocus
                        Exit Function
                    
                    Case Else
                        Call File_Error(sts, BtOpGetGreater, "構成マスタ")
                        Exit Function
                End Select
            End If
        
        Case ptxK_SEQNO             '個装資材　追番
        
            If Option1(poptK_ITEM).Value Then
                If Not IsNumeric(Text1(ptxK_SEQNO).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(ptxK_SEQNO).SetFocus
                    Exit Function
                Else
                    If CInt(Text1(ptxK_SEQNO).Text) <= 0 Then
                        MsgBox "入力した項目はエラーです。"
                        Text1(ptxK_SEQNO).SetFocus
                        Exit Function
                    Else
                        Text1(ptxK_SEQNO).Text = Format(CInt(Text1(ptxK_SEQNO).Text), "000")
                    
                    
                        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
                        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                                    
                        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
                    
                        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_KOSOU)
                        Call UniCode_Conv(K0_P_COMPO.SEQNO, Text1(ptxK_SEQNO).Text)
                    
                    
                        sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Select Case sts
                            Case BtNoErr
                                
                                
                                
                            Case BtErrKeyNotFound
                                If List1(plstK_ITEM).ListCount >= 5 Then
                                    MsgBox "入力した項目はエラーです。"
                                    Text1(ptxK_SEQNO).SetFocus
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpGetGreater, "構成マスタ")
                                Exit Function
                        End Select
                    
                    
                    
                    End If
                End If
            
            End If
        
        Case ptxK_HIN_GAI          '個装資材　品番
        
            If Option1(poptK_ITEM).Value Then
                
                '親品番と等しいはエラー
                
                
                If Trim(Text1(ptxK_HIN_GAI).Text) = Trim(Text1(ptxHIN_GAI).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(ptxK_HIN_GAI).SetFocus
                    Exit Function
                End If
                
                
                '最初は仕向け先のコードで読み込み
                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxK_HIN_GAI).Text)


                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        
                        
                    Case BtErrKeyNotFound
                    
                        '資材で再読み込み
                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxK_HIN_GAI).Text)
                    
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                                
                                
                            Case BtErrKeyNotFound
                            
                                Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番です。")
                                
'                                MsgBox "入力した項目はエラーです。"
'                                Text1(ptxK_HIN_GAI).SetFocus
'                                Exit Function
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                Exit Function
                        End Select
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                End Select
            
                Text1(ptxK_HIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                txtK_KEY.Text = StrConv(ITEMREC.JGYOBU, vbUnicode) & StrConv(ITEMREC.NAIGAI, vbUnicode)
            
            
            End If
        
        
        Case ptxK_KO_QTY           '個装資材　員数
        
            If Option1(poptK_ITEM).Value Then
                If Not IsNumeric(Text1(ptxK_KO_QTY).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(ptxK_KO_QTY).SetFocus
                    Exit Function
                Else
                    If CDbl(Text1(ptxK_KO_QTY).Text) <= 0 Then
                        MsgBox "入力した項目はエラーです。"
                        Text1(ptxK_KO_QTY).SetFocus
                        Exit Function
                    Else
                        Text1(ptxK_KO_QTY).Text = Format(CDbl(Text1(ptxK_KO_QTY).Text), "#0.00")
                    End If
                End If
            
            End If
        
        Case ptxG_SEQNO            '外装資材　追番
        
            If Option1(poptG_ITEM).Value Then
                If Not IsNumeric(Text1(ptxG_SEQNO).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(ptxG_SEQNO).SetFocus
                    Exit Function
                Else
                    If CInt(Text1(ptxG_SEQNO).Text) <= 0 Then
                        MsgBox "入力した項目はエラーです。"
                        Text1(ptxG_SEQNO).SetFocus
                        Exit Function
                    Else
                        Text1(ptxG_SEQNO).Text = Format(CInt(Text1(ptxG_SEQNO).Text), "000")
                    
                    
                    
                        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
                        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                                    
                        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
                    
                        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_GAISOU)
                        Call UniCode_Conv(K0_P_COMPO.SEQNO, Text1(ptxG_SEQNO).Text)
                    
                    
                        sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Select Case sts
                            Case BtNoErr
                                
                                
                                
                            Case BtErrKeyNotFound
                                If List1(plstG_ITEM).ListCount >= 3 Then
                                    MsgBox "入力した項目はエラーです。"
                                    Text1(ptxG_SEQNO).SetFocus
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpGetGreater, "構成マスタ")
                                Exit Function
                        End Select
                    
                    
                    
                    
                    
                    
                    End If
                End If
            
            End If
        
        Case ptxG_HIN_GAI          '外装資材　品番
        
            If Option1(poptG_ITEM).Value Then
                
                If Trim(Text1(ptxG_HIN_GAI).Text) = Trim(Text1(ptxHIN_GAI).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(ptxG_HIN_GAI).SetFocus
                    Exit Function
                End If
                
                
                
                '最初は仕向け先のコードで読み込み
                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxG_HIN_GAI).Text)


                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        
                        
                    Case BtErrKeyNotFound
                    
                        '資材で再読み込み
                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxG_HIN_GAI).Text)
                    
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                                
                                
                            Case BtErrKeyNotFound
                            
                                Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番です。")
                                
'                                MsgBox "入力した項目はエラーです。"
'                                Text1(ptxG_HIN_GAI).SetFocus
'                                Exit Function
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                Exit Function
                        End Select
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                End Select
            
                Text1(ptxG_HIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                txtG_KEY.Text = StrConv(ITEMREC.JGYOBU, vbUnicode) & StrConv(ITEMREC.NAIGAI, vbUnicode)
            
            
            
            End If
        
        
        Case ptxG_KO_QTY           '外装資材　員数
        
            If Option1(poptG_ITEM).Value Then
                If Not IsNumeric(Text1(ptxG_KO_QTY).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(ptxG_KO_QTY).SetFocus
                    Exit Function
                Else
                    If CDbl(Text1(ptxG_KO_QTY).Text) <= 0 Then
                        MsgBox "入力した項目はエラーです。"
                        Text1(ptxG_KO_QTY).SetFocus
                        Exit Function
                    Else
                        Text1(ptxG_KO_QTY).Text = Format(CDbl(Text1(ptxG_KO_QTY).Text), "#0.00")
                    End If
                End If
            
            End If
        
        
        
        Case ptxD_SEQNO            '同梱/構成　追番
        
            If Option1(poptD_ITEM).Value Then
                If Not IsNumeric(Text1(ptxD_SEQNO).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(ptxD_SEQNO).SetFocus
                    Exit Function
                Else
                    If CInt(Text1(ptxD_SEQNO).Text) <= 0 Then
                        MsgBox "入力した項目はエラーです。"
                        Text1(ptxD_SEQNO).SetFocus
                        Exit Function
                    Else
                        Text1(ptxD_SEQNO).Text = Format(CInt(Text1(ptxD_SEQNO).Text), "000")
                    
                    
                        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
                        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                                    
                        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
                    
                        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_DOUKON)
                        Call UniCode_Conv(K0_P_COMPO.SEQNO, Text1(ptxD_SEQNO).Text)
                    
                    
                        sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Select Case sts
                            Case BtNoErr
                                
                                
                                
                            Case BtErrKeyNotFound
                            
                                If List1(plstD_ITEM).ListCount >= 50 Then
                                    MsgBox "入力した項目はエラーです。"
                                    Text1(ptxD_SEQNO).SetFocus
                                    Exit Function
                                End If
                            
                            Case Else
                                Call File_Error(sts, BtOpGetGreater, "構成マスタ")
                                Exit Function
                        End Select
                    
                    
                    
                    
                    End If
                End If
            
            End If
        
        Case ptxD_HIN_GAI          '同梱/構成　品番
        
            If Option1(poptD_ITEM).Value Then
                
                If Trim(Text1(ptxD_HIN_GAI).Text) = Trim(Text1(ptxHIN_GAI).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(ptxG_HIN_GAI).SetFocus
                    Exit Function
                End If
                
                
                '最初は仕向け先のコードで読み込み
                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxD_HIN_GAI).Text)


                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        
                        
                    Case BtErrKeyNotFound
                    
                        '資材で再読み込み
                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxD_HIN_GAI).Text)
                    
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                                
                                
                            Case BtErrKeyNotFound
                            
                                Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番です。")
                                
'                                MsgBox "入力した項目はエラーです。"
'                                Text1(ptxD_HIN_GAI).SetFocus
'                                Exit Function
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                Exit Function
                        End Select
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                End Select
            
            
                Text1(ptxD_HIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                txtD_KEY.Text = StrConv(ITEMREC.JGYOBU, vbUnicode) & StrConv(ITEMREC.NAIGAI, vbUnicode)
            Else
                If Trim(Text1(ptxD_HIN_GAI)) <> "" Then
                    MsgBox "「同梱／構成」にチェックマークありません。", vbExclamation
                End If
            End If
        
        
        Case ptxD_KO_QTY           '同梱/構成　員数
        
            If Option1(poptD_ITEM).Value Then
                If Not IsNumeric(Text1(ptxD_KO_QTY).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(ptxD_KO_QTY).SetFocus
                    Exit Function
                Else
                    If CDbl(Text1(ptxD_KO_QTY).Text) <= 0 Then
                        MsgBox "入力した項目はエラーです。"
                        Text1(ptxD_KO_QTY).SetFocus
                        Exit Function
                    Else
                        Text1(ptxD_KO_QTY).Text = Format(CDbl(Text1(ptxD_KO_QTY).Text), "#0.00")
                    End If
                End If
            
            End If
        
    End Select
        
    Error_Check_Proc = False


End Function
Private Function Item_Disp_Proc(CODE As String) As Integer
'----------------------------------------------------------------------------
'                   画面表示
'----------------------------------------------------------------------------
Dim sts As Integer
Dim i   As Integer


    Item_Disp_Proc = True
    
    '構成ﾏｽﾀ読み込み
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(CODE, 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(CODE, 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(CODE, 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Mid(CODE, 5, 20))
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_HEAD)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
    
    
    sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    Select Case sts
        Case BtNoErr
            
            'ﾚｺｰﾄﾞ内容の表示
            For i = 0 To Combo1(pcmbSHIMUKE).ListCount - 1
            
                If StrConv(P_COMPO_O_REC.SHIMUKE_CODE, vbUnicode) = Left(Right(Combo1(pcmbSHIMUKE).List(i), 4), 2) Then
            
                    Combo1(pcmbSHIMUKE).ListIndex = i
                    
                    Exit For
            
                End If
            
            Next
                                            '品目ｺｰﾄﾞ
            Text1(ptxHIN_GAI).Text = Trim(StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode))
                                            '品名(品目ﾏｽﾀより)
            Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(CODE, 3, 1))
            Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(CODE, 4, 1))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Mid(CODE, 5, 20))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    PM000502.Visible = False
                    INIT_FLG = False
                    Exit Function
            End Select
            Text1(ptxHIN_NAME).Text = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
    
            Text1(ptxCLASS_CODE).Text = Trim(StrConv(P_COMPO_O_REC.CLASS_CODE, vbUnicode))
                    
            'ｸﾗｽﾏｽﾀ読み込み
            Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
            Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxCLASS_CODE).Text)
                
            sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
            Select Case sts
                Case BtNoErr
'                    Text1(ptxCLASS_NAME).Text = StrConv(P_CLASSREC.CLASS_NAME, vbUnicode)
                Case BtErrKeyNotFound
'                    Text1(ptxCLASS_NAME).Text = ""
                
                Case Else
                    Call File_Error(sts, BtOpGetGreater, "クラスマスタ")
                    PM000502.Visible = False
                    INIT_FLG = False
                    Exit Function
            End Select
                                        '備考
            RichTextBox1(prchBIKOU).Text = Trim(StrConv(P_COMPO_O_REC.BIKOU, vbUnicode))
        
        Case BtErrKeyNotFound
        
            MsgBox "他端末で変更されています。前画面に戻ります。"
            PM000502.Visible = False
            INIT_FLG = False
            
            Exit Function
                    
        Case Else
            Call File_Error(sts, BtOpGetEqual, "構成マスタ")
            PM000502.Visible = False
            INIT_FLG = False
            Exit Function
    
    End Select
                                                
        
        
                    
        
        
    If K_List_Disp_Proc() Then
        PM000502.Visible = False
        INIT_FLG = False
        Exit Function
    End If

    If G_List_Disp_Proc() Then
        PM000502.Visible = False
        INIT_FLG = False
        Exit Function
    End If

    If D_List_Disp_Proc() Then
        PM000502.Visible = False
        INIT_FLG = False
        Exit Function
    End If
        
        
        
        

    Item_Disp_Proc = False

End Function

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   構成マスタ出力
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim com     As Integer
Dim i       As Integer

    Update_Proc = True
    
    '--------------------------------------------   ﾍｯﾀﾞｰﾚｺｰﾄﾞ
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_HEAD)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")

    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "構成マスタ")
                Exit Function
        
        End Select


    Loop


    If com = BtOpInsert Then
        Call UniCode_Conv(P_COMPO_O_REC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
        Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
        Call UniCode_Conv(P_COMPO_O_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
        Call UniCode_Conv(P_COMPO_O_REC.HIN_GAI, Text1(ptxHIN_GAI).Text)
        Call UniCode_Conv(P_COMPO_O_REC.DATA_KBN, P_HEAD)
        Call UniCode_Conv(P_COMPO_O_REC.SEQNO, "000")
    
        Call UniCode_Conv(P_COMPO_O_REC.FILLER, "")
    
    
    End If



    Call UniCode_Conv(P_COMPO_O_REC.CLASS_CODE, Text1(ptxCLASS_CODE).Text)                  'ｸﾗｽｺｰﾄﾞ
    
    Call UniCode_Conv(P_COMPO_O_REC.BIKOU, RichTextBox1(prchBIKOU%).Text)                   '備考

    Call UniCode_Conv(P_COMPO_O_REC.BIKOU, "AAAAA" & ChrW(1))                      '備考




    Call UniCode_Conv(P_COMPO_O_REC.UPD_TANTO, "")                                          '更新担当者ｺｰﾄﾞ
                                                                                            '更新日時
    Call UniCode_Conv(P_COMPO_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))


    Do
        
        DoEvents
        
        sts = BTRV(com, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                    Update_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, com, "構成マスタ")
                Exit Function
        End Select
    
    Loop
    
    
    
    '--------------------------------------------   個装資材
    If Option1(poptK_ITEM).Value Then
    
        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_KOSOU)
        Call UniCode_Conv(K0_P_COMPO.SEQNO, Text1(ptxK_SEQNO).Text)
    
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound
                    com = BtOpInsert
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Update_Proc = False
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "構成マスタ")
                    Exit Function
            
            End Select
    
    
        Loop
    
    
        If com = BtOpInsert Then
            Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
            Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
            Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
            Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, Text1(ptxHIN_GAI).Text)
            Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_KOSOU)
            Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(CInt(Text1(ptxK_SEQNO).Text), "000"))
        
            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, "")
            
            Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, "")
        
            Call UniCode_Conv(P_COMPO_K_REC.FILLER, "")
        
        
        End If
    
    
        Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, Mid(txtK_KEY.Text, 1, 1))                        '子　事業部
        Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, Mid(txtK_KEY.Text, 2, 1))                        '子　国内外
        Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, Text1(ptxK_HIN_GAI).Text)                       '子　品番
        Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(CDbl(Text1(ptxK_KO_QTY).Text), "000.00"))    '子　員数
    
    
        Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, "")                                              '更新担当者ｺｰﾄﾞ
                                                                                                    '更新日時
        Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
        Do
            
            DoEvents
            
            sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Update_Proc = False
                        Exit Do
                    End If
                Case Else
                    Call File_Error(sts, com, "構成マスタ")
                    Exit Function
            End Select
        
        Loop
    
    
    
    
    
    End If
    
    
    '--------------------------------------------   外装資材
    If Option1(poptG_ITEM).Value Then
    
        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_GAISOU)
        Call UniCode_Conv(K0_P_COMPO.SEQNO, Text1(ptxG_SEQNO).Text)
    
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound
                    com = BtOpInsert
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Update_Proc = False
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "構成マスタ")
                    Exit Function
            
            End Select
    
    
        Loop
    
    
        If com = BtOpInsert Then
            Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
            Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
            Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
            Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, Text1(ptxHIN_GAI).Text)
            Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_GAISOU)
            Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(CInt(Text1(ptxG_SEQNO).Text), "000"))
        
            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, "")
            
            Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, "")
        
            Call UniCode_Conv(P_COMPO_K_REC.FILLER, "")
        
        
        End If
    
    
        Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, Mid(txtG_KEY.Text, 1, 1))                        '子　事業部
        Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, Mid(txtG_KEY.Text, 2, 1))                        '子　国内外
        Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, Text1(ptxG_HIN_GAI).Text)                       '子　品番
        Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(CDbl(Text1(ptxG_KO_QTY).Text), "000.00"))    '子　員数
    
    
    
        Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, "")                                              '更新担当者ｺｰﾄﾞ
                                                                                                    '更新日時
        Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
        Do
            
            DoEvents
            
            sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Update_Proc = False
                        Exit Do
                    End If
                Case Else
                    Call File_Error(sts, com, "構成マスタ")
                    Exit Function
            End Select
        
        Loop
    
    
    
    
    
    End If
    
    '--------------------------------------------   同梱・構成
    If Option1(poptD_ITEM).Value Then
    
        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_DOUKON)
        Call UniCode_Conv(K0_P_COMPO.SEQNO, Text1(ptxD_SEQNO).Text)
    
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound
                    com = BtOpInsert
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Update_Proc = False
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "構成マスタ")
                    Exit Function
            
            End Select
    
    
        Loop
    
    
        If com = BtOpInsert Then
            Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
            Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
            Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
            Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, Text1(ptxHIN_GAI).Text)
            Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_DOUKON)
            Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(CInt(Text1(ptxD_SEQNO).Text), "000"))
        
        
            Call UniCode_Conv(P_COMPO_K_REC.FILLER, "")
        
        
        End If
    
    
        Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, Mid(txtD_KEY.Text, 1, 1))                        '子　事業部
        Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, Mid(txtD_KEY.Text, 2, 1))                        '子　国内外
        Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, Text1(ptxD_HIN_GAI).Text)                       '子　品番
        
        Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, Right(Combo1(pcmbD_SYUBETSU).Text, 2))
        
        
        Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(CDbl(Text1(ptxD_KO_QTY).Text), "000.00"))    '子　員数
    
        Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, Text1(ptxD_BIKOU).Text)                           '子　備考
    
    
    
        Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, "")                                              '更新担当者ｺｰﾄﾞ
                                                                                                    '更新日時
        Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
        Do
            
            DoEvents
            
            sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Update_Proc = False
                        Exit Do
                    End If
                Case Else
                    Call File_Error(sts, com, "構成マスタ")
                    Exit Function
            End Select
        
        Loop
    
    
    
    
    
    End If
    
    
    Update_Proc = False


End Function
Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   構成マスタ削除（行単位）
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer

    Delete_Proc = True
    
    '--------------------------------------------   個装資材
    If Option1(poptK_ITEM).Value Then
    
        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_KOSOU)
        Call UniCode_Conv(K0_P_COMPO.SEQNO, Text1(ptxK_SEQNO).Text)
    
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound
                    Delete_Proc = False
                    Exit Function
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Delete_Proc = False
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "構成マスタ")
                    Exit Function
            
            End Select
    
        Loop
    
    
        Do
            
            DoEvents
            
            sts = BTRV(BtOpDelete, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Delete_Proc = False
                        Exit Do
                    End If
                Case Else
                    Call File_Error(sts, BtOpDelete, "構成マスタ")
                    Exit Function
            End Select
        Loop

    End If

    '--------------------------------------------   外装資材
    If Option1(poptG_ITEM).Value Then
    
        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_GAISOU)
        Call UniCode_Conv(K0_P_COMPO.SEQNO, Text1(ptxG_SEQNO).Text)
    
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound
                    Delete_Proc = False
                    Exit Function
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Delete_Proc = False
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "構成マスタ")
                    Exit Function
            
            End Select
    
    
        Loop
    
    
        Do
            
            DoEvents
            
            sts = BTRV(BtOpDelete, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Delete_Proc = False
                        Exit Do
                    End If
                Case Else
                    Call File_Error(sts, BtOpDelete, "構成マスタ")
                    Exit Function
            End Select
        Loop

    End If

    '--------------------------------------------   同梱・構成
    If Option1(poptD_ITEM).Value Then
    
        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_DOUKON)
        Call UniCode_Conv(K0_P_COMPO.SEQNO, Text1(ptxD_SEQNO).Text)
    
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound
                    Delete_Proc = False
                    Exit Function
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Delete_Proc = False
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "構成マスタ")
                    Exit Function
            
            End Select
    
    
        Loop
    
    
        Do
            
            DoEvents
            
            sts = BTRV(BtOpDelete, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Delete_Proc = False
                        Exit Do
                    End If
                Case Else
                    Call File_Error(sts, BtOpDelete, "構成マスタ")
                    Exit Function
            End Select
        Loop

    End If

    Delete_Proc = False


End Function



Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
    
    Call Tab_Ctrl(Shift)        '移動

End Sub


Private Sub Command1_Click(Index As Integer)

Dim ans     As Integer
Dim i       As Integer

    Select Case Index
        Case P_CMD_Upd                      '更新
            
            For i = ptxHIN_GAI To ptxD_BIKOU
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
            
            Beep
            ans = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Update_Proc() Then
                    PM000502.Visible = False
                    INIT_FLG = False
                End If
            Else
                Exit Sub
            End If
                                
            '個装資材処理だったら
            If Option1(poptK_ITEM).Value Then
                For i = ptxK_SEQNO To ptxK_KO_QTY
                    Text1(i).Text = ""
                Next i
        
                If K_List_Disp_Proc() Then
                    PM000502.Visible = False
                    INIT_FLG = False
                End If
            Else
                '外装資材処理だったら
                If Option1(poptG_ITEM).Value Then
                    For i = ptxG_SEQNO To ptxG_KO_QTY
                        Text1(i).Text = ""
                    Next i
            
                    If G_List_Disp_Proc() Then
                        PM000502.Visible = False
                        INIT_FLG = False
                    End If
                Else
                    '同梱／構成処理だったら
                    If Option1(poptD_ITEM).Value Then
                        For i = ptxD_SEQNO To ptxD_BIKOU
                            Text1(i).Text = ""
                        Next i
                
                        Combo1(pcmbD_SYUBETSU).ListIndex = 0
                
                
                
                        If D_List_Disp_Proc() Then
                            PM000502.Visible = False
                            INIT_FLG = False
                        End If
                    
                    Else
                        'ﾍｯﾀﾞｰ対応
                        PM000502.Visible = False
                        INIT_FLG = False
                    End If
                End If
            End If
        
        Case P_CMD_DEL                      '削除
            ans = MsgBox("削除しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Delete_Proc() Then
                    PM000502.Visible = False
                    INIT_FLG = False
                End If
            Else
                Exit Sub
            End If
        
            '個装資材処理だったら
            If Option1(poptK_ITEM).Value Then
                For i = ptxK_SEQNO To ptxK_KO_QTY
                    Text1(i).Text = ""
                Next i
        
                If K_List_Disp_Proc() Then
                    PM000502.Visible = False
                    INIT_FLG = False
                End If
            End If
            '外装資材処理だったら
            If Option1(poptG_ITEM).Value Then
                For i = ptxG_SEQNO To ptxG_KO_QTY
                    Text1(i).Text = ""
                Next i
        
                If G_List_Disp_Proc() Then
                    PM000502.Visible = False
                    INIT_FLG = False
                End If
            End If
            '同梱／構成処理だったら
            If Option1(poptD_ITEM).Value Then
                For i = ptxD_SEQNO To ptxD_KO_QTY
                    Text1(i).Text = ""
                Next i
        
                If D_List_Disp_Proc() Then
                    PM000502.Visible = False
                    INIT_FLG = False
                End If
            
            End If
        
        
        
        Case P_CMD_DSP                      '検索/表示
        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        Case P_CMD_PRT                      '印刷
        
        Case P_CMD_End                      '終了
            PM000502.Visible = False
            INIT_FLG = False
    End Select

End Sub

Private Sub Command2_Click(Index As Integer)

Dim ans         As Integer
Dim K_Err_Mode  As Integer
Dim G_Err_Mode  As Integer
Dim D_Err_Mode  As Integer

Dim Messeg      As String


    Select Case Index
        Case pcmbALLDEL     '一括削除
        
            ans = MsgBox("[" & Trim(Text1(ptxHIN_GAI).Text) & "] の一括削除しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
            
                If ALLDEL_Proc() Then
                    Unload Me
                End If
            
                MsgBox "一括削除が終了しました。"
            
                PM000502.Visible = False
                INIT_FLG = False
            
            End If
                    
        
        
        Case pcmbALLCHK     '一括ﾁｪｯｸ
            
            ans = MsgBox("[" & Trim(Text1(ptxHIN_GAI).Text) & "] の一括ﾁｪｯｸしますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
            
                If ALLCHK_Proc(K_Err_Mode, G_Err_Mode, D_Err_Mode) Then
                    Unload Me
                End If
            
                 
                If K_Err_Mode = 0 And G_Err_Mode = 0 And D_Err_Mode = 0 Then
                
                    MsgBox "一括ﾁｪｯｸは、正常終了しました。"
                                
                    Command2(pcmbALLCHK).SetFocus
            
                Else
            
                    Messeg = ""
                    If K_Err_Mode = 1 Then
                        Messeg = "「個装資材」"
                    End If
                    If G_Err_Mode = 1 Then
                        Messeg = "「外装資材」"
                    End If
                    If D_Err_Mode = 1 Then
                        Messeg = "「同梱／構成」"
                    End If
                    
                    MsgBox Messeg & "に未登録品番があります。"
                    
                    If K_Err_Mode = 1 Then
            
                        Option1(poptK_ITEM).Value = True
                        Option1(poptG_ITEM).Value = False
                        Option1(poptD_ITEM).Value = False
            
                        List1(plstK_ITEM).SetFocus
                        If List1(plstK_ITEM).ListCount > 0 Then
                            List1(plstK_ITEM).ListIndex = 0
                        Else
                            Text1(ptxK_HIN_GAI).SetFocus
                        End If
            
                    Else
                        If G_Err_Mode = 1 Then
                
                            Option1(poptK_ITEM).Value = False
                            Option1(poptG_ITEM).Value = True
                            Option1(poptD_ITEM).Value = False
                
                            List1(plstG_ITEM).SetFocus
                            If List1(plstG_ITEM).ListCount > 0 Then
                                List1(plstG_ITEM).ListIndex = 0
                            Else
                                Text1(ptxG_HIN_GAI).SetFocus
                            End If
                
                        
                        Else
                            Option1(poptK_ITEM).Value = False
                            Option1(poptG_ITEM).Value = False
                            Option1(poptD_ITEM).Value = True
                
                            List1(plstD_ITEM).SetFocus
                            If List1(plstD_ITEM).ListCount > 0 Then
                                List1(plstD_ITEM).ListIndex = 0
                            Else
                                Text1(ptxD_HIN_GAI).SetFocus
                            End If
                        
                        End If
                    End If
                End If
            End If
        
        
        
        Case pcmbRENUM      '追番ﾘﾅﾝﾊﾞｰ
            ans = MsgBox("[" & Trim(Text1(ptxHIN_GAI).Text) & "] のﾘﾅﾝﾊﾞｰしますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
            
                If ALLRENUM_Proc() Then
                    Unload Me
                End If
            
                 
                
                MsgBox "ﾘﾅﾝﾊﾞｰは、正常終了しました。"
                            
                Command2(pcmbALLCHK).SetFocus
            
            
            
            End If
    End Select


End Sub

Private Sub Form_Activate()
    
Dim i       As Integer
Dim CODE    As String
    
    If INIT_FLG Then
        Exit Sub
    End If


    Select Case G_SCREEN_FLG
        Case G_SCREEN_INS       '新規
                
            Combo1(pcmbSHIMUKE).BackColor = G_INPUT_OK
            Combo1(pcmbSHIMUKE).TabStop = True
            Combo1(pcmbSHIMUKE).Locked = False
                
                
            Text1(ptxHIN_GAI).BackColor = G_INPUT_OK
            Text1(ptxHIN_GAI).TabStop = True
            Text1(ptxHIN_GAI).Locked = False
                
            For i = ptxHIN_GAI To ptxD_BIKOU
                Text1(i).Text = ""
            Next i
                
                
            For i = pcmbSHIMUKE To pcmbD_SYUBETSU
                
                
                If Combo1(i).ListCount > 0 Then
                    Combo1(i).ListIndex = 0
                End If
            Next i
                
            For i = plstK_ITEM To plstD_ITEM
                List1(i).Clear
            Next i
                
            For i = poptK_ITEM To poptD_ITEM
                
                Option1(i).Value = False
            
            Next i
                
                
            Combo1(pcmbSHIMUKE).SetFocus
            Combo1(pcmbSHIMUKE).ListIndex = 0
                
                
                
        
        Case G_SCREEN_UPD       '更新
    
            Combo1(pcmbSHIMUKE).BackColor = G_INPUT_NG
            Combo1(pcmbSHIMUKE).TabStop = False
            Combo1(pcmbSHIMUKE).Locked = True
                
    
    
            Text1(ptxHIN_GAI).BackColor = G_INPUT_NG
            Text1(ptxHIN_GAI).TabStop = False
            Text1(ptxHIN_GAI).Locked = True
        
            '2019.05.28 これが無いと、前回画面内容が残る！  高沢
            For i = ptxHIN_GAI To ptxD_BIKOU
                Text1(i).Text = ""
            Next i
            
            
            
            CODE = PM000501.txSEL_KEY.Text
            
            If Item_Disp_Proc(CODE) Then
                Exit Sub
            End If
    
            Text1(ptxCLASS_CODE).SetFocus
    
    End Select


    INIT_FLG = True

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

    
    '仕向け先名のセット
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD, 0) Then
        Unload Me
    End If
    
    '種別名のセット
    If Code_Set_Proc(pcmbD_SYUBETSU, P_KBN06_CD, 1) Then
        Unload Me
    End If
    
    
    
    
    INIT_FLG = False
    
    
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            
                                            
                                            
                                            '構成マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "構成マスタ")
        End If
    End If
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
                                            'クラスマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "クラスマスタ")
        End If
    End If
    
    sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PM000501 = Nothing
    Set PM000502 = Nothing

    End
End Sub



Private Sub List1_DblClick(Index As Integer)
    
    
    
    Select Case Index
        Case plstK_ITEM     '個装資材
        
            Option1(poptK_ITEM).Value = True
            Option1(poptG_ITEM).Value = False
            Option1(poptD_ITEM).Value = False
        
        
    
            If K_Item_Disp_Proc(Left(List1(Index).List(List1(Index).ListIndex), 3)) Then
                PM000502.Visible = False
                INIT_FLG = False
            End If
                
            If txtK_KEY.Text = "" Then
                If List1(Index).ListCount > 0 Then
                    List1(Index).SetFocus
                    List1(Index).ListIndex = 0
                Else
                    Text1(ptxK_SEQNO).SetFocus
                End If
            Else
                Text1(ptxK_SEQNO).SetFocus
            End If
        
        Case plstG_ITEM     '外装資材
        
            Option1(poptK_ITEM).Value = False
            Option1(poptG_ITEM).Value = True
            Option1(poptD_ITEM).Value = False
        
    
            If G_Item_Disp_Proc(Left(List1(Index).List(List1(Index).ListIndex), 3)) Then
                PM000502.Visible = False
                INIT_FLG = False
            End If
                
            If txtK_KEY.Text = "" Then
                If List1(Index).ListCount > 0 Then
                    List1(Index).SetFocus
                    List1(Index).ListIndex = 0
                Else
                    Text1(ptxG_SEQNO).SetFocus
                End If
            Else
                Text1(ptxG_SEQNO).SetFocus
            End If
    
        Case plstD_ITEM     '同梱／構成
        
            Option1(poptK_ITEM).Value = False
            Option1(poptG_ITEM).Value = False
            Option1(poptD_ITEM).Value = True
        
        
    
            If D_Item_Disp_Proc(Left(List1(Index).List(List1(Index).ListIndex), 3)) Then
                PM000502.Visible = False
                INIT_FLG = False
            End If
                
            If txtK_KEY.Text = "" Then
                If List1(Index).ListCount > 0 Then
                    List1(Index).SetFocus
                    List1(Index).ListIndex = 0
                Else
                    Text1(ptxD_SEQNO).SetFocus
                End If
            Else
                Text1(ptxD_SEQNO).SetFocus
            End If
    
    
    End Select

End Sub

Private Sub List1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    If Shift = vbShiftMask Then
        Call Tab_Ctrl(Shift)        '移動
    Else
        Select Case Index
            Case plstK_ITEM     '個装資材
            
                Option1(poptK_ITEM).Value = True
                Option1(poptG_ITEM).Value = False
                Option1(poptD_ITEM).Value = False
            
            
        
                If K_Item_Disp_Proc(Right(List1(Index).List(List1(Index).ListIndex), 3)) Then
                    PM000502.Visible = False
                    INIT_FLG = False
                End If
                    
                If txtK_KEY.Text = "" Then
                    If List1(Index).ListCount > 0 Then
                        List1(Index).SetFocus
                        List1(Index).ListIndex = 0
                    Else
                        Text1(ptxK_SEQNO).SetFocus
                    End If
                Else
                    Text1(ptxK_SEQNO).SetFocus
                End If
            
            Case plstG_ITEM     '外装資材
            
                Option1(poptK_ITEM).Value = False
                Option1(poptG_ITEM).Value = True
                Option1(poptD_ITEM).Value = False
            
            
        
                If G_Item_Disp_Proc(Right(List1(Index).List(List1(Index).ListIndex), 3)) Then
                    PM000502.Visible = False
                    INIT_FLG = False
                End If
                    
                If txtK_KEY.Text = "" Then
                    If List1(Index).ListCount > 0 Then
                        List1(Index).SetFocus
                        List1(Index).ListIndex = 0
                    Else
                        Text1(ptxG_SEQNO).SetFocus
                    End If
                Else
                    Text1(ptxG_SEQNO).SetFocus
                End If
        
            Case plstD_ITEM     '同梱／構成
            
                Option1(poptK_ITEM).Value = False
                Option1(poptG_ITEM).Value = False
                Option1(poptD_ITEM).Value = True
            
            
        
                If G_Item_Disp_Proc(Right(List1(Index).List(List1(Index).ListIndex), 3)) Then
                    PM000502.Visible = False
                    INIT_FLG = False
                End If
                    
                If txtK_KEY.Text = "" Then
                    If List1(Index).ListCount > 0 Then
                        List1(Index).SetFocus
                        List1(Index).ListIndex = 0
                    Else
                        Text1(ptxD_SEQNO).SetFocus
                    End If
                Else
                    Text1(ptxD_SEQNO).SetFocus
                End If
        
        
        End Select
    End If

End Sub

Private Sub Option1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
        
        
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

    Select Case Index
        Case ptxCLASS_CODE      'ｸﾗｽｺｰﾄﾞ
            
            Option1(poptK_ITEM).Value = False
            Option1(poptG_ITEM).Value = False
            Option1(poptD_ITEM).Value = False
    
    End Select


End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    If Error_Check_Proc(Index) Then     'エラーチェック
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Function Code_Set_Proc(Index As Integer, KBN As String, Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   コードマスタをコンボにセットする。
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim Key_Len     As Integer
Dim Option1     As Integer
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
    
    If Mode = 1 Then
        Combo1(Index).AddItem Space(Key_Len)
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
            wkOption = Trim(StrConv(P_CODEREC.Option1, vbUnicode))
        End If
        If P_KBN_TBL(i).KBN_OP2 Then
            wkOption = wkOption & Trim(StrConv(P_CODEREC.OPTION2, vbUnicode))
        End If
        
        
        
        Combo1(Index).AddItem StrConv(P_CODEREC.C_RNAME, vbUnicode) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), Key_Len) & wkOption
        
        
        com = BtOpGetNext
    
    Loop

    Code_Set_Proc = False
    



End Function

Private Function K_List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   個装資材情報をﾘｽﾄﾎﾞｯｸに表示する
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim com     As Integer

Dim KO_QTY  As String * 6

        K_List_Disp_Proc = True


        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                    
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_KOSOU)
        Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
            
        com = BtOpGetGreater
            
        List1(plstK_ITEM).Clear
            
            
        Do
            
            DoEvents
            
            sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    
                    If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                        Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHIN_GAI).Text) Or _
                        StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_KOSOU Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, BtOpGetGreater, "構成マスタ")
                    Exit Function
            End Select


            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))


            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    
                Case BtErrKeyNotFound
                    Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番です。")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
            
            KO_QTY = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
            KO_QTY = Space(Len(KO_QTY) - Len(Trim(KO_QTY))) & Trim(KO_QTY)

            List1(plstK_ITEM).AddItem StrConv(P_COMPO_K_REC.SEQNO, vbUnicode) & "  " & _
                                        StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode) & "  " & _
                                        StrConv(ITEMREC.HIN_NAME, vbUnicode) & " " & _
                                        KO_QTY & "          " & _
                                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) & _
                                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode)
            com = BtOpGetNext
        
        Loop

        K_List_Disp_Proc = False

End Function

Private Function G_List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   外装資材情報をﾘｽﾄﾎﾞｯｸに表示する
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim com     As Integer

Dim KO_QTY  As String * 6

        G_List_Disp_Proc = True


        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                    
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_GAISOU)
        Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
            
        com = BtOpGetGreater
            
        List1(plstG_ITEM).Clear
            
            
        Do
            
            DoEvents
            
            sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    
                    If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                        Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHIN_GAI).Text) Or _
                        StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_GAISOU Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, BtOpGetGreater, "構成マスタ")
                    Exit Function
            End Select


            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))


            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    
                Case BtErrKeyNotFound
                    Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番です。")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
            
            KO_QTY = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
            KO_QTY = Space(Len(KO_QTY) - Len(Trim(KO_QTY))) & Trim(KO_QTY)

            List1(plstG_ITEM).AddItem StrConv(P_COMPO_K_REC.SEQNO, vbUnicode) & "  " & _
                                        StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode) & "  " & _
                                        StrConv(ITEMREC.HIN_NAME, vbUnicode) & " " & _
                                        KO_QTY & "          " & _
                                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) & _
                                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode)
            com = BtOpGetNext
        
        Loop

        G_List_Disp_Proc = False

End Function


Private Function D_List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   同梱／構成情報をﾘｽﾄﾎﾞｯｸに表示する
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim KO_QTY      As String * 6

Dim SYUBETSU    As String * 4

        D_List_Disp_Proc = True


        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                    
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_DOUKON)
        Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
            
        com = BtOpGetGreater
            
        List1(plstD_ITEM).Clear
            
            
        Do
            
            DoEvents
            
            sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    
                    If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                        Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHIN_GAI).Text) Or _
                        StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_DOUKON Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, BtOpGetGreater, "構成マスタ")
                    Exit Function
            End Select


            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))


            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    
                Case BtErrKeyNotFound
                    Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番です。")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
            
            
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                    
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_CODEREC.C_RNAME, "　　　")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "コードマスタ")
                    Exit Function
            End Select
            SYUBETSU = StrConv(P_CODEREC.C_RNAME, vbUnicode)
            
            
            
            KO_QTY = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
            KO_QTY = Space(Len(KO_QTY) - Len(Trim(KO_QTY))) & Trim(KO_QTY)

            List1(plstD_ITEM).AddItem StrConv(P_COMPO_K_REC.SEQNO, vbUnicode) & "  " & _
                                        SYUBETSU & "    " & _
                                        StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode) & "  " & _
                                        Left(StrConv(ITEMREC.HIN_NAME, vbUnicode), 30) & " " & _
                                        KO_QTY & "  " & _
                                        StrConv(P_COMPO_K_REC.KO_BIKOU, vbUnicode) & "   " & _
                                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) & _
                                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode)
            com = BtOpGetNext
        
        Loop

        D_List_Disp_Proc = False

End Function
Private Function ALLDEL_Proc() As Integer
'----------------------------------------------------------------------------
'                   指定品番の子部品を全て削除する
'----------------------------------------------------------------------------
Dim sts As Integer
Dim com As Integer
Dim ans As Integer


    ALLDEL_Proc = True
    
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "")
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
    
    com = BtOpGetGreaterEqual
    
    
    
    Do
        DoEvents
        
        
        Do
            sts = BTRV(com + BtSNoWait, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    
                    If StrConv(P_COMPO_O_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2) Or _
                        StrConv(P_COMPO_O_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1) Or _
                        StrConv(P_COMPO_O_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1) Or _
                        Trim(StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHIN_GAI).Text) Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        
                        ALLDEL_Proc = False
                        Exit Function
                    
                    End If
                    
                    Exit Do
                Case BtErrEOF
                    ALLDEL_Proc = False
                    Exit Function
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        ALLDEL_Proc = False
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "構成マスタ")
                    Exit Function
            
            End Select
        Loop
    
    
        Do
            
            DoEvents
            
            sts = BTRV(BtOpDelete, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        ALLDEL_Proc = False
                        Exit Do
                    End If
                Case Else
                    Call File_Error(sts, BtOpDelete, "構成マスタ")
                    Exit Function
            End Select
        Loop

        com = BtOpGetNext
    
    Loop

    ALLDEL_Proc = False


End Function

Private Function ALLCHK_Proc(K_Err_Mode As Integer, G_Err_Mode As Integer, D_Err_Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   指定品番の子部品を全てﾁｪｯｸする
'   ﾁｪｯｸ項目　：　品目マスタの有無
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim com     As Integer

    ALLCHK_Proc = True

    K_Err_Mode = 0
    G_Err_Mode = 0
    D_Err_Mode = 0
    '----------------------------   品目マスタの有無チェック（全構成）
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_KOSOU)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
    
    com = BtOpGetGreaterEqual
    
        
    Do
        
        DoEvents
        
        sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
                
                If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2) Or _
                    StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1) Or _
                    StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1) Or _
                    Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHIN_GAI).Text) Then
                    
                    Exit Do
                
                End If
                
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "構成マスタ")
                Exit Function
        
        End Select
    
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
    
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                
                
            Case BtErrKeyNotFound
            
                Select Case StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode)
                
                    Case P_KOSOU            '個装資材
                    
                        K_Err_Mode = 1
                    
                    Case P_GAISOU           '外装資材
                        
                        G_Err_Mode = 1
                    
                    Case P_DOUKON           '同梱・構成
                    
                        D_Err_Mode = 1
                    
                End Select
            
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        
        End Select
    
    
        com = BtOpGetNext
    
    Loop



    ALLCHK_Proc = False


End Function

Private Function K_Item_Disp_Proc(Item_Key As String) As String
'----------------------------------------------------------------------------
'                   個装資材の指定情報の表示
'----------------------------------------------------------------------------
Dim sts As Integer

    
    K_Item_Disp_Proc = True

    txtK_KEY.Text = ""

    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_KOSOU)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, Item_Key)

    sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    Select Case sts
        Case BtNoErr
            
            
        Case BtErrKeyNotFound
            MsgBox "データが変更されています。「ＯＫ」で再表示を行います。"
            If K_List_Disp_Proc() Then
                Exit Function
            End If
            K_Item_Disp_Proc = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "構成マスタ")
            Exit Function
    
    End Select

    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            
            
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番です。")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    
    End Select

    Text1(ptxK_SEQNO).Text = StrConv(P_COMPO_K_REC.SEQNO, vbUnicode)                            '追番
    Text1(ptxK_HIN_GAI).Text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)                     '子品番
    Text1(ptxK_HIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)                            '子品名
    Text1(ptxK_KO_QTY).Text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")   '員数

    txtK_KEY.Text = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode) & StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)


    K_Item_Disp_Proc = False

End Function

Private Function G_Item_Disp_Proc(Item_Key As String) As String
'----------------------------------------------------------------------------
'                   外装資材の指定情報の表示
'----------------------------------------------------------------------------
Dim sts As Integer

    
    G_Item_Disp_Proc = True

    txtK_KEY.Text = ""

    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_GAISOU)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, Item_Key)

    sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    Select Case sts
        Case BtNoErr
            
            
        Case BtErrKeyNotFound
            MsgBox "データが変更されています。「ＯＫ」で再表示を行います。"
            If G_List_Disp_Proc() Then
                Exit Function
            End If
            G_Item_Disp_Proc = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "構成マスタ")
            Exit Function
    
    End Select

    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            
            
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番です。")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    
    End Select

    Text1(ptxG_SEQNO).Text = StrConv(P_COMPO_K_REC.SEQNO, vbUnicode)                            '追番
    Text1(ptxG_HIN_GAI).Text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)                     '子品番
    Text1(ptxG_HIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)                            '子品名
    Text1(ptxG_KO_QTY).Text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")   '員数

    txtG_KEY.Text = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode) & StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)


    G_Item_Disp_Proc = False

End Function


Private Function D_Item_Disp_Proc(Item_Key As String) As String
'----------------------------------------------------------------------------
'                   同梱／構成の指定情報の表示
'----------------------------------------------------------------------------
Dim sts As Integer
Dim i   As Integer
    
    D_Item_Disp_Proc = True

    txtD_KEY.Text = ""

    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_DOUKON)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, Item_Key)

    sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    Select Case sts
        Case BtNoErr
            
            
        Case BtErrKeyNotFound
            MsgBox "データが変更されています。「ＯＫ」で再表示を行います。"
            If D_List_Disp_Proc() Then
                Exit Function
            End If
            D_Item_Disp_Proc = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "構成マスタ")
            Exit Function
    
    End Select

    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            
            
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番です。")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    
    End Select

    Text1(ptxD_SEQNO).Text = StrConv(P_COMPO_K_REC.SEQNO, vbUnicode)                        '追番
    
    For i = 0 To Combo1(pcmbD_SYUBETSU).ListCount - 1                                       '種別
        If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = Right(Combo1(pcmbD_SYUBETSU).List(i), P_KBN06_Len) Then
            Combo1(pcmbD_SYUBETSU).ListIndex = i
            Exit For
        End If
    Next i
    Text1(ptxD_HIN_GAI).Text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)                     '子品番
    Text1(ptxD_HIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)                            '子品名
    Text1(ptxD_KO_QTY).Text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")   '員数
    Text1(ptxD_BIKOU).Text = StrConv(P_COMPO_K_REC.KO_BIKOU, vbUnicode)                         '備考


    txtG_KEY.Text = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode) & StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)


    D_Item_Disp_Proc = False

End Function
Private Function ALLRENUM_Proc() As Integer
'----------------------------------------------------------------------------
'                   子部品の追番を振り直す（１０毎）
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim com     As Integer
Dim SEQNO   As Integer
Dim ans     As Integer


    ALLRENUM_Proc = True
    
    
    '-------------------------------------  個装資材の処理
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_KOSOU)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
    
    com = BtOpGetGreaterEqual
    
    SEQNO = 0
    
    Do
        DoEvents
        
        
        Do
            sts = BTRV(com + BtSNoWait, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    
                    If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2) Or _
                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1) Or _
                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1) Or _
                        Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHIN_GAI).Text) Or _
                        StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_KOSOU Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                    
                        sts = BtErrEOF
                        Exit Do
                    
                    End If
                    
                    Exit Do
                Case BtErrEOF
                
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        ALLRENUM_Proc = False
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "構成マスタ")
                    Exit Function
            
            End Select
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
        
        SEQNO = SEQNO + 10
        
        Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(SEQNO, "000"))
        
        
        Do
            
            DoEvents
            
            sts = BTRV(BtOpUpdate, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        ALLRENUM_Proc = False
                        Exit Do
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "構成マスタ")
                    Exit Function
            End Select
        Loop

        com = BtOpGetNext
    
    Loop
    '-------------------------------------  外装資材の処理
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_GAISOU)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
    
    com = BtOpGetGreaterEqual
    
    SEQNO = 0
    
    Do
        DoEvents
        
        
        Do
            sts = BTRV(com + BtSNoWait, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    
                    If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2) Or _
                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1) Or _
                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1) Or _
                        Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHIN_GAI).Text) Or _
                        StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_GAISOU Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                    
                        sts = BtErrEOF
                        Exit Do
                    
                    End If
                    
                    Exit Do
                Case BtErrEOF
                
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        ALLRENUM_Proc = False
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "構成マスタ")
                    Exit Function
            
            End Select
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
        
        SEQNO = SEQNO + 10
        
        Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(SEQNO, "000"))
        
        
        Do
            
            DoEvents
            
            sts = BTRV(BtOpUpdate, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        ALLRENUM_Proc = False
                        Exit Do
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "構成マスタ")
                    Exit Function
            End Select
        Loop

        com = BtOpGetNext
    
    Loop

    '-------------------------------------  同梱／構成の処理
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).Text)
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_DOUKON)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
    
    com = BtOpGetGreaterEqual
    
    SEQNO = 0
    
    Do
        DoEvents
        
        
        Do
            sts = BTRV(com + BtSNoWait, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    
                    If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE).Text, 5), 1, 3) Or _
                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE).Text, 5), 4, 1) Or _
                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE).Text, 5), 5, 1) Or _
                        Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHIN_GAI).Text) Or _
                        StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_DOUKON Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                    
                        sts = BtErrEOF
                        Exit Do
                    
                    End If
                    
                    Exit Do
                Case BtErrEOF
                
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        ALLRENUM_Proc = False
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "構成マスタ")
                    Exit Function
            
            End Select
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
        
        SEQNO = SEQNO + 10
        
        Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(SEQNO, "000"))
        
        
        Do
            
            DoEvents
            
            sts = BTRV(BtOpUpdate, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        ALLRENUM_Proc = False
                        Exit Do
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "構成マスタ")
                    Exit Function
            End Select
        Loop

        com = BtOpGetNext
    
    Loop

    ALLRENUM_Proc = False


End Function
