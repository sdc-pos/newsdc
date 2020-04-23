VERSION 5.00
Begin VB.Form F1200351 
   BackColor       =   &H00FFFFFF&
   Caption         =   "標準棚番集計処理"
   ClientHeight    =   6375
   ClientLeft      =   2325
   ClientTop       =   2910
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
   ScaleHeight     =   6375
   ScaleWidth      =   11295
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Index           =   13
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1332
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Index           =   12
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1332
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Index           =   11
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1332
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Index           =   8
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1332
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Index           =   14
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5040
      Width           =   852
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Index           =   15
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5040
      Width           =   492
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Index           =   16
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5040
      Width           =   492
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Index           =   10
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1332
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Index           =   9
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1332
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   7
      Left            =   5760
      MaxLength       =   2
      TabIndex        =   7
      Top             =   480
      Width           =   372
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   6
      Left            =   5280
      MaxLength       =   2
      TabIndex        =   6
      Top             =   480
      Width           =   372
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   5
      Left            =   4800
      MaxLength       =   2
      TabIndex        =   5
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   4
      Left            =   4320
      MaxLength       =   2
      TabIndex        =   4
      Top             =   480
      Width           =   372
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   3
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   3
      Top             =   480
      Width           =   372
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   2
      Left            =   3255
      MaxLength       =   2
      TabIndex        =   2
      Top             =   480
      Width           =   372
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   1
      Left            =   2760
      MaxLength       =   2
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   0
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   0
      Top             =   480
      Width           =   372
   End
   Begin VB.CommandButton Command 
      Caption         =   "終　了"
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
      TabIndex        =   25
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
      Index           =   9
      Left            =   8640
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "品削除"
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
      TabIndex        =   22
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
      Index           =   6
      Left            =   5640
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
      Index           =   5
      Left            =   4800
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
      Index           =   4
      Left            =   3960
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
      Index           =   3
      Left            =   2640
      TabIndex        =   17
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
      Index           =   1
      Left            =   960
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "実　行"
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "ケ月間入庫なし"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   4200
      TabIndex        =   50
      Top             =   3360
      Width           =   2415
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTuki 
      BackColor       =   &H80000009&
      Caption         =   "ｎ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   49
      Top             =   3360
      Width           =   375
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "在庫無し"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   4200
      TabIndex        =   48
      Top             =   2880
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "総品目数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   2280
      TabIndex        =   47
      Top             =   3600
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "―"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   3840
      TabIndex        =   46
      Top             =   4080
      Width           =   255
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "＝"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   5880
      TabIndex        =   45
      Top             =   4080
      Width           =   255
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "在庫無し"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   4200
      TabIndex        =   41
      Top             =   1680
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "年"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   12
      Left            =   6720
      TabIndex        =   40
      Top             =   5160
      Width           =   372
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "月"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   13
      Left            =   7800
      TabIndex        =   39
      Top             =   5160
      Width           =   372
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "日現在"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   14
      Left            =   8880
      TabIndex        =   38
      Top             =   5160
      Width           =   1092
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "＝"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   5760
      TabIndex        =   37
      Top             =   2160
      Width           =   255
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "―"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3720
      TabIndex        =   36
      Top             =   2160
      Width           =   255
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "-"
      Height          =   252
      Index           =   8
      Left            =   5640
      TabIndex        =   35
      Top             =   600
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "-"
      Height          =   252
      Index           =   5
      Left            =   5160
      TabIndex        =   34
      Top             =   600
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "-"
      Height          =   252
      Index           =   4
      Left            =   4680
      TabIndex        =   33
      Top             =   600
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "-"
      Height          =   252
      Index           =   7
      Left            =   3600
      TabIndex        =   32
      Top             =   600
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "-"
      Height          =   252
      Index           =   2
      Left            =   3120
      TabIndex        =   31
      Top             =   600
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "-"
      Height          =   252
      Index           =   1
      Left            =   2640
      TabIndex        =   30
      Top             =   600
      Width           =   252
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
      TabIndex        =   29
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "総品目数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   16.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   2160
      TabIndex        =   28
      Top             =   1680
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "～"
      Height          =   252
      Index           =   3
      Left            =   4080
      TabIndex        =   27
      Top             =   600
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H80000009&
      Caption         =   "対象棚番"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   26
      Top             =   600
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Checked         =   -1  'True
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1200351"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxS_Soko_No% = 0             '開始　倉庫№
Private Const ptxS_Retu% = 1                '開始　列
Private Const ptxS_Ren% = 2                 '開始　連
Private Const ptxS_Dan% = 3                 '開始　段
Private Const ptxE_Soko_No% = 4             '終了　倉庫№
Private Const ptxE_Retu% = 5                '終了　列
Private Const ptxE_Ren% = 6                 '終了　連
Private Const ptxE_Dan% = 7                 '終了　段

Private Const ptxAbove_ITEM_SU% = 8         '総品目数
Private Const ptxAbove_ZAIKO_NASI% = 9      '在庫なし数
Private Const ptxAbove_ZAIKO_ARI% = 10      '入荷あり数
Private Const ptxBelow_ITEM_SU% = 11        '総品目数
Private Const ptxBelow_NYUKA_NASI% = 12     '入荷なし数（在庫なし）
Private Const ptxBelow_NYUKA_ARI% = 13      '入荷あり数（在庫なし）


Private Const ptxNOW_DT_YY% = 14            '現在日付　年
Private Const ptxNOW_DT_MM% = 15            '現在日付　月
Private Const ptxNOW_DT_DD% = 16            '現在日付　日

Private Const Text_Max% = 16                '画面項目別最大ｲﾝﾃﾞｯｸｽ

Dim Tuki_Suu    As Integer                  '対象月数
Dim Taisyo_YMD  As String * 8               '対象月日

Private Function Main_Proc() As Integer
'----------------------------------------------------------------------------
'                   データ集計メイン処理
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer
Dim com_IDO     As Integer

Dim ALL_ITEM_SU As Long                     '総品目数
Dim Zaiko_ARI   As Long                     '在庫あり品目数
Dim NYUKA_NASI  As Long                     '在庫なし入荷なし
                                            
Dim Nyuka_Flg   As Boolean                  '入荷有無フラグ
                                            
                                            
                                            
    Main_Proc = True
                                            
    Call Input_Lock
                                            
                                            
    ALL_ITEM_SU = 0
    Zaiko_ARI = 0
    NYUKA_NASI = 0
                                        '品目マスタ読み込み開始
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, "")
    Call UniCode_Conv(K0_ITEM.HIN_GAI, "")
    
    com = BtOpGetGreater
    Do
        DoEvents
        
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                                        
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Exit Function
        End Select
        
        If Len(Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode))) = 0 Then
                                    '標準倉庫棚未設定
        Else
                        
            If (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)) _
                < ((Text(ptxS_Soko_No).Text & Text(ptxS_Retu).Text & Text(ptxS_Ren).Text & Text(ptxS_Dan).Text)) Or _
                    (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)) _
                > ((Text(ptxE_Soko_No).Text & Text(ptxE_Retu).Text & Text(ptxE_Ren).Text & Text(ptxE_Dan).Text)) Then
                                    '棚番範囲範囲外
            Else
                ALL_ITEM_SU = ALL_ITEM_SU + 1   '全品目数
                                            '在庫のチェック
                Call UniCode_Conv(K4_ZAIKO.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K4_ZAIKO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K4_ZAIKO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(K4_ZAIKO.Soko_No, "")
                Call UniCode_Conv(K4_ZAIKO.Retu, "")
                Call UniCode_Conv(K4_ZAIKO.Ren, "")
                Call UniCode_Conv(K4_ZAIKO.Dan, "")
            
                sts = BTRV(BtOpGetGreater, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K4_ZAIKO, Len(K4_ZAIKO), 4)
                Select Case sts
                    Case BtNoErr
                        If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                            StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> StrConv(ITEMREC.NAIGAI, vbUnicode) Or _
                            StrConv(ZAIKOREC.HIN_GAI, vbUnicode) <> StrConv(ITEMREC.HIN_GAI, vbUnicode) Then
                            '在庫なしなら入荷のチェック
                            Call UniCode_Conv(K1_IDO.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                            Call UniCode_Conv(K1_IDO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                            Call UniCode_Conv(K1_IDO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                            Call UniCode_Conv(K1_IDO.JITU_DT, Taisyo_YMD)
                            Call UniCode_Conv(K1_IDO.JITU_TM, "")
                        
                            Nyuka_Flg = False
                        
                            com_IDO = BtOpGetGreaterEqual
                            
                            Do
                                DoEvents
                                sts = BTRV(com_IDO, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
                                Select Case sts
                                    Case BtNoErr
                                        If StrConv(IDOREC.JGYOBU, vbUnicode) <> StrConv(ITEMREC.JGYOBU, vbUnicode) Or _
                                            StrConv(IDOREC.NAIGAI, vbUnicode) <> StrConv(ITEMREC.NAIGAI, vbUnicode) Or _
                                            Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                                            Exit Do
                                        End If
                                    
                                        If StrConv(IDOREC.RIRK_ID, vbUnicode) = YOIN_TU_NYUKA Then
                                            Nyuka_Flg = True
                                            Exit Do
                                        End If
                                    
                                    Case BtErrEOF
                                        Exit Do
                                    Case Else
                                        Call File_Error(sts, com_IDO, "在庫移動歴")
                                        Exit Function
                                End Select
                            
                                com_IDO = BtOpGetNext
                            
                            Loop
                            If Not Nyuka_Flg Then
                                NYUKA_NASI = NYUKA_NASI + 1
                            End If
                        Else
                            Zaiko_ARI = Zaiko_ARI + 1
                        End If
                    Case BtErrEOF
                    
                        '在庫なしなら入荷のチェック
                        Call UniCode_Conv(K1_IDO.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                        Call UniCode_Conv(K1_IDO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                        Call UniCode_Conv(K1_IDO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                        Call UniCode_Conv(K1_IDO.JITU_DT, Taisyo_YMD)
                        Call UniCode_Conv(K1_IDO.JITU_TM, "")
                        
                        Nyuka_Flg = False
                        
                        com_IDO = BtOpGetGreaterEqual
                            
                        Do
                            DoEvents
                            sts = BTRV(com_IDO, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
                            Select Case sts
                                Case BtNoErr
                                    If StrConv(IDOREC.JGYOBU, vbUnicode) <> StrConv(ITEMREC.JGYOBU, vbUnicode) Or _
                                        StrConv(IDOREC.NAIGAI, vbUnicode) <> StrConv(ITEMREC.NAIGAI, vbUnicode) Or _
                                        Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                                        Exit Do
                                    End If
                                    
                                    If StrConv(IDOREC.RIRK_ID, vbUnicode) = YOIN_TU_NYUKA Then
                                        Nyuka_Flg = True
                                        Exit Do
                                    End If
                                    
                                Case BtErrEOF
                                    Exit Do
                                Case Else
                                    Call File_Error(sts, com_IDO, "在庫移動歴")
                                    Exit Function
                            End Select
                            
                            com_IDO = BtOpGetNext
                        
                        Loop
                            
                        If Not Nyuka_Flg Then
                            NYUKA_NASI = NYUKA_NASI + 1
                        End If
                    
                    Case Else
                        Call File_Error(sts, BtOpGetGreater, "在庫データ")
                        Exit Function
                End Select
            
            End If
        End If
        
        com = BtOpGetNext
    
    Loop


    Text(ptxAbove_ITEM_SU).Text = Format(ALL_ITEM_SU, "#,##0")
    Text(ptxAbove_ZAIKO_NASI).Text = Format((ALL_ITEM_SU - Zaiko_ARI), "#,##0")
    Text(ptxAbove_ZAIKO_ARI).Text = Format(Zaiko_ARI, "#,##0")
    
    Text(ptxBelow_ITEM_SU).Text = Format(ALL_ITEM_SU, "#,##0")
    Text(ptxBelow_NYUKA_NASI).Text = Format(NYUKA_NASI, "#,##0")
    Text(ptxBelow_NYUKA_ARI).Text = Format((ALL_ITEM_SU - NYUKA_NASI), "#,##0")
    
    
    Text(ptxNOW_DT_YY).Text = Left(Format(Date, "yyyymmdd"), 4)
    Text(ptxNOW_DT_MM).Text = Mid(Format(Date, "yyyymmdd"), 5, 2)
    Text(ptxNOW_DT_DD).Text = Right(Format(Date, "yyyymmdd"), 2)

                                            
    Call Input_UnLock
    
    Main_Proc = False

End Function

Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        Case 0              '実行
            
            If Err_Chk() Then
                Exit Sub
            End If
            
            Beep
            ans = MsgBox("「標準棚番集計処理」実行しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Main_Proc() Then
                    Unload Me
                End If
'                Call Clear_Field
            End If
            Text(ptxS_Soko_No).SetFocus
                    
        Case 8              '品番削除
                
                
            If Err_Chk() Then
                Exit Sub
            End If
                
                
            ans = MsgBox("「品番削除処理」実行しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Delete_Proc() Then
                    Unload Me
                End If
'                Call Clear_Field
            End If
            Text(ptxS_Soko_No).SetFocus
        
        
        
        
        Case 11             '終了
            Unload Me
        Case Else
            Beep
    End Select
    

End Sub

Private Sub Form_Activate()
    
    Text(ptxS_Soko_No).SetFocus


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

Dim i           As Integer
Dim c           As String * 128
Dim sts         As Integer
        

    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = Trim(c)
                                '「通常入荷」要因取り込み
    If GetIni("YOIN", "YOIN_TU_NYUKA", "SYS", c) Then
        Beep
        MsgBox "「通常入荷」要因の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    YOIN_TU_NYUKA = Trim(c)
                                
                                '入庫対象月数
    If GetIni(App.EXEName, "TUKI", "SYS", c) Then
        Tuki_Suu = 1
    Else
        Tuki_Suu = CInt(Trim(c))
    End If
                                
    lblTuki.Caption = StrConv(Format(Tuki_Suu, "0"), vbWide)
                                
    Taisyo_YMD = Format(DateAdd("m", Tuki_Suu * (-1), Now), "YYYYMMDD")
                                
                                
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).Code = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).Code = Last_JGYOBU Then
            F1200351.Caption = "標準棚番集計処理（" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫移動歴ＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If

End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            '品目ﾏｽﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '在庫ＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫ﾃﾞｰﾀ")
        End If
    End If
                                            '在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1200351 = Nothing

    End
End Sub



Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    
                                    'メニューより終了要求
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If
    
    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).Code = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1200351.Caption = "標準棚番集計処理処理（" + RTrim(JGYOBU_T(Index).NAME) + "）"
    Last_JGYOBU = JGYOBU_T(Index).Code
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub
Private Function Err_Chk() As Integer
    
Dim i As Integer
    
    Err_Chk = True


    For i = ptxS_Soko_No To ptxE_Dan
        If Len(Text(i).Text) = 0 Then
            Select Case i
                Case ptxS_Soko_No
                Case ptxS_Retu, ptxS_Ren, ptxS_Dan
                    Text(i).Text = ""
                Case ptxE_Soko_No
                    Text(i).Text = "zz"
                Case ptxE_Retu, ptxE_Ren, ptxE_Dan
                    Text(i).Text = "99"
            End Select
        Else
            If i <> ptxS_Soko_No And i <> ptxS_Soko_No Then
                If IsNumeric(Text(i).Text) Then
                    Text(i).Text = Format(CInt(Text(i).Text), "00")
                End If
            End If
        End If
    Next i
    
    Err_Chk = False

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1200351.MousePointer = vbHourglass

    Call Ctrl_Lock(F1200351)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1200351)


    F1200351.MousePointer = vbDefault

End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If


End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
Dim i   As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    For i = Index + 1 To Text_Max
        If Text(i).Visible And Text(i).Enabled And Text(i).TabStop Then
            Text(i).SetFocus
            Exit For
        End If
    Next i

End Sub

Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   品目削除処理
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim com_IDO     As Integer
Dim ans         As Integer
                                            
Dim Nyuka_Flg   As Boolean                  '入荷有無フラグ
                                            
                                            
                                            
    Delete_Proc = True
                                            
    Call Input_Lock
                                        '品目マスタ読み込み開始
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, "")
    Call UniCode_Conv(K0_ITEM.HIN_GAI, "")
    
    com = BtOpGetGreater
    Do
        DoEvents
        
        
        Do
            sts = BTRV(com + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                                            
                    If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                        sts = BtErrEOF
                        Exit Do
                    End If
                    
                    Exit Do
                    
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, com, "品目マスタ")
                    Exit Function
            End Select
        Loop
        
        If sts = BtErrEOF Then
            Exit Do
        End If
                        
        If (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)) _
            < ((Text(ptxS_Soko_No).Text & Text(ptxS_Retu).Text & Text(ptxS_Ren).Text & Text(ptxS_Dan).Text)) Or _
                (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)) _
            > ((Text(ptxE_Soko_No).Text & Text(ptxE_Retu).Text & Text(ptxE_Ren).Text & Text(ptxE_Dan).Text)) Then
                                '棚番範囲範囲外
        Else
                                        '在庫のチェック
            Call UniCode_Conv(K4_ZAIKO.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K4_ZAIKO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K4_ZAIKO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
            Call UniCode_Conv(K4_ZAIKO.Soko_No, "")
            Call UniCode_Conv(K4_ZAIKO.Retu, "")
            Call UniCode_Conv(K4_ZAIKO.Ren, "")
            Call UniCode_Conv(K4_ZAIKO.Dan, "")
        
            sts = BTRV(BtOpGetGreater, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K4_ZAIKO, Len(K4_ZAIKO), 4)
            Select Case sts
                Case BtNoErr
                    If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                        StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> StrConv(ITEMREC.NAIGAI, vbUnicode) Or _
                        StrConv(ZAIKOREC.HIN_GAI, vbUnicode) <> StrConv(ITEMREC.HIN_GAI, vbUnicode) Then
                        '在庫なしなら入荷のチェック
                        Call UniCode_Conv(K1_IDO.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                        Call UniCode_Conv(K1_IDO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
'                        Call UniCode_Conv(K1_IDO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                        Call UniCode_Conv(K1_IDO.HIN_GAI, "7")
                        Call UniCode_Conv(K1_IDO.JITU_DT, Taisyo_YMD)
                        Call UniCode_Conv(K1_IDO.JITU_TM, "")
                    
                        Nyuka_Flg = False
                    
                        com_IDO = BtOpGetGreaterEqual
                        
                        Do
                            DoEvents
                            sts = BTRV(com_IDO, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
                            Select Case sts
                                Case BtNoErr
                                    If StrConv(IDOREC.JGYOBU, vbUnicode) <> StrConv(ITEMREC.JGYOBU, vbUnicode) Or _
                                        StrConv(IDOREC.NAIGAI, vbUnicode) <> StrConv(ITEMREC.NAIGAI, vbUnicode) Or _
                                        Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                                        Exit Do
                                    End If
                                
                                    If StrConv(IDOREC.RIRK_ID, vbUnicode) = YOIN_TU_NYUKA Then
                                        Nyuka_Flg = True
                                        Exit Do
                                    End If
                                
                                Case BtErrEOF
                                    Exit Do
                                Case Else
                                    Call File_Error(sts, com_IDO, "在庫移動歴")
                                    Exit Function
                            End Select
                        
                            com_IDO = BtOpGetNext
                        
                        Loop
                        If Not Nyuka_Flg Then
                            Do
                                sts = BTRV(BtOpDelete, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                        ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                        If ans = vbCancel Then
                                            Exit Function
                                        End If
                                    Case Else
                                        Call File_Error(sts, BtOpDelete, "品目マスタ")
                                        Exit Function
                                End Select
                            Loop
                            '品目、削除対象
                            Call Log_Out(LOG_F, "削除　品目=[" & StrConv(ITEMREC.JGYOBU, vbUnicode) & "][" & StrConv(ITEMREC.NAIGAI, vbUnicode) & "][" & StrConv(ITEMREC.HIN_GAI, vbUnicode) & "]")
                        End If
                    End If
                Case BtErrEOF
                
                    '在庫なしなら入荷のチェック
                    Call UniCode_Conv(K1_IDO.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K1_IDO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
'                    Call UniCode_Conv(K1_IDO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                    Call UniCode_Conv(K1_IDO.HIN_GAI, "7")
                    Call UniCode_Conv(K1_IDO.JITU_DT, Taisyo_YMD)
                    Call UniCode_Conv(K1_IDO.JITU_TM, "")
                    
                    Nyuka_Flg = False
                    
                    com_IDO = BtOpGetGreaterEqual
                        
                    Do
                        DoEvents
                        sts = BTRV(com_IDO, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
                        Select Case sts
                            Case BtNoErr
                                If StrConv(IDOREC.JGYOBU, vbUnicode) <> StrConv(ITEMREC.JGYOBU, vbUnicode) Or _
                                    StrConv(IDOREC.NAIGAI, vbUnicode) <> StrConv(ITEMREC.NAIGAI, vbUnicode) Or _
                                    Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                                    Exit Do
                                End If
                                
                                If StrConv(IDOREC.RIRK_ID, vbUnicode) = YOIN_TU_NYUKA Then
                                    Nyuka_Flg = True
                                    Exit Do
                                End If
                                
                            Case BtErrEOF
                                Exit Do
                            Case Else
                                Call File_Error(sts, com_IDO, "在庫移動歴")
                                Exit Function
                        End Select
                        
                        com_IDO = BtOpGetNext
                    
                    Loop
                    If Not Nyuka_Flg Then
                        '品目、削除対象
                        Do
                            sts = BTRV(BtOpDelete, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                    If ans = vbCancel Then
                                        Exit Function
                                    End If
                                Case Else
                                    Call File_Error(sts, BtOpDelete, "品目マスタ")
                                    Exit Function
                            End Select
                        Loop
                        Call Log_Out(LOG_F, "削除　品目=[" & StrConv(ITEMREC.JGYOBU, vbUnicode) & "][" & StrConv(ITEMREC.NAIGAI, vbUnicode) & "][" & StrConv(ITEMREC.HIN_GAI, vbUnicode) & "]")
                        
                    End If
                
                Case Else
                    Call File_Error(sts, BtOpGetGreater, "在庫データ")
                    Exit Function
            End Select
        
        End If
        
        com = BtOpGetNext
    
    Loop


    sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpUnlock, "品目マスタ")
        Exit Function
    End If
                                            
    Call Input_UnLock
    
    Delete_Proc = False

End Function
