VERSION 5.00
Begin VB.Form F1010511 
   BackColor       =   &H00FFFFFF&
   Caption         =   "品目マスタメンテナンス（削除機能付き）"
   ClientHeight    =   6915
   ClientLeft      =   1920
   ClientTop       =   2295
   ClientWidth     =   14055
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   14055
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   33
      Left            =   11160
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   89
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   32
      Left            =   12480
      MaxLength       =   4
      TabIndex        =   87
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   31
      Left            =   8880
      MaxLength       =   1
      TabIndex        =   84
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   30
      Left            =   5760
      MaxLength       =   13
      TabIndex        =   82
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   29
      Left            =   2280
      MaxLength       =   13
      TabIndex        =   80
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   21
      Left            =   10200
      MaxLength       =   8
      TabIndex        =   22
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   1800
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   20
      Left            =   9720
      MaxLength       =   2
      TabIndex        =   21
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   25
      Left            =   4200
      MaxLength       =   2
      TabIndex        =   26
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   24
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   25
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   23
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   24
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   22
      Left            =   12840
      MaxLength       =   1
      TabIndex        =   23
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   19
      Left            =   8520
      MaxLength       =   2
      TabIndex        =   20
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   18
      Left            =   7680
      MaxLength       =   2
      TabIndex        =   19
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   17
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   18
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   16
      Left            =   4200
      MaxLength       =   2
      TabIndex        =   17
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   15
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   16
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   14
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   15
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   13
      Left            =   13080
      MaxLength       =   2
      TabIndex        =   14
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   12
      Left            =   12360
      MaxLength       =   2
      TabIndex        =   13
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   11
      Left            =   11640
      MaxLength       =   2
      TabIndex        =   12
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   10
      Left            =   10920
      MaxLength       =   2
      TabIndex        =   11
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   9
      Left            =   8280
      MaxLength       =   2
      TabIndex        =   10
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   7440
      MaxLength       =   2
      TabIndex        =   9
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   6360
      MaxLength       =   4
      TabIndex        =   8
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   5160
      MaxLength       =   13
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   4440
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   5
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   6  '半角ｶﾀｶﾅ
      Index           =   2
      Left            =   8040
      MaxLength       =   25
      TabIndex        =   3
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   1800
      MaxLength       =   13
      TabIndex        =   1
      Top             =   600
      Width           =   1695
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
      TabIndex        =   39
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
      TabIndex        =   38
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
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "データ"
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
      TabIndex        =   36
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
      TabIndex        =   35
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
      TabIndex        =   34
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
      TabIndex        =   33
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
      TabIndex        =   32
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
      TabIndex        =   31
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
      TabIndex        =   30
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
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "更  新"
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
      Index           =   0
      Left            =   120
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   2700
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   28
      Top             =   3000
      Width           =   8655
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   26
      Left            =   6720
      MaxLength       =   4
      TabIndex        =   76
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   27
      Left            =   7800
      MaxLength       =   2
      TabIndex        =   77
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   28
      Left            =   8640
      MaxLength       =   2
      TabIndex        =   78
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "月平均出荷数"
      Height          =   255
      Index           =   24
      Left            =   9360
      TabIndex        =   88
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "個装箱№"
      Height          =   255
      Index           =   42
      Left            =   11280
      TabIndex        =   86
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(0:要　1:不要)"
      Height          =   255
      Index           =   41
      Left            =   9240
      TabIndex        =   85
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "商品化有無"
      Height          =   255
      Index           =   40
      Left            =   7560
      TabIndex        =   83
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "読替えコード"
      Height          =   255
      Index           =   39
      Left            =   4080
      TabIndex        =   81
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ｊａｎコード"
      Height          =   255
      Index           =   38
      Left            =   600
      TabIndex        =   79
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "（最新照合日付"
      Height          =   255
      Index           =   37
      Left            =   4920
      TabIndex        =   75
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "／"
      Height          =   255
      Index           =   36
      Left            =   7440
      TabIndex        =   74
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "／"
      Height          =   255
      Index           =   35
      Left            =   8280
      TabIndex        =   73
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "）"
      Height          =   255
      Index           =   34
      Left            =   9120
      TabIndex        =   72
      Top             =   2160
      Width           =   135
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
      TabIndex        =   71
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   255
      Index           =   33
      Left            =   840
      TabIndex        =   70
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "備考"
      Height          =   255
      Index           =   32
      Left            =   9120
      TabIndex        =   69
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "）"
      Height          =   255
      Index           =   31
      Left            =   4680
      TabIndex        =   68
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "／"
      Height          =   255
      Index           =   30
      Left            =   3840
      TabIndex        =   67
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "／"
      Height          =   255
      Index           =   29
      Left            =   3000
      TabIndex        =   66
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "（最終入荷日付"
      Height          =   255
      Index           =   28
      Left            =   480
      TabIndex        =   65
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "サンプル数"
      Height          =   255
      Index           =   27
      Left            =   11400
      TabIndex        =   64
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "）"
      Height          =   255
      Index           =   23
      Left            =   9000
      TabIndex        =   63
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "／"
      Height          =   255
      Index           =   22
      Left            =   8160
      TabIndex        =   62
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "／"
      Height          =   255
      Index           =   21
      Left            =   7320
      TabIndex        =   61
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "（最終出庫日付"
      Height          =   255
      Index           =   20
      Left            =   4800
      TabIndex        =   60
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "）"
      Height          =   255
      Index           =   19
      Left            =   4680
      TabIndex        =   59
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "／"
      Height          =   255
      Index           =   18
      Left            =   3840
      TabIndex        =   58
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "／"
      Height          =   255
      Index           =   17
      Left            =   3000
      TabIndex        =   57
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "（最終入庫日付"
      Height          =   255
      Index           =   16
      Left            =   480
      TabIndex        =   56
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "）"
      Height          =   255
      Index           =   15
      Left            =   13560
      TabIndex        =   55
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   13
      Left            =   12840
      TabIndex        =   54
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   12
      Left            =   12120
      TabIndex        =   53
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   11
      Left            =   11400
      TabIndex        =   52
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "（前回入庫棚"
      Height          =   255
      Index           =   10
      Left            =   9000
      TabIndex        =   51
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "）"
      Height          =   255
      Index           =   9
      Left            =   8760
      TabIndex        =   50
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "／"
      Height          =   255
      Index           =   8
      Left            =   7920
      TabIndex        =   49
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "／"
      Height          =   255
      Index           =   7
      Left            =   7080
      TabIndex        =   48
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品番（内部）"
      Height          =   255
      Index           =   14
      Left            =   3600
      TabIndex        =   47
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "（設定日付"
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   46
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   5
      Left            =   4200
      TabIndex        =   45
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   44
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   43
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "標準入庫棚"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   42
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品 名"
      Height          =   255
      Index           =   1
      Left            =   7200
      TabIndex        =   41
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品番（外部）"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   40
      Top             =   720
      Width           =   1575
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1010511"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim LIST_MAX    As Integer              'リストボックス最大表示件数

Dim Text_Max    As Integer              '画面項目別最大ｲﾝﾃﾞｯｸｽ
Dim Combo_Max   As Integer
Dim Command_Max As Integer
''Dim JIGYOBU_BEF As String * 1         'ｱｲﾛﾝ元事業部
Dim ITEM_CSV    As String



Private Function List_Disp()
Dim sts As Integer
Dim com As Integer
Dim i As Integer
Dim Sv_Naigai As String * 1
Dim Edit As String
    
    List_Disp = False
    
    List1.Clear
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    If Combo(0).Text = NAIGAI1$ Then
        Sv_Naigai = NAIGAI_NAI$
    Else
        Sv_Naigai = NAIGAI_GAI$
    End If
    Call UniCode_Conv(K0_ITEM.NAIGAI, Sv_Naigai)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, RTrim(Text(0).Text))
    
    com = BtOpGetGreaterEqual
    For i = 0 To LIST_MAX - 1
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) <> Sv_Naigai Then
                    Exit For
                End If
            Case BtErrEOF
                Exit Function
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                List_Disp = True
                Exit Function
        End Select
        
        Edit = StrConv(ITEMREC.HIN_GAI, vbUnicode) & " " & StrConv(ITEMREC.HIN_NAI, vbUnicode) & " " & StrConv(ITEMREC.HIN_NAME, vbUnicode) & " "
        Edit = Edit & StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" + StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode)
        List1.AddItem Edit
        
        com = BtOpGetNext
    Next i
    
End Function
                                    '画面初期状態を設定する
Private Sub Clear_Field(Mode As Integer)
Dim i As Integer

    If (Mode = 0) Then
        Text(0).Text = ""
    End If
    
    For i = 1 To 32
        Text(i).Text = ""
    Next i
End Sub

'                                       入力項目のエラーチェック
Private Function Del_Chk() As Integer
            
Dim RetBuf As String
Dim i As Integer
Dim sts As Integer


    Del_Chk = False
    
    
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, "")
    Call UniCode_Conv(K1_ZAIKO.Soko_No, "")
    Call UniCode_Conv(K1_ZAIKO.Retu, "")
    Call UniCode_Conv(K1_ZAIKO.Ren, "")
    Call UniCode_Conv(K1_ZAIKO.Dan, "")

    sts = BTRV(BtOpGetGreaterEqual, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
    Select Case sts
        Case BtNoErr
            If StrConv(ZAIKOREC.JGYOBU, vbUnicode) = StrConv(ZAIKOREC.JGYOBU, vbUnicode) And _
                StrConv(ZAIKOREC.NAIGAI, vbUnicode) = StrConv(ITEMREC.NAIGAI, vbUnicode) And _
                StrConv(ZAIKOREC.HIN_GAI, vbUnicode) = StrConv(ITEMREC.HIN_GAI, vbUnicode) Then
                Beep
                MsgBox "有効在庫残有り！！削除できません。"
                Text(0).SelStart = 0
                Text(0).SelLength = Len(Text(0).Text)
                Text(0).SetFocus
                Del_Chk = True
                Exit Function
            End If
        Case BtErrEOF
        Case Else
            Call File_Error(sts, BtOpGetGreaterEqual, "在庫データ")
            Del_Chk = SYS_ERR
    End Select

End Function
'                                       入力項目のエラーチェック
Private Function Err_Chk() As Integer
            
Dim RetBuf As String
Dim i As Integer
Dim sts As Integer


    Err_Chk = False
    
    If Len(RTrim(Text(0).Text)) = 0 Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text(0).SetFocus
        Err_Chk = True
        Exit Function
    End If
    
                                            '標準入庫棚チェック
    If Len(RTrim(Text(3).Text)) = 0 Then
        Text(4).Text = ""
        Text(5).Text = ""
        Text(6).Text = ""
    Else
        For i = 4 To 6
            If Not IsNumeric(Text(i).Text) Then
                Beep
                MsgBox "入力した項目はエラーです。"
                Text(i).SetFocus
                Err_Chk = True
                Exit Function
            Else
                Text(i).Text = Format(CInt(Text(i).Text), "00")
            End If
        Next i
        Call UniCode_Conv(K0_SOKO.Soko_No, Text(3).Text)
        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG$ Then
                    If StrConv(SOKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                        Beep
                        MsgBox "入力した項目はエラーです。（混載エラー）"
                        Text(3).SetFocus
                        Err_Chk = True
                        Exit Function
                    End If
                End If
            Case BtErrKeyNotFound
                Beep
                MsgBox "入力した項目はエラーです。（未登録エラー）"
                Text(3).SetFocus
                Err_Chk = True
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
                                                '棚マスタ読み込み
        Call UniCode_Conv(K0_TANA.Soko_No, Text(3).Text)
        Call UniCode_Conv(K0_TANA.Retu, Text(4).Text)
        Call UniCode_Conv(K0_TANA.Ren, Text(5).Text)
        Call UniCode_Conv(K0_TANA.Dan, Text(6).Text)
        sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Beep
                MsgBox "入力した項目はエラーです。（未登録エラー）"
                Text(3).SetFocus
                Err_Chk = True
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "棚マスタ")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
    End If
                                        'サンプル数
    If Len(RTrim(Text(22).Text)) = 0 Then
        Text(22).Text = "1"
    End If
    If Not IsNumeric(Text(22).Text) Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text(22).SetFocus
        Err_Chk = True
        Exit Function
    Else
        Text(22).Text = Format(CLng(Text(22).Text), "0")
    End If
                                        'JANコード
    If Len(Trim(Text(29).Text)) <> 0 Then
        If Len(RTrim(Text(29).Text)) <> Text(29).MaxLength Then
            Beep
            MsgBox "入力した項目はエラーです。"
            Text(22).SetFocus
            Err_Chk = True
            Exit Function
        End If

        If Not IsNumeric(Text(29).Text) Then
            Beep
            MsgBox "入力した項目はエラーです。"
            Text(22).SetFocus
            Err_Chk = True
            Exit Function
        End If


        Call UniCode_Conv(K4_ITEM.JGYOBU, Last_JGYOBU)
        If Combo(0).Text = NAIGAI1$ Then
            Call UniCode_Conv(K4_ITEM.NAIGAI, NAIGAI_NAI$)
        Else
            Call UniCode_Conv(K4_ITEM.NAIGAI, NAIGAI_GAI$)
        End If
        Call UniCode_Conv(K4_ITEM.JAN_CODE, Text(29).Text)
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K4_ITEM, Len(K4_ITEM), 4)
        Select Case sts
            Case BtNoErr
                If Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) <> Trim(Text(0).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。(登録済み)"
                    Text(29).SetFocus
                    Err_Chk = True
                    Exit Function
                End If
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
    End If
                                        '読替えコード
    If Len(Trim(Text(30).Text)) <> 0 Then
        Call UniCode_Conv(K5_ITEM.JGYOBU, Last_JGYOBU)
        If Combo(0).Text = NAIGAI1$ Then
            Call UniCode_Conv(K5_ITEM.NAIGAI, NAIGAI_NAI$)
        Else
            Call UniCode_Conv(K5_ITEM.NAIGAI, NAIGAI_GAI$)
        End If
        Call UniCode_Conv(K5_ITEM.HIN_CHANGE, Text(30).Text)
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K5_ITEM, Len(K5_ITEM), 5)
        Select Case sts
            Case BtNoErr
                If Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) <> Trim(Text(0).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。(登録済み)"
                    Text(30).SetFocus
                    Err_Chk = True
                    Exit Function
                End If
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
    End If
                                        '商品化有無
    If Text(31).Text = "" Then
        Text(31).Text = "0"
    Else
        If Text(31).Text <> "0" And Text(31).Text <> "1" Then
            Beep
            MsgBox "入力した項目はエラーです。"
            Text(31).SetFocus
            Err_Chk = True
            Exit Function
        End If
    End If
                                        '個装箱№
    If Text(32).Text <> "" Then
        Call UniCode_Conv(K0_PACKING.PACKING_NO, Text(32).Text)
        sts = BTRV(BtOpGetEqual, PACKING_POS, PACKINGREC, Len(PACKINGREC), K0_PACKING, Len(K0_PACKING), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Beep
                MsgBox "入力した項目はエラーです。（未登録エラー）"
                Text(32).SetFocus
                Err_Chk = True
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "個装箱マスタ")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
    End If

End Function

Private Sub Item_Dsp()
Dim sts         As Integer
Dim Work_Date   As String * 8
Dim RetBuf      As String
            
    Text(0).Text = RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
    Text(1).Text = RTrim(StrConv(ITEMREC.HIN_NAI, vbUnicode))
    Text(2).Text = RTrim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
    Text(3).Text = StrConv(ITEMREC.ST_SOKO, vbUnicode)
    Text(4).Text = StrConv(ITEMREC.ST_RETU, vbUnicode)
    Text(5).Text = StrConv(ITEMREC.ST_REN, vbUnicode)
    Text(6).Text = StrConv(ITEMREC.ST_DAN, vbUnicode)
    Work_Date = StrConv(ITEMREC.ST_SET_DT, vbUnicode)
    Text(7).Text = Mid(Work_Date, 1, 4)
    Text(8).Text = Mid(Work_Date, 5, 2)
    Text(9).Text = Mid(Work_Date, 7, 2)
    Text(10).Text = StrConv(ITEMREC.BEF_SOKO, vbUnicode)
    Text(11).Text = StrConv(ITEMREC.BEF_RETU, vbUnicode)
    Text(12).Text = StrConv(ITEMREC.BEF_REN, vbUnicode)
    Text(13).Text = StrConv(ITEMREC.BEF_DAN, vbUnicode)
    Work_Date = StrConv(ITEMREC.LAST_NYU_DT, vbUnicode)
    Text(14).Text = Mid(Work_Date, 1, 4)
    Text(15).Text = Mid(Work_Date, 5, 2)
    Text(16).Text = Mid(Work_Date, 7, 2)
    Work_Date = StrConv(ITEMREC.LAST_SYU_DT, vbUnicode)
    Text(17).Text = Mid(Work_Date, 1, 4)
    Text(18).Text = Mid(Work_Date, 5, 2)
    Text(19).Text = Mid(Work_Date, 7, 2)
    Text(20).Text = RTrim(StrConv(ITEMREC.BIKOU_SOKO, vbUnicode))
    Text(21).Text = RTrim(StrConv(ITEMREC.BIKOU_TANA, vbUnicode))
    Text(22).Text = StrConv(ITEMREC.SAMPLE_QTY, vbUnicode)
    Work_Date = StrConv(ITEMREC.LAST_INP_DT, vbUnicode)
    Text(23).Text = Mid(Work_Date, 1, 4)
    Text(24).Text = Mid(Work_Date, 5, 2)
    Text(25).Text = Mid(Work_Date, 7, 2)
    Work_Date = StrConv(ITEMREC.LAST_CHK_DT, vbUnicode)
    Text(26).Text = Mid(Work_Date, 1, 4)
    Text(27).Text = Mid(Work_Date, 5, 2)
    Text(28).Text = Mid(Work_Date, 7, 2)

    Text(29).Text = StrConv(ITEMREC.JAN_CODE, vbUnicode)
    Text(30).Text = StrConv(ITEMREC.HIN_CHANGE, vbUnicode)
    Text(31).Text = StrConv(ITEMREC.GOODS_KBN, vbUnicode)
    Text(32).Text = StrConv(ITEMREC.PACKING_NO, vbUnicode)
    If IsNumeric(CLng(StrConv(ITEMREC.AVE_SYUKA, vbUnicode))) Then
        Text(33).Text = Format(CDbl(StrConv(ITEMREC.AVE_SYUKA, vbUnicode)), "#0.0")
    Else
        Text(33).Text = ""

    End If
End Sub

                                            '品目マスタの追加／訂正
Private Function Update_Proc() As Integer
Dim sts As Integer
Dim ans As Integer
Dim com As Integer
Dim Sv_Naigai As String * 1
Dim RetBuf As String
Dim Edit As String
Dim i As Integer

    Update_Proc = False

    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    If Combo(0).Text = NAIGAI1 Then
        Sv_Naigai = NAIGAI_NAI$
    Else
        Sv_Naigai = NAIGAI_GAI$
    End If
    Call UniCode_Conv(K0_ITEM.NAIGAI, Sv_Naigai)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(0).Text)
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Call Clear_Field(0)
                    Text(0).SetFocus
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Update_Proc = True
        End Select
    Loop
                                            'レコード内容編集
    Call UniCode_Conv(ITEMREC.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(ITEMREC.NAIGAI, Sv_Naigai)
    Call UniCode_Conv(ITEMREC.HIN_GAI, Text(0).Text)
    Call UniCode_Conv(ITEMREC.HIN_NAI, Text(1).Text)
    Call UniCode_Conv(ITEMREC.HIN_NAME, Text(2).Text)
    If Len(RTrim(Text(3).Text)) = 0 Then
        Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
    Else
        If com = BtOpUpdate Then
            If (StrConv(ITEMREC.ST_SOKO, vbUnicode) <> Text(3).Text) Or _
                (StrConv(ITEMREC.ST_RETU, vbUnicode) <> Text(4).Text) Or _
                (StrConv(ITEMREC.ST_REN, vbUnicode) <> Text(5).Text) Or _
                (StrConv(ITEMREC.ST_DAN, vbUnicode) <> Text(6).Text) Then
                Call UniCode_Conv(ITEMREC.ST_SET_DT, (Left(Format(Date, "yyyymmdd"), 4) + Mid(Format(Date, "yyyymmdd"), 5, 2) + Mid(Format(Date, "yyyymmdd"), 7, 2)))
            End If
        Else
            Call UniCode_Conv(ITEMREC.ST_SET_DT, (Left(Format(Date, "yyyymmdd"), 4) + Mid(Format(Date, "yyyymmdd"), 5, 2) + Mid(Format(Date, "yyyymmdd"), 7, 2)))
        End If
    End If
    Call UniCode_Conv(ITEMREC.ST_SOKO, Text(3).Text)
    Call UniCode_Conv(ITEMREC.ST_RETU, Text(4).Text)
    Call UniCode_Conv(ITEMREC.ST_REN, Text(5).Text)
    Call UniCode_Conv(ITEMREC.ST_DAN, Text(6).Text)
    If com = BtOpInsert Then
        Call UniCode_Conv(ITEMREC.BEF_SOKO, "")
        Call UniCode_Conv(ITEMREC.BEF_RETU, "")
        Call UniCode_Conv(ITEMREC.BEF_REN, "")
        Call UniCode_Conv(ITEMREC.BEF_DAN, "")
    
        Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")
        Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")
        Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")
        
        Call UniCode_Conv(ITEMREC.LOCK_F, LOCK_OFF)
        Call UniCode_Conv(ITEMREC.WEL_ID, "")
        Call UniCode_Conv(ITEMREC.PRG_ID, "")
        Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")
        Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "00000000")
        Call UniCode_Conv(ITEMREC.FILLER, "")
    
        Call UniCode_Conv(ITEMREC.SIZAI_CD, "")
        Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")
        Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")
    
    End If
    
    Call UniCode_Conv(ITEMREC.HIN_NAI, Text(1).Text)
    Call UniCode_Conv(ITEMREC.BIKOU_SOKO, Text(20).Text)
    Call UniCode_Conv(ITEMREC.BIKOU_TANA, Text(21).Text)
    Call UniCode_Conv(ITEMREC.SAMPLE_QTY, Format(CInt(Text(22).Text), "0"))
    Call UniCode_Conv(ITEMREC.JAN_CODE, Text(29).Text)
    Call UniCode_Conv(ITEMREC.HIN_CHANGE, Text(30).Text)
    Call UniCode_Conv(ITEMREC.GOODS_KBN, Text(31).Text)
    Call UniCode_Conv(ITEMREC.PACKING_NO, Text(32).Text)
            
            
    Do
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Call Clear_Field(0)
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Beep
                MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly + vbCritical
                Update_Proc = True
        End Select
    Loop
                                        'リストボックス追加
    If com = BtOpUpdate Then
        For i = 0 To List1.ListCount - 1
            If RTrim(Left$(List1.List(i), 13)) = RTrim(Text(0).Text) Then
                List1.RemoveItem i
            End If
        Next i
    End If
        
    Edit = StrConv(ITEMREC.HIN_GAI, vbUnicode) + " " + StrConv(ITEMREC.HIN_NAI, vbUnicode) + " " + StrConv(ITEMREC.HIN_NAME, vbUnicode) + " "
    Edit = Edit + StrConv(ITEMREC.ST_SOKO, vbUnicode) + "-" + StrConv(ITEMREC.ST_RETU, vbUnicode) + "-" + StrConv(ITEMREC.ST_REN, vbUnicode) + "-" + StrConv(ITEMREC.ST_DAN, vbUnicode)
    List1.AddItem Edit
                                        '画面クリアー
    Call Clear_Field(0)
'

End Function
                                            '品目マスタの削除
Private Function Delete_Proc() As Integer
Dim sts As Integer
Dim ans As Integer
Dim com As Integer
Dim Sv_Naigai As String * 1
Dim RetBuf As String
Dim Edit As String
Dim i As Integer
    
    Delete_Proc = False

    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    If Combo(0).Text = NAIGAI1 Then
        Sv_Naigai = NAIGAI_NAI$
    Else
        Sv_Naigai = NAIGAI_GAI$
    End If
    Call UniCode_Conv(K0_ITEM.NAIGAI, Sv_Naigai)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(0).Text)
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
'                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
'                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Call Clear_Field(0)
                    Text(0).SetFocus
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Delete_Proc = True
                Exit Function
        End Select
    Loop
            
    If sts = BtNoErr Then
            
        Do
            sts = BTRV(BtOpDelete, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Call Clear_Field(0)
                        Text(0).SetFocus
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpDelete, "品目マスタ")
                    Beep
                    MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly + vbCritical
                    Delete_Proc = True
                    Exit Function
            End Select
        Loop
    End If
                                        'リストボックス削除
'    For i = 0 To List1.ListCount - 1
'        If RTrim(Left$(List1.List(i), 13)) = RTrim(Text(0).Text) Then
'            List1.RemoveItem i
'        End If
'    Next i
                                        '画面クリアー
    Call Clear_Field(0)
'

End Function

Private Sub Combo_DblClick(Index As Integer)
    
    Call Clear_Field(0)
    List1.Clear

End Sub


Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Select Case Index
                Case 0
                    Call Clear_Field(0)
                    List1.Clear
                    Text(0).SetFocus
            End Select
    End Select

End Sub


Private Sub Combo_LostFocus(Index As Integer)
'        Call Clear_Field(0)
'        List1.Clear

End Sub

Private Sub Command_Click(Index As Integer)

Dim yn  As Integer
Dim i   As Integer
Dim sts As Integer
    
    Select Case Index
        Case 0                              'データ更新
                                            
            
            Beep
            yn = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
                Text(0).SetFocus
            End If
        Case 3                              '削除
            sts = Del_Chk()
                
            Select Case sts
                Case False
                Case True
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
            
            Beep
            yn = MsgBox("削除しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                If Delete_Proc() Then
                    Unload Me
                End If
                Text(0).SetFocus
            End If
        Case 8                              'データ出力
                        
            If CSV_OUTPUT_Proc() Then
                Unload Me
            End If
        Case 11                             '終了
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
Dim i       As Integer
Dim c       As String * 128
Dim sts     As Integer


    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
    
    Text_Max = 1                '印刷指示画面項目別最大ｲﾝﾃﾞｯｸｽ
    Command_Max = 2
    
    Show
    
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                
''                                'ｱｲﾛﾝ事業部元取り込み
''    If GetIni("SENTAKU", "JIGYOBU_BEF", "SYS", c) Then
''        Beep
''        MsgBox "ｱｲﾛﾝ事業部元の獲得に失敗しました。処理を中止して下さい。"
''        End
''    End If
''    JIGYOBU_BEF = RTrim(c)
    
                                'ＣＳＶファイル名取り込み
    If GetIni("FILE", "ITEM_CSV", "SYS", c) Then
        Beep
        MsgBox "品目マスタデータ出力用ファイル[ITEM_CSV]の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    ITEM_CSV = Trim(c)
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
            F1010511.Caption = "品目マスタメンテナンス（削除機能付き）（" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
                                'リストボックス最大表示件数表示
    If GetIni("F101051", "LISTMAX", "SYS", c) Then
        Beep
        MsgBox "最大表示件数の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LIST_MAX = CInt(RTrim(c))
                                '国内外取り込み
    Combo(0).AddItem NAIGAI1$
    Combo(0).AddItem NAIGAI2$
    Combo(0).Text = NAIGAI1$
                                
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '棚マスタＯＰＥＮ
    If TANA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '個装箱マスタＯＰＥＮ
    If PACKING_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '画面初期設定
    Call Clear_Field(0)
    
    Combo(0).SetFocus
    
    End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫マスタ")
        End If
    End If
                                            '棚マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "棚マスタ")
        End If
    End If
                                            '個装箱マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, PACKING_POS, PACKINGREC, Len(PACKINGREC), K0_PACKING, Len(K0_PACKING), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "個装箱マスタ")
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
    
    sts = BTRV(BtOpReset, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1010511 = Nothing

    End
End Sub

Private Sub List1_DblClick()
Dim sts As Integer
                                        'リストボックスより項目表示
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Mid$(List1.List(List1.ListIndex), 1, 13))
                                                '外部品番で読み込み
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            Call Item_Dsp
            Text(1).SetFocus
        Case BtErrKeyNotFound           'これは無いはず
            MsgBox "マスタ内容が変更されています。最新情報を再表示します。"
            If List_Disp() Then
                Unload Me
            End If
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Beep
            MsgBox "システム異常が発生しました。処理を中止して下さい。"
            Unload Me
    End Select

End Sub


Private Sub List1_GotFocus()
    
    If List1.ListCount > 0 Then
        If List1.ListIndex <= 0 Then
            List1.ListIndex = 0
        End If
    End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sts As Integer
                                        
    If List1.ListCount = 0 Then
        Exit Sub
    End If
                                        'リストボックスより項目表示
    Select Case KeyCode
        Case vbKeyReturn
            Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Mid$(List1.List(List1.ListIndex), 1, 13))
                                                '外部品番で読み込み
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Call Item_Dsp
                    Text(1).SetFocus
                Case BtErrKeyNotFound           'これは無いはず
                    MsgBox "マスタ内容が変更されています。最新情報を再表示します。"
                    If List_Disp() Then
                        Unload Me
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Unload Me
            End Select
    End Select
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
    F1010511.Caption = "品目マスタメンテナンス（削除機能付き）（" + RTrim(JGYOBU_T(Index).NAME) + "）"
    Last_JGYOBU = JGYOBU_T(Index).Code
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
Dim RetBuf As String
Dim i As Integer
Dim sts As Integer

    Select Case KeyCode
        Case vbKeyReturn
            Select Case Index
                Case 0
                        If List_Disp() Then
                            Unload Me
                        End If
                                                '外部品番で読み込み
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                        If Combo(0).Text = NAIGAI1$ Then
                            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI$)
                        Else
                            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI$)
                        End If
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, RTrim(Text(0).Text))
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                                Call Item_Dsp
                            Case BtErrKeyNotFound
                                Call Clear_Field(1)
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                Unload Me
                        End Select
                Case 1
                Case 29             'Janコード
                Case 30             '読替えコード
            End Select
            For i = Index + 1 To 32
                If Text(i).Enabled Then
                    Text(i).SetFocus
                    Exit For
                End If
            Next i
    End Select
End Sub



Private Function CSV_OUTPUT_Proc() As Integer

Dim FileNo          As Integer
Dim fileName        As String
Dim Ret             As Integer

Dim com             As Integer
Dim sts             As Integer

Dim c               As String * 128

Dim Soko_No         As String * 2

    Call Input_Lock
    
    FileNo = FreeFile
    fileName = ITEM_CSV
    
    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)

    On Error GoTo Error_Proc

    Open (fileName) For Output As FileNo
    On Error GoTo 0
    
'    Write #FileNo, "事業部", "内外", "品番（外部）", "品名", "標準棚番設定日", "標準棚番", "前回棚番", "最終入庫日", "最終出庫日", "品番（内部）", "ホスト倉庫", "ホスト棚番", "サンプル数", "最終入荷日", "最終照合日", "照合時在庫数", "印刷備考", "印刷入り数", "ＪＡＮ", "読み替え", "商品化有無", "個装箱№"
    Write #FileNo, "内外", "品番（外部）", "品名", "標準棚番設定日", "標準棚番", "前回棚番", "最終入庫日", "最終出庫日", "品番（内部）", "ホスト倉庫", "ホスト棚番", "サンプル数", "最終入荷日", "最終照合日", "照合時在庫数", "印刷備考", "印刷入り数", "ＪＡＮ", "読み替え", "商品化有無", "個装箱№"

    com = BtOpGetFirst
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
                
                Call Input_UnLock
                Exit Function
        End Select
    
'        Write #FileNo, StrConv(ITEMREC.JGYOBU, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.NAIGAI, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.HIN_GAI, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.HIN_NAME, vbUnicode),
        
        If IsDate(StrConv(ITEMREC.ST_SET_DT, vbUnicode)) Then
            Write #FileNo, Format(StrConv(ITEMREC.ST_SET_DT, vbUnicode), "YYYY/MM/DD"),
        Else
            Write #FileNo, ,
        End If
        
        
        If GetIni("SOKO_NO", StrConv(ITEMREC.ST_SOKO, vbUnicode), "SYS", c) Then
            Soko_No = StrConv(ITEMREC.ST_SOKO, vbUnicode)
        Else
            Soko_No = Trim(c)
        End If
        Write #FileNo, Soko_No & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode),
        
        If GetIni("SOKO_NO", StrConv(ITEMREC.BEF_SOKO, vbUnicode), "SYS", c) Then
            Soko_No = StrConv(ITEMREC.BEF_SOKO, vbUnicode)
        Else
            Soko_No = Trim(c)
        End If
        Write #FileNo, Soko_No & "-" & StrConv(ITEMREC.BEF_RETU, vbUnicode) & "-" & StrConv(ITEMREC.BEF_REN, vbUnicode) & "-" & StrConv(ITEMREC.BEF_DAN, vbUnicode),
    
        Write #FileNo, StrConv(ITEMREC.LAST_NYU_DT, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.LAST_SYU_DT, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.HIN_NAI, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.BIKOU_SOKO, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.BIKOU_TANA, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.SAMPLE_QTY, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.LAST_INP_DT, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.LAST_CHK_DT, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.LAST_CHK_QTY, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.BIKOU, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.IRI_QTY, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.JAN_CODE, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.HIN_CHANGE, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.GOODS_KBN, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.PACKING_NO, vbUnicode)
    
    
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
        Call Input_UnLock
        CSV_OUTPUT_Proc = False
    Else
        MsgBox "Err.Number= " & Err.Number
        CSV_OUTPUT_Proc = True
    End If

    Call Input_UnLock

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    F1010511.MousePointer = vbHourglass

    Call Ctrl_Lock(F1010511)

End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1010511)

    F1010511.MousePointer = vbDefault

End Sub

