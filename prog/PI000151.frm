VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form PI000151 
   Caption         =   "商品化指図票発行(受注機能付き)"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16545
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
   ScaleHeight     =   9525
   ScaleWidth      =   16545
   StartUpPosition =   2  '画面の中央
   Begin VB.ComboBox Combo2 
      Height          =   360
      Index           =   8
      Left            =   1200
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   115
      Top             =   8400
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Index           =   7
      Left            =   1200
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   106
      Top             =   8040
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Index           =   6
      Left            =   1200
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   97
      Top             =   7680
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Index           =   5
      Left            =   1200
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   88
      Top             =   7320
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Index           =   4
      Left            =   1200
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   79
      Top             =   6960
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Index           =   3
      Left            =   1200
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   70
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   1590
      MaxLength       =   10
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   12
      Left            =   14160
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   11
      Left            =   13200
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1080
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   1800
      TabIndex        =   176
      Top             =   2160
      Width           =   5925
      Begin VB.OptionButton Option1 
         Caption         =   "再梱包"
         Height          =   375
         Index           =   3
         Left            =   4410
         TabIndex        =   180
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "欠品解除"
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "事前"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   178
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "スポット"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   98
      Left            =   11520
      TabIndex        =   122
      Top             =   8400
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   97
      Left            =   10920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   8400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   96
      Left            =   9480
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   120
      TabStop         =   0   'False
      Top             =   8400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   95
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   119
      TabStop         =   0   'False
      Top             =   8400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   94
      Left            =   7560
      MaxLength       =   6
      TabIndex        =   118
      Top             =   8400
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   93
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   117
      TabStop         =   0   'False
      Top             =   8400
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   92
      Left            =   2520
      MaxLength       =   20
      TabIndex        =   116
      Top             =   8400
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   91
      Left            =   11520
      TabIndex        =   113
      Top             =   8040
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   90
      Left            =   10920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   8040
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   89
      Left            =   9480
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   88
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   87
      Left            =   7560
      MaxLength       =   6
      TabIndex        =   109
      Top             =   8040
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   86
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   8040
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   85
      Left            =   2520
      MaxLength       =   20
      TabIndex        =   107
      Top             =   8040
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   84
      Left            =   11520
      TabIndex        =   104
      Top             =   7680
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   83
      Left            =   10920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   7680
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   82
      Left            =   9480
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   81
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   80
      Left            =   7560
      MaxLength       =   6
      TabIndex        =   100
      Top             =   7680
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   79
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   7680
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   78
      Left            =   2520
      MaxLength       =   20
      TabIndex        =   98
      Top             =   7680
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   77
      Left            =   11520
      TabIndex        =   95
      Top             =   7320
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   76
      Left            =   10920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   7320
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   75
      Left            =   9480
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   74
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   73
      Left            =   7560
      MaxLength       =   6
      TabIndex        =   91
      Top             =   7320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   72
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   7320
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   71
      Left            =   2520
      MaxLength       =   20
      TabIndex        =   89
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   70
      Left            =   11520
      TabIndex        =   86
      Top             =   6960
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   69
      Left            =   10920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   6960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   68
      Left            =   9480
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   67
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   66
      Left            =   7560
      MaxLength       =   6
      TabIndex        =   82
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   65
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   6960
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   64
      Left            =   2520
      MaxLength       =   20
      TabIndex        =   80
      Top             =   6960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   56
      Left            =   13320
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   55
      Left            =   12240
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   54
      Left            =   11400
      MaxLength       =   6
      TabIndex        =   66
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   53
      Left            =   9240
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   52
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   64
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   51
      Left            =   13320
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   50
      Left            =   12240
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   49
      Left            =   11400
      MaxLength       =   6
      TabIndex        =   61
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   48
      Left            =   9240
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   47
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   59
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   46
      Left            =   13320
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   45
      Left            =   12240
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   44
      Left            =   11400
      MaxLength       =   6
      TabIndex        =   56
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   43
      Left            =   9240
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   42
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   54
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   41
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   40
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   39
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   51
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   38
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   5760
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   37
      Left            =   600
      MaxLength       =   20
      TabIndex        =   49
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   36
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   35
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   34
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   46
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   33
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   5400
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   32
      Left            =   600
      MaxLength       =   20
      TabIndex        =   44
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   31
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   30
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   29
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   41
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   28
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   27
      Left            =   600
      MaxLength       =   20
      TabIndex        =   39
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   26
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   25
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   24
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   36
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   16
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   19
      Top             =   1800
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1095
      Index           =   0
      Left            =   8040
      TabIndex        =   28
      Top             =   2640
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1931
      _Version        =   393217
      TextRTF         =   $"PI000151.frx":0000
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   8
      Left            =   240
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   114
      Top             =   8400
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   7
      Left            =   240
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   105
      Top             =   8040
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   6
      Left            =   240
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   96
      Top             =   7680
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   5
      Left            =   240
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   87
      Top             =   7320
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   4
      Left            =   240
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   78
      Top             =   6960
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   3
      Left            =   240
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   69
      Top             =   6600
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "出力対象"
      Height          =   855
      Left            =   240
      TabIndex        =   158
      Top             =   2880
      Width           =   6735
      Begin VB.CheckBox Check1 
         Caption         =   "機種ラベル"
         Height          =   375
         Index           =   4
         Left            =   6240
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "外装ラベル"
         Height          =   375
         Index           =   3
         Left            =   4320
         TabIndex        =   26
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "パーツラベル"
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "指図票"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   23
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   22
      Left            =   600
      MaxLength       =   20
      TabIndex        =   34
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   21
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   20
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   19
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   31
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Index           =   18
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   17
      Left            =   600
      MaxLength       =   20
      TabIndex        =   29
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   6
      Left            =   8850
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   13
      Left            =   240
      MaxLength       =   5
      TabIndex        =   15
      Top             =   1800
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "見本作成"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   21
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   2
      Left            =   9960
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   20
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   15
      Left            =   5640
      MaxLength       =   10
      TabIndex        =   18
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   14
      Left            =   4200
      MaxLength       =   10
      TabIndex        =   17
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   1
      Left            =   960
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   16
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   10
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   9
      Left            =   9960
      MaxLength       =   8
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   8
      Left            =   5640
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   3120
      MaxLength       =   20
      TabIndex        =   8
      Top             =   1080
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   240
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   7
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   8130
      MaxLength       =   5
      TabIndex        =   5
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   4
      Left            =   5970
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   5250
      MaxLength       =   5
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   3450
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   240
      MaxLength       =   8
      TabIndex        =   0
      Top             =   360
      Width           =   1155
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
      Left            =   10440
      TabIndex        =   134
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取 消"
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
      Left            =   9600
      TabIndex        =   133
      Top             =   9000
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
      Left            =   8760
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "印 刷"
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
      Left            =   7920
      TabIndex        =   131
      Top             =   9000
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
      Left            =   6600
      TabIndex        =   130
      TabStop         =   0   'False
      Top             =   9000
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
      Left            =   5760
      TabIndex        =   129
      TabStop         =   0   'False
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "構成部品"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4920
      TabIndex        =   128
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4080
      TabIndex        =   127
      TabStop         =   0   'False
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "M更新"
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
      Left            =   2760
      TabIndex        =   126
      Top             =   9000
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
      Left            =   1920
      TabIndex        =   125
      TabStop         =   0   'False
      Top             =   9000
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
      Left            =   1080
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   9000
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
      Left            =   240
      TabIndex        =   123
      Top             =   9000
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   63
      Left            =   11520
      TabIndex        =   77
      Top             =   6600
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   62
      Left            =   10920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   6600
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   61
      Left            =   9480
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   60
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   59
      Left            =   7560
      MaxLength       =   6
      TabIndex        =   73
      Top             =   6600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   58
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   6600
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   57
      Left            =   2520
      MaxLength       =   20
      TabIndex        =   71
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "事業部"
      Height          =   255
      Index           =   25
      Left            =   1200
      TabIndex        =   181
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "注文№"
      Height          =   255
      Index           =   24
      Left            =   1560
      TabIndex        =   179
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "商品化済"
      Height          =   255
      Index           =   23
      Left            =   14040
      TabIndex        =   14
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "未商品"
      Height          =   255
      Index           =   17
      Left            =   13200
      TabIndex        =   177
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "棚番"
      Height          =   255
      Index           =   16
      Left            =   9840
      TabIndex        =   175
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "数量"
      Height          =   255
      Index           =   22
      Left            =   8880
      TabIndex        =   174
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "品名"
      Height          =   255
      Index           =   21
      Left            =   5040
      TabIndex        =   173
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "内職ｸﾗｽ"
      Height          =   255
      Index           =   20
      Left            =   7320
      TabIndex        =   172
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "備考"
      Height          =   255
      Index           =   19
      Left            =   12240
      TabIndex        =   171
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "在庫"
      Height          =   255
      Index           =   18
      Left            =   11040
      TabIndex        =   170
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "員数"
      Height          =   255
      Index           =   15
      Left            =   7800
      TabIndex        =   169
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "品番"
      Height          =   255
      Index           =   14
      Left            =   2760
      TabIndex        =   168
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "種別"
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   167
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "外装資材№"
      Enabled         =   0   'False
      Height          =   375
      Index           =   17
      Left            =   7560
      TabIndex        =   166
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "①"
      Enabled         =   0   'False
      Height          =   375
      Index           =   16
      Left            =   7560
      TabIndex        =   165
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "②"
      Enabled         =   0   'False
      Height          =   375
      Index           =   15
      Left            =   7560
      TabIndex        =   164
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "③"
      Enabled         =   0   'False
      Height          =   375
      Index           =   14
      Left            =   7560
      TabIndex        =   163
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "品名"
      Enabled         =   0   'False
      Height          =   375
      Index           =   13
      Left            =   9240
      TabIndex        =   162
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "入数"
      Enabled         =   0   'False
      Height          =   375
      Index           =   12
      Left            =   11400
      TabIndex        =   161
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "数量"
      Enabled         =   0   'False
      Height          =   375
      Index           =   11
      Left            =   12240
      TabIndex        =   160
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "棚番"
      Enabled         =   0   'False
      Height          =   375
      Index           =   10
      Left            =   13320
      TabIndex        =   159
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "棚番"
      Enabled         =   0   'False
      Height          =   375
      Index           =   9
      Left            =   6000
      TabIndex        =   157
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "数量"
      Enabled         =   0   'False
      Height          =   375
      Index           =   8
      Left            =   4920
      TabIndex        =   156
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "員数"
      Enabled         =   0   'False
      Height          =   375
      Index           =   7
      Left            =   4080
      TabIndex        =   155
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "品名"
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   1920
      TabIndex        =   154
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "⑤"
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   153
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "④"
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   152
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "③"
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   151
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "②"
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   150
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "①"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   149
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "個装資材№"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   148
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "備考"
      Height          =   255
      Index           =   13
      Left            =   8040
      TabIndex        =   147
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "収単/担当者"
      Height          =   255
      Index           =   11
      Left            =   9960
      TabIndex        =   146
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "付加ｸﾗｽ"
      Height          =   255
      Index           =   10
      Left            =   5880
      TabIndex        =   145
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "商品化ｸﾗｽ"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   9
      Left            =   4320
      TabIndex        =   144
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "手配先"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   143
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "標準棚番"
      Height          =   255
      Index           =   7
      Left            =   11520
      TabIndex        =   142
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "数量"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   10560
      TabIndex        =   141
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "品番"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   140
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "仕向け先"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   139
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "承認"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   8130
      TabIndex        =   138
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "担当者"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   5250
      TabIndex        =   137
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "発行日"
      Height          =   255
      Index           =   1
      Left            =   3330
      TabIndex        =   136
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "指図票№"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   135
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "PI000151"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private MTS_CODE    As String * 8
Private SS_CODE     As String * 8
Private CYU_KBN     As String * 1
Private CYU_KBN_N   As String * 1

Private G_Kisyu_F   As Integer
Private L_URIKIN1   As Double
Private L_URIKIN2   As Double
Private L_URIKIN3   As Double

Private SAIKON_F    As String * 1           '再梱包F    2007.11.09

Private TEHAI       As String
    
'--------------------------------------------------- 大阪  部材対応　2012.03.09

Private Type ZAIKO_FUSOKU_T
    JGYOBU          As String * 1
    NAIGAI          As String * 1
    HIN_GAI         As String * 20
    USE_QTY         As Double
    ZAIKO_QTY       As Double

    SAI_QTY         As Double


    IDO_SUMI        As String * 1
    HIKIATE_QTY     As Double


    IDO_SUMI_QTY    As Double               '2012.04.13

End Type


Private ZAIKO_FUSOKU() _
                    As ZAIKO_FUSOKU_T       '在庫不足品番




'--------------------------------------------------- 大阪  部材対応　2012.03.09
    
    
'テキスト用添字
Private Const ptxSHIJI_NO% = 0              '指図票№
'--------------------------------------------------- 大阪  部材対応　2012.03.18
'Private Const ptxORDER_DT% = 1              '受注日
Private Const ptxORDER_NO% = 1              '注文№
'--------------------------------------------------- 大阪  部材対応　2012.03.18
Private Const ptxHAKKO_DT% = 2              '発行日
Private Const ptxTANTO_CODE% = 3            '担当者ｺｰﾄﾞ
Private Const ptxTANTO_NAME% = 4            '担当者名称
Private Const ptxSHONIN_CODE% = 5           '承認者ｺｰﾄﾞ
Private Const ptxSHONIN_NAME% = 6           '承認者名称
Private Const ptxHIN_GAI% = 7               '品番
Private Const ptxHIN_NAME% = 8              '品名
Private Const ptxSHIJI_QTY% = 9             '数量
Private Const ptxST_LOCATION% = 10          '標準棚番
Private Const ptxMI_QTY% = 11               '未商品
Private Const ptxSUMI_QTY% = 12             '商品化済
Private Const ptxUKEHARAI_CODE% = 13        '手配先ｺｰﾄﾞ
Private Const ptxS_CLASS_CODE% = 14         '商品化ｸﾗｽ
Private Const ptxF_CLASS_CODE% = 15         '付加ｸﾗｽ
Private Const ptxN_CLASS_CODE% = 16         '内職ｸﾗｽ

    
Private Const ptxK_HIN_GAI01% = 17          '①　個装資材№
Private Const ptxK_HIN_NAME01% = 18         '①　個装資材名称
Private Const ptxK_QTY01% = 19              '①　員数
Private Const ptxK_SHIJI_QTY01% = 20        '①　数量
Private Const ptxK_ST_LOCATION01% = 21      '①　棚番

Private Const ptxK_HIN_GAI02% = 22          '②　個装資材№
Private Const ptxK_HIN_NAME02% = 23         '②　個装資材名称
Private Const ptxK_QTY02% = 24              '②　員数
Private Const ptxK_SHIJI_QTY02% = 25        '②　数量
Private Const ptxK_ST_LOCATION02% = 26      '②　棚番
    
Private Const ptxK_HIN_GAI03% = 27          '③　個装資材№
Private Const ptxK_HIN_NAME03% = 28         '③　個装資材名称
Private Const ptxK_QTY03% = 29              '③　員数
Private Const ptxK_SHIJI_QTY03% = 30        '③　数量
Private Const ptxK_ST_LOCATION03% = 31      '③　棚番
    
Private Const ptxK_HIN_GAI04% = 32          '④　個装資材№
Private Const ptxK_HIN_NAME04% = 33         '④　個装資材名称
Private Const ptxK_QTY04% = 34              '④　員数
Private Const ptxK_SHIJI_QTY04% = 35        '④　数量
Private Const ptxK_ST_LOCATION04% = 36      '④　棚番
    
Private Const ptxK_HIN_GAI05% = 37          '⑤　個装資材№
Private Const ptxK_HIN_NAME05% = 38         '⑤　個装資材名称
Private Const ptxK_QTY05% = 39              '⑤　員数
Private Const ptxK_SHIJI_QTY05% = 40        '⑤　数量
Private Const ptxK_ST_LOCATION05% = 41      '⑤　棚番
    
    
Private Const ptxG_HIN_GAI01% = 42          '①　外装資材№
Private Const ptxG_HIN_NAME01% = 43         '①　外装資材名称
Private Const ptxG_QTY01% = 44              '①　員数
Private Const ptxG_SHIJI_QTY01% = 45        '①　数量
Private Const ptxG_ST_LOCATION01% = 46      '①　棚番
    
Private Const ptxG_HIN_GAI02% = 47          '②　外装資材№
Private Const ptxG_HIN_NAME02% = 48         '②　外装資材名称
Private Const ptxG_QTY02% = 49              '②　員数
Private Const ptxG_SHIJI_QTY02% = 50        '②　数量
Private Const ptxG_ST_LOCATION02% = 51      '②　棚番
    
Private Const ptxG_HIN_GAI03% = 52          '③　外装資材№
Private Const ptxG_HIN_NAME03% = 53         '③　外装資材名称
Private Const ptxG_QTY03% = 54              '③　員数
Private Const ptxG_SHIJI_QTY03% = 55        '③　数量
Private Const ptxG_ST_LOCATION03% = 56      '③　棚番
    
Private Const ptxD_HIN_GAI01% = 57          '①　同梱／構成品番
Private Const ptxD_HIN_NAME01% = 58         '①　同梱／構成品目
Private Const ptxD_QTY01% = 59              '①　員数
Private Const ptxD_SHIJI_QTY01% = 60        '①　数量
Private Const ptxD_ST_LOCATION01% = 61      '①　棚番
Private Const ptxD_ZAIKO_QTY01% = 62        '①　在庫数
Private Const ptxD_BIKOU01% = 63            '①　備考
    
Private Const ptxD_HIN_GAI02% = 64          '②　同梱／構成品番
Private Const ptxD_HIN_NAME02% = 65         '②　同梱／構成品目
Private Const ptxD_QTY02% = 66              '②　員数
Private Const ptxD_SHIJI_QTY02% = 67        '②　数量
Private Const ptxD_ST_LOCATION02% = 68      '②　棚番
Private Const ptxD_ZAIKO_QTY02% = 69        '②　在庫数
Private Const ptxD_BIKOU02% = 70            '②　備考
    
Private Const ptxD_HIN_GAI03% = 71          '③　同梱／構成品番
Private Const ptxD_HIN_NAME03% = 72         '③　同梱／構成品目
Private Const ptxD_QTY03% = 73              '③　員数
Private Const ptxD_SHIJI_QTY03% = 74        '③　数量
Private Const ptxD_ST_LOCATION03% = 75      '③　棚番
Private Const ptxD_ZAIKO_QTY03% = 76        '③　在庫数
Private Const ptxD_BIKOU03% = 77            '③　備考
    
Private Const ptxD_HIN_GAI04% = 78          '④　同梱／構成品番
Private Const ptxD_HIN_NAME04% = 79         '④　同梱／構成品目
Private Const ptxD_QTY04% = 80              '④　員数
Private Const ptxD_SHIJI_QTY04% = 81        '④　数量
Private Const ptxD_ST_LOCATION04% = 82      '④　棚番
Private Const ptxD_ZAIKO_QTY04% = 83        '④　在庫数
Private Const ptxD_BIKOU04% = 84            '④　備考
    
Private Const ptxD_HIN_GAI05% = 85          '⑤　同梱／構成品番
Private Const ptxD_HIN_NAME05% = 86         '⑤　同梱／構成品目
Private Const ptxD_QTY05% = 87              '⑤　員数
Private Const ptxD_SHIJI_QTY05% = 88        '⑤　数量
Private Const ptxD_ST_LOCATION05% = 89      '⑤　棚番
Private Const ptxD_ZAIKO_QTY05% = 90        '⑤　在庫数
Private Const ptxD_BIKOU05% = 91            '⑤　備考
    
Private Const ptxD_HIN_GAI06% = 92          '⑥　同梱／構成品番
Private Const ptxD_HIN_NAME06% = 93         '⑥　同梱／構成品目
Private Const ptxD_QTY06% = 94              '⑥　員数
Private Const ptxD_SHIJI_QTY06% = 95        '⑥　数量
Private Const ptxD_ST_LOCATION06% = 96      '⑥　棚番
Private Const ptxD_ZAIKO_QTY06% = 97        '⑥　在庫数
Private Const ptxD_BIKOU06% = 98            '⑥　備考
    
    
    
 


'コンボ用添字
Private Const pcmbSHIMUKE% = 0              '仕向け先
Private Const pcmbUKEHARAI% = 1             '手配先
Private Const pcmbS_TANTO% = 2              '収単／担当者コード

Private Const pcmbD_SYUBETSU01% = 3         '①　種別
Private Const pcmbD_SYUBETSU02% = 4         '②　種別
Private Const pcmbD_SYUBETSU03% = 5         '③　種別
Private Const pcmbD_SYUBETSU04% = 6         '④　種別
Private Const pcmbD_SYUBETSU05% = 7         '⑤　種別
Private Const pcmbD_SYUBETSU06% = 8         '⑥　種別


Private Const pcmbD_JGYOBU01% = 3           '①　事業部 2016.01.27
Private Const pcmbD_JGYOBU02% = 4           '②　事業部 2016.01.27
Private Const pcmbD_JGYOBU03% = 5           '③　事業部 2016.01.27
Private Const pcmbD_JGYOBU04% = 6           '④　事業部 2016.01.27
Private Const pcmbD_JGYOBU05% = 7           '⑤　事業部 2016.01.27
Private Const pcmbD_JGYOBU06% = 8           '⑥　事業部 2016.01.27



'チェック用添字
Private Const pchkSAMPLE_F% = 0             '見本作成
Private Const pchkPRI_SHIJI% = 1            '出力対象　指図票
Private Const pchkPRI_PARTS% = 2            '出力対象　ﾊﾟｰﾂﾗﾍﾞﾙ
Private Const pchkPRI_GAISOU% = 3           '出力対象　外装ﾗﾍﾞﾙ
Private Const pchkPRI_KISHU% = 4            '出力対象　機種ﾗﾍﾞﾙ

'ｵﾌﾟｼｮﾝﾎﾞﾀﾝ用添字
Private Const poptSHIJI_NORMAL% = 0         '通常
Private Const poptSHIJI_SPOT% = 1           'スポット
Private Const poptSHIJI_KEPPIN% = 2         '欠品解除
Private Const poptSHIJI_SAIKON% = 3         '再梱包 2007.11.09


'リッチテキスト用添字
Private Const prchBIKOU% = 0                '備考





Private Const cmdMUPDATE% = 3               'ﾏｽﾀ更新

Private Const cmdNext% = 5                  '構成部品画面へ
Private Const cmdCen% = 10                  '取り消し


'Private Const LAST_UPDATE_DAY$ = "([PI00015] 2017.10.17 09:30)"
'Private Const LAST_UPDATE_DAY$ = "([PI00015] 2017.12.15 10:15)"
Private Const LAST_UPDATE_DAY$ = "([PI00015] 2020.05.07 12:30) 作業実績項目名変更"


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PI000151.MousePointer = vbHourglass

    Call Ctrl_Lock(PI000151)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PI000151)


    PI000151.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer, chk As Integer, flg As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts         As Integer
    
    
Dim Mi_Qty      As Long
Dim Sumi_Qty    As Long
    
Dim i           As Integer
Dim j           As Integer
Dim k           As Integer
    
Dim yn          As Integer
    
Dim wkJgyobu    As String * 1
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        Case ptxSHIJI_NO    '指図票№
        
            If Text1(ptxSHIJI_NO).Locked Then
            Else
            
            
                    
            
            
                If Trim(Text1(ptxSHIJI_NO).text) = "" Then
                Else
                    
                    If IsNumeric(Text1(ptxSHIJI_NO).text) Then
                        Text1(ptxSHIJI_NO).text = Format(CLng(Text1(ptxSHIJI_NO).text), "00000000")
                    End If
                    
                    '指図票ﾃﾞｰﾀのﾁｪｯｸ
                    sts = P_SSHIJI_Read_Proc()
                    Select Case sts
                        Case False, BtNoErr
''                            If CLng(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) <> 0 Then
''                                MsgBox "受入完了もしくは受入中です。この画面では処理できません"
''                                Text1(Mode).SetFocus
''                                Exit Function
''                            End If
                            If CLng(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) <> 0 Then
                                yn = MsgBox("受入完了もしくは受入中です。" & Chr(13) & Chr(10) & _
                                        "強制編集する場合は、「はい」をクリック。", vbYesNo + vbDefaultButton2, "確認入力")
                                If yn = vbNo Then
                                    Text1(Mode).SetFocus
                                    Exit Function
                                End If
                            End If
                            
                            If StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
                                MsgBox "キャンセル済です。この画面では処理できません"
                                Text1(Mode).SetFocus
                                Exit Function
                            End If
                        
                        Case BtErrKeyNotFound
                            MsgBox "入力した項目はエラーです。"
                            Text1(Mode).SetFocus
                            Exit Function
                        Case Else
                            Exit Function
                    End Select
                End If
            
            
            
                Text1(ptxSHIJI_NO).BackColor = G_INPUT_NG
                Text1(ptxSHIJI_NO).Locked = True
                Text1(ptxSHIJI_NO).TabStop = False
            
            
            
            End If
        
        
        Case ptxORDER_NO    '受注日     2012.03.18 受注№に変更
            
            If chk = 1 Then
            Else
                
'--------------------------------------------------- 大阪  部材対応　2012.03.18
'                If Trim(Text1(ptxORDER_DT).text) = "" Then
'                Else
'
'                    If Not IsDate(Text1(ptxORDER_DT).text) Then
'                        MsgBox "入力した項目はエラーです。(受注日)"
'                        Text1(Mode).SetFocus
'                        Exit Function
'                    Else
'                        Text1(ptxORDER_DT).text = Format(CDate(Text1(ptxORDER_DT).text), "YYYY/MM/DD")
'                    End If
'                End If
            
            
                If Trim(Text1(ptxORDER_NO).text) = "" Then
                    Call Disp_Lock_Proc(False)
                Else
                    Call UniCode_Conv(K6_ODR_ORDER.ORDER_NO, Text1(ptxORDER_NO).text)
                
                    sts = BTRV(BtOpGetEqual, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K6_ODR_ORDER, Len(K6_ODR_ORDER), 6)
                    Select Case sts
                        Case BtNoErr
                        
                            If Trim(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode)) <> "" Then
                                MsgBox "入力した項目はエラーです。(親品番注文№:完了済みです)"
                                Text1(Mode).SetFocus
                                Exit Function
                            End If
                        
                            Text1(ptxHIN_GAI).text = RTrim(StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode))
                        
                        
                            Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                            Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)
                    
                    
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                                    
                                    Text1(ptxHIN_NAME).text = ""
                                    Text1(ptxST_LOCATION).text = ""
                                    Text1(ptxMI_QTY).text = ""
                                    Text1(ptxSUMI_QTY).text = ""
                            
                                    MsgBox "入力した項目はエラーです。(品番)"
                                    Text1(Mode).SetFocus
                                    Exit Function
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                    Exit Function
                            
                            End Select
                        
                        
                        
                            Text1(ptxHIN_NAME).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                            
                            Text1(ptxSHIJI_QTY).text = Val(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode))
                            
                            
                            If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                                Text1(ptxST_LOCATION).text = ""
                            Else
                                Text1(ptxST_LOCATION).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                                StrConv(ITEMREC.ST_DAN, vbUnicode)
                            End If
                
                            '>>>>>>>>>>>>>>>>>> 2013.01.07
                            'If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI Then
                            '    wkJgyobu = BUZAI
                            'Else
                            '    'wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)  '2012.04.04
                            '    wkJgyobu = YUKO_JGYOBU                          '2012.04.04
                            'End If
                            wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
                            '>>>>>>>>>>>>>>>>>> 2013.01.07


                            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, wkJgyobu, _
                                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                                                    , , , Jyogai_Soko_umu) Then
                                Exit Function

                            End If

                            Text1(ptxMI_QTY).text = Format(Mi_Qty, "#0")
                            Text1(ptxSUMI_QTY).text = Format(Sumi_Qty, "#0")

                            If flg = 1 Then
                            Else

                                If Trim(Text1(ptxSHIJI_NO).text) = "" Then
                                    If StrConv(P_COMPO_O_REC.SHIMUKE_CODE, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) And _
                                        StrConv(P_COMPO_O_REC.JGYOBU, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) And _
                                        StrConv(P_COMPO_O_REC.NAIGAI, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) And _
                                        Trim(StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode)) = Trim(Text1(ptxHIN_GAI).text) Then
                                    Else
                                        sts = P_COMPO_Disp_Proc()
                                        Select Case sts
                                            Case BtNoErr
                                            Case BtErrKeyNotFound
                    '                            MsgBox "入力した項目はエラーです。"
                    '                            Text1(Mode).SetFocus
                    '                            Exit Function
                                            Case Else
                                                Call File_Error(sts, BtOpGetEqual, "構成マスタ")
                                                Exit Function
                                        End Select
                                            
                                    End If
                                Else
                                    If StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) And _
                                        StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) And _
                                        StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) And _
                                        Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)) = Trim(Text1(ptxHIN_GAI).text) Then
                                    Else
                                        sts = P_COMPO_Disp_Proc()
                                        Select Case sts
                                            Case BtNoErr
                                            Case BtErrKeyNotFound
                    '                            MsgBox "入力した項目はエラーです。"
                    '                            Text1(Mode).SetFocus
                    '                            Exit Function
                                            Case Else
                                                Call File_Error(sts, BtOpGetEqual, "構成マスタ")
                                                Exit Function
                                        End Select
                                    End If
                                End If
                            End If
                        
                        
                                             '個装資材再計算
                        For i = ptxK_QTY01 To ptxK_QTY05 Step 5
                        
                            If IsNumeric(Text1(i).text) Then
                                Text1(i + 1).text = Format(CDbl(CLng(Text1(ptxSHIJI_QTY).text) * CDbl(Text1(i).text)), "#0.00")
                            Else
                                Text1(i + 1).text = ""
                            End If
                        Next i
                    
                    
                        '外装資材再計算
                        For i = ptxG_QTY01 To ptxG_QTY03 Step 5
                        
                            If IsNumeric(Text1(i).text) Then
                                Text1(i + 1).text = Format(Int(CDbl(CLng(Text1(ptxSHIJI_QTY).text) / CDbl(Text1(i).text))), "#0")
                            Else
                                Text1(i + 1).text = ""
                            End If
                        Next i
                    
                        '同梱／構成再計算
                        For i = ptxD_QTY01 To ptxD_QTY06 Step 7
                        
                            k = 0
                            j = Mode - ptxD_QTY01
                            Do
                                j = j - 5
                                If j < 0 Then
                                    Exit Do
                                End If
                                k = k + 1
                            Loop
                        
                            
                            
                            
                            If IsNumeric(Text1(i).text) Then
                                Text1(i + 1).text = Format(CDbl(CLng(Text1(ptxSHIJI_QTY).text) * CDbl(Text1(i).text)), "#0.00")
                            
                            
                            Else
                                Text1(i + 1).text = ""
                            
                            
                            
                            End If
                        Next i
                    
                    
                        For i = 0 To UBound(D_Item_Tbl)
                            If Trim(D_Item_Tbl(i).JGYOBU) = "" Then
                            Else
                                D_Item_Tbl(i).SHIJI_QTY = CDbl(Text1(ptxSHIJI_QTY).text) * D_Item_Tbl(i).QTY
                            End If
                        Next i
   
                        
                        
                        
                        Case BtErrKeyNotFound
                            MsgBox "入力した項目はエラーです。(親品番注文№)"
                            Text1(Mode).SetFocus
                            Exit Function
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "親品番注文F")
                            Exit Function
                    
                    End Select
                
                    Call Disp_Lock_Proc(True)
                
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    再発行時の警告  2012.04.13
                    If StrConv(ODR_ORDER_REC.PRT_FLG, vbUnicode) = "F" Then
                        yn = MsgBox("指図票発行済みです。処理を継続しますか？", vbYesNo + vbDefaultButton2, "確認入力")
                        If yn = vbNo Then
                            Call Disp_Lock_Proc(False)
                            Exit Function
                        End If
                    End If
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    再発行時の警告  2012.04.13
                
                
                
                End If
'--------------------------------------------------- 大阪  部材対応　2012.03.18
            End If
        
        
        
        Case ptxHAKKO_DT    '発行日
            
            If chk = 1 Then
            Else
                If Trim(Text1(ptxHAKKO_DT).text) = "" Then
                Else
                    If Not IsDate(Text1(ptxHAKKO_DT).text) Then
                        MsgBox "入力した項目はエラーです。(発行日)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Else
                        Text1(ptxHAKKO_DT).text = Format(CDate(Text1(ptxHAKKO_DT).text), "YYYY/MM/DD")
                    End If
                End If
            End If
        
        Case ptxTANTO_CODE      '担当者
        
            If chk = 1 Then
            Else
                Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTANTO_CODE).text)
    
                sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
                Select Case sts
                    Case BtNoErr
                        Text1(ptxTANTO_NAME).text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                    Case BtErrKeyNotFound
                        Text1(ptxTANTO_NAME).text = ""
                
                        MsgBox "入力した項目はエラーです。(担当者)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                        Exit Function
                
                End Select
            End If
    
        Case ptxSHONIN_CODE     '承認者
        
            If chk = 1 Then
            Else
                Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxSHONIN_CODE).text)
    
                sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
                Select Case sts
                    Case BtNoErr
                        Text1(ptxSHONIN_NAME).text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                    Case BtErrKeyNotFound
                        Text1(ptxSHONIN_NAME).text = ""
                
                        MsgBox "入力した項目はエラーです。(承認者)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                        Exit Function
                    
                
                
                End Select
            End If
        Case ptxHIN_GAI         '品番
    
                    
            Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
            Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)
    
    
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    
                    Text1(ptxHIN_NAME).text = ""
                    Text1(ptxST_LOCATION).text = ""
                    Text1(ptxMI_QTY).text = ""
                    Text1(ptxSUMI_QTY).text = ""
            
                    MsgBox "入力した項目はエラーです。(品番)"
                    Text1(Mode).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            
            End Select
            
            Text1(ptxHIN_NAME).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
            
            If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                Text1(ptxST_LOCATION).text = ""
            Else
                Text1(ptxST_LOCATION).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_DAN, vbUnicode)
            End If

            '>>>>>>>>>>>>>> 2013.01.07
            'If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI Then
            '    wkJgyobu = BUZAI
            'Else
            '    'wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)  2012.04.04
            '    wkJgyobu = YUKO_JGYOBU                          '2012.04.04
            'End If
            wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
            '>>>>>>>>>>>>>> 2013.01.07

            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, wkJgyobu, _
                                                   StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                   StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                                   , , , Jyogai_Soko_umu) Then
                Exit Function
            
            End If

            Text1(ptxMI_QTY).text = Format(Mi_Qty, "#0")
            Text1(ptxSUMI_QTY).text = Format(Sumi_Qty, "#0")
    
            
            If flg = 1 Then
            Else
            
                If Trim(Text1(ptxSHIJI_NO).text) = "" Then
                    If StrConv(P_COMPO_O_REC.SHIMUKE_CODE, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) And _
                        StrConv(P_COMPO_O_REC.JGYOBU, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) And _
                        StrConv(P_COMPO_O_REC.NAIGAI, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) And _
                        Trim(StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode)) = Trim(Text1(ptxHIN_GAI).text) Then
                    Else
                        sts = P_COMPO_Disp_Proc()
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
    '                            MsgBox "入力した項目はエラーです。"
    '                            Text1(Mode).SetFocus
    '                            Exit Function
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "構成マスタ")
                                Exit Function
                        End Select
                            
                    End If
                Else
                    If StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) And _
                        StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) And _
                        StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) And _
                        Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)) = Trim(Text1(ptxHIN_GAI).text) Then
                    Else
                        sts = P_COMPO_Disp_Proc()
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
    '                            MsgBox "入力した項目はエラーです。"
    '                            Text1(Mode).SetFocus
    '                            Exit Function
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "構成マスタ")
                                Exit Function
                        End Select
                    End If
                End If
            End If
        Case ptxSHIJI_QTY       '数量
    
            If chk = 1 Then
            Else
                If Not IsNumeric(Text1(ptxSHIJI_QTY).text) Then
                    MsgBox "入力した項目はエラーです。(数量)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxSHIJI_QTY).text = Format(CLng(Text1(ptxSHIJI_QTY).text), "#0")
                
                    '個装資材再計算
                    For i = ptxK_QTY01 To ptxK_QTY05 Step 5
                    
                        If IsNumeric(Text1(i).text) Then
                            Text1(i + 1).text = Format(CDbl(CLng(Text1(ptxSHIJI_QTY).text) * CDbl(Text1(i).text)), "#0.00")
                        Else
                            Text1(i + 1).text = ""
                        End If
                    Next i
                
                
                    '外装資材再計算
                    For i = ptxG_QTY01 To ptxG_QTY03 Step 5
                    
                        If IsNumeric(Text1(i).text) Then
                            Text1(i + 1).text = Format(Int(CDbl(CLng(Text1(ptxSHIJI_QTY).text) / CDbl(Text1(i).text))), "#0")
                        Else
                            Text1(i + 1).text = ""
                        End If
                    Next i
                
                    '同梱／構成再計算
                    For i = ptxD_QTY01 To ptxD_QTY06 Step 7
                    
                        k = 0
                        j = Mode - ptxD_QTY01
                        Do
                            j = j - 5
                            If j < 0 Then
                                Exit Do
                            End If
                            k = k + 1
                        Loop
                    
                        
                        
                        
                        If IsNumeric(Text1(i).text) Then
                            Text1(i + 1).text = Format(CDbl(CLng(Text1(ptxSHIJI_QTY).text) * CDbl(Text1(i).text)), "#0.00")
                        
                        
                        Else
                            Text1(i + 1).text = ""
                        
                        
                        
                        End If
                    Next i
                
                
                    For i = 0 To UBound(D_Item_Tbl)
                        If Trim(D_Item_Tbl(i).JGYOBU) = "" Then
                        Else
                            D_Item_Tbl(i).SHIJI_QTY = CDbl(Text1(ptxSHIJI_QTY).text) * D_Item_Tbl(i).QTY
                        End If
                    Next i
                End If
            End If
    
        Case ptxUKEHARAI_CODE   '手配先
            
            If chk = 1 Then
            Else
               Combo1(pcmbUKEHARAI).ListIndex = -1
               For i = 0 To Combo1(pcmbUKEHARAI).ListCount - 1
                   If Text1(ptxUKEHARAI_CODE).text = Trim(Right(Combo1(pcmbUKEHARAI).List(i), 5)) Then
                       Combo1(pcmbUKEHARAI).ListIndex = i
                       Exit For
                   End If
               
               Next i
        
               If i > Combo1(pcmbUKEHARAI).ListCount - 1 Then
                   MsgBox "入力した項目はエラーです。(手配先)"
                   Text1(Mode).SetFocus
                   Exit Function
               End If
            End If
    
    
    
        Case ptxS_CLASS_CODE    '商品化ｸﾗｽ
        
            If Trim(Text1(ptxS_CLASS_CODE).text) = "" Then
            Else
            
            
                Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 1, 2))
                Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxS_CLASS_CODE).text)
    
                sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                
                        MsgBox "入力した項目はエラーです。(商品化ｸﾗｽ)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "商品化ｸﾗｽ")
                        Exit Function
                
                End Select
            End If
        Case ptxF_CLASS_CODE    '付加ｸﾗｽ
        
            If Trim(Text1(ptxF_CLASS_CODE).text) = "" Then
            Else
                Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 1, 2))
                Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxF_CLASS_CODE).text)
    
                sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                
                        MsgBox "入力した項目はエラーです。(付加ｸﾗｽ)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "商品化ｸﾗｽ")
                        Exit Function
                
                End Select
            End If
    
        Case ptxN_CLASS_CODE    '内職ｸﾗｽ
        
            If Trim(Text1(ptxN_CLASS_CODE).text) = "" Then
            Else
                Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 1, 2))
                Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxN_CLASS_CODE).text)
    
                sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                
                        MsgBox "入力した項目はエラーです。(内職ｸﾗｽ)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "商品化ｸﾗｽ")
                        Exit Function
                
                End Select
            End If
    
                                '個装資材№
        Case ptxK_HIN_GAI01, ptxK_HIN_GAI02, ptxK_HIN_GAI03, ptxK_HIN_GAI04, ptxK_HIN_GAI05
            If Trim(Text1(Mode).text) = "" Then
                Text1(Mode + 1).text = ""
                Text1(Mode + 2).text = ""
                Text1(Mode + 3).text = ""
                Text1(Mode + 4).text = ""
            Else
                
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2014.03.24
'                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
'                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)
'
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                Select Case sts
'                    Case BtNoErr
'                    Case BtErrKeyNotFound
'                        '資材品で読み替え
'
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)
'
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                        Select Case sts
'                            Case BtNoErr
'                            Case BtErrKeyNotFound
'
'                                If HIN_INV Then
'                                    '未登録品番　可　資材としておく
'                                    Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
'                                    Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
'                                    Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番")
'                                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
'                                Else
'                                    MsgBox "入力した項目はエラーです。(個装資材　品番)"
'                                    Text1(Mode).SetFocus
'                                    Exit Function
'                                End If
'                            Case Else
'                                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
'                                Exit Function
'
'                        End Select
'
'                    Case Else
'                        Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
'                        Exit Function
'
'                End Select

                sts = Item_Read_Proc(Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1), Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1), Text1(Mode).text)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        MsgBox "入力した項目はエラーです。(個装資材　品番)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Exit Function
                End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2014.03.24
    
                '品名
                Text1(Mode + 1).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                '標準棚番
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                Else
                    Text1(Mode + 4).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
                
                i = 0
                j = Mode - ptxK_HIN_GAI01
                Do
                    j = j - 5
                    If j < 0 Then
                        Exit Do
                    End If
                    i = i + 1
                Loop
            
                K_Item_Tbl(i).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                K_Item_Tbl(i).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
            
            
            End If
                                '個装資材　員数
        Case ptxK_QTY01, ptxK_QTY02, ptxK_QTY03, ptxK_QTY04, ptxK_QTY05
            
            If Trim(Text1(Mode).text) = "" Then
                If Trim(Text1(Mode - 1).text) <> "" Then
                    MsgBox "入力した項目はエラーです。(個装資材　員数)"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            Else
                If Trim(Text1(Mode - 1).text) = "" Then
                    MsgBox "入力した項目はエラーです。(個装資材　員数)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    If Not IsNumeric(Text1(Mode).text) Then
                        MsgBox "入力した項目はエラーです。(個装資材　員数)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Else
                        Text1(Mode).text = Format(CDbl(Text1(Mode).text), "#0.00")
                        '数量
                        If IsNumeric(Text1(ptxSHIJI_QTY).text) Then
                            Text1(Mode + 1).text = Format(CDbl(CLng(Text1(ptxSHIJI_QTY).text) * CDbl(Text1(Mode).text)), "#0.00")
                        
                        
                        
                        Else
                            Text1(Mode + 1).text = ""
                        End If
                    
                    End If
                End If
            End If
    
    
    
    
    
                                '外装資材№
        Case ptxG_HIN_GAI01, ptxG_HIN_GAI02, ptxG_HIN_GAI03
            If Trim(Text1(Mode).text) = "" Then
                Text1(Mode + 1).text = ""
                Text1(Mode + 2).text = ""
                Text1(Mode + 3).text = ""
                Text1(Mode + 4).text = ""
            Else
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2014.03.24
'                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
'                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)
'
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                Select Case sts
'                    Case BtNoErr
'                    Case BtErrKeyNotFound
'                        '資材品で読み替え
'
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)
'
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                        Select Case sts
'                            Case BtNoErr
'                            Case BtErrKeyNotFound
'
'                                If HIN_INV Then
'                                    '未登録品番　可　資材としておく
'                                    Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
'                                    Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
'                                    Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番")
'                                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
'                                Else
'
'                                    MsgBox "入力した項目はエラーです。(外装資材　品番)"
'                                    Text1(Mode).SetFocus
'                                    Exit Function
'                                End If
'                            Case Else
'                                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
'                                Exit Function
'
'                        End Select
'
'                    Case Else
'                        Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
'                        Exit Function
'
'                End Select
                sts = Item_Read_Proc(Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1), Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1), Text1(Mode).text)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        MsgBox "入力した項目はエラーです。(外装資材　品番)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Exit Function
                End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2014.03.24
    
                '品名
                Text1(Mode + 1).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                '標準棚番
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                Else
                    Text1(Mode + 4).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            
                i = 0
                j = Mode - ptxG_HIN_GAI01
                Do
                    j = j - 5
                    If j < 0 Then
                        Exit Do
                    End If
                    i = i + 1
                Loop
            
                G_Item_Tbl(i).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                G_Item_Tbl(i).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
            
            
            End If
                                '外装資材　員数
        Case ptxG_QTY01, ptxG_QTY02, ptxG_QTY03
            
            If Trim(Text1(Mode).text) = "" Then
                If Trim(Text1(Mode - 1).text) <> "" Then
                    MsgBox "入力した項目はエラーです。(外装資材　員数)"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            Else
                If Trim(Text1(Mode - 1).text) = "" Then
                    MsgBox "入力した項目はエラーです。(外装資材　員数)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    If Not IsNumeric(Text1(Mode).text) Then
                        MsgBox "入力した項目はエラーです。(外装資材　員数)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Else
                        Text1(Mode).text = Format(CDbl(Text1(Mode).text), "#0.00")
                        '数量
                        If IsNumeric(Text1(ptxSHIJI_QTY).text) Then
                            Text1(Mode + 1).text = Format(Int(CDbl(CLng(Text1(ptxSHIJI_QTY).text) / CDbl(Text1(Mode).text))), "#0")
                        
                        Else
                            Text1(Mode + 1).text = ""
                        End If
                    
                    End If
                End If
            End If
    
                                '同梱／構成　品番
        Case ptxD_HIN_GAI01, ptxD_HIN_GAI02, ptxD_HIN_GAI03, ptxD_HIN_GAI04, ptxD_HIN_GAI05, ptxD_HIN_GAI06
            If Trim(Text1(Mode).text) = "" Then
                Text1(Mode + 1).text = ""
                Text1(Mode + 2).text = ""
                Text1(Mode + 3).text = ""
                Text1(Mode + 4).text = ""
                Text1(Mode + 5).text = ""
                Text1(Mode + 6).text = ""
            
            
                i = 0
                j = Mode - ptxD_HIN_GAI01
                Do
                    j = j - 7
                    If j < 0 Then
                        Exit Do
                    End If
                    i = i + 1
                Loop
            
                D_Item_Tbl(i).JGYOBU = ""
                D_Item_Tbl(i).NAIGAI = ""
                D_Item_Tbl(i).HIN_GAI = ""
                D_Item_Tbl(i).QTY = 0
                D_Item_Tbl(i).SHIJI_QTY = 0
                D_Item_Tbl(i).BIKOU = ""
            
            
            Else
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2014.03.24
'                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
'                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)
'
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                Select Case sts
'                    Case BtNoErr
'                    Case BtErrKeyNotFound
'
'                        '品番（内）で読み替え
'                        Call UniCode_Conv(K2_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
'                        Call UniCode_Conv(K2_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
'                        Call UniCode_Conv(K2_ITEM.HIN_NAI, Text1(Mode).text)
'
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
'                        Select Case sts
'                            Case BtNoErr
'                            Case BtErrKeyNotFound
'
'
'
'
'
'
'
'                                '資材品で読み替え
'                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
'
'                                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
'                                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)
'
'                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                                Select Case sts
'                                    Case BtNoErr
'                                    Case BtErrKeyNotFound
'
'                                        If HIN_INV Then
'                                            '未登録品番　可　資材としておく
'                                            Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
'                                            Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
'                                            Call UniCode_Conv(ITEMREC.HIN_GAI, Text1(Mode).text)
'                                            Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番")
'                                            Call UniCode_Conv(ITEMREC.ST_SOKO, "")
'
'                                        Else
'
'                                            MsgBox "入力した項目はエラーです。(同梱／構成　品番)"
'                                            Text1(Mode).SetFocus
'                                            Exit Function
'                                        End If
'                                    Case Else
'                                        Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
'                                        Exit Function
'
'                                End Select
'
'                            Case Else
'                                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
'                                Exit Function
'                       End Select
'
'                    Case Else
'                        Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
'                        Exit Function
'
'                End Select
                
                
                
'>>>>>>>>>>>>>>>>>>>>>>>>>  品番読込み変更 2016.01.27
'                sts = Item_Read_Proc(Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1), Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1), Text1(Mode).text)
                
                Select Case Mode
                    Case ptxD_HIN_GAI01
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo2(pcmbD_JGYOBU01).text, 1))
                    Case ptxD_HIN_GAI02
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo2(pcmbD_JGYOBU02).text, 1))
                    Case ptxD_HIN_GAI03
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo2(pcmbD_JGYOBU03).text, 1))
                    Case ptxD_HIN_GAI04
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo2(pcmbD_JGYOBU04).text, 1))
                    Case ptxD_HIN_GAI05
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo2(pcmbD_JGYOBU05).text, 1))
                    Case ptxD_HIN_GAI06
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo2(pcmbD_JGYOBU06).text, 1))
                End Select
                
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'>>>>>>>>>>>>>>>>>>>>>>>>>  品番読込み変更 2016.01.27
                
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        MsgBox "入力した項目はエラーです。(同梱／構成　品番)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Exit Function
                End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2014.03.24
    
                '品名
                Text1(Mode + 1).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
'                '標準棚番
'                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
'                Else
'                    Text1(Mode + 4).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
'                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
'                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
'                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
'                End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
            
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.01.07
                '在庫数
                'If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI Then
                '    wkJgyobu = BUZAI
                'Else
                '    'wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)  2012.04.04
                '    wkJgyobu = YUKO_JGYOBU                          '2012.04.04
                'End If
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2016.01.27
'                Select Case StrConv(ITEMREC.JGYOBU, vbUnicode)
'                    Case SHIZAI
'                        wkJgyobu = BUZAI
'                    Case SETSUBI
'                        wkJgyobu = YUKO_JGYOBU
'                    Case Else
'                        wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
'                End Select

                wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2016.01.27

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.01.07
                
                '標準棚番を設定 2013.01.11
                If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, wkJgyobu, _
                                                        StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                        StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                                        (StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_DAN, vbUnicode)), _
                                                        , , Jyogai_Soko_umu) Then
                    Exit Function
                
                End If
            
                Text1(Mode + 5).text = Format(Sumi_Qty + Mi_Qty, "#0")
            
            
            
                i = 0
                j = Mode - ptxD_HIN_GAI01
                Do
                    j = j - 7
                    If j < 0 Then
                        Exit Do
                    End If
                    i = i + 1
                Loop
            
' 2013.01.07 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                D_Item_Tbl(i).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
'                D_Item_Tbl(i).JGYOBU = BUZAI            '同梱/構成の事業部を「部材」固定に変更
' 2013.01.07 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                D_Item_Tbl(i).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
                D_Item_Tbl(i).HIN_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
            
            
            
            
            
            
            
            
'>>>>>>>>>>>>>>>>>>>>>  2013.01.07 標準棚番の表示を資材--＞部材ｾﾝﾀｰ
                
'>>>>>>>>>> 廃止　2016.01.27
'                Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
'
'                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)
'
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                Select Case sts
'                    Case BtNoErr
'                    Case BtErrKeyNotFound
'                        Call UniCode_Conv(ITEMREC.ST_SOKO, "")
'                    Case Else
'                        Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
'                        Exit Function
'
'                End Select
'
'                '標準棚番
'                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
'                Else
'                    Text1(Mode + 4).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
'                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
'                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
'                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
'                End If
'>>>>>>>>>> 廃止　2016.01.27


'>>>>>>>>>>>>>>>>>>>>>  2013.01.07

            
            
            End If
                                '同梱／構成　員数
        Case ptxD_QTY01, ptxD_QTY02, ptxD_QTY03, ptxD_QTY04, ptxD_QTY05, ptxD_QTY06
            
            If Trim(Text1(Mode).text) = "" Then
                If Trim(Text1(Mode - 2).text) <> "" Then
                    MsgBox "入力した項目はエラーです。(同梱／構成　員数)"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            Else
                If Trim(Text1(Mode - 2).text) = "" Then
                    MsgBox "入力した項目はエラーです。(同梱／構成　員数)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    If Not IsNumeric(Text1(Mode).text) Then
                        MsgBox "入力した項目はエラーです。(同梱／構成　員数)"
                        Text1(Mode).SetFocus
                        Exit Function
                    Else
                        Text1(Mode).text = Format(CDbl(Text1(Mode).text), "#0.00")
                        
                        
                        i = 0
                        j = Mode - ptxD_QTY01
                        Do
                            j = j - 7
                            If j < 0 Then
                                Exit Do
                            End If
                            i = i + 1
                        Loop
                    
                        D_Item_Tbl(i).QTY = CDbl(Text1(Mode).text)
                        
                        
                        '数量
                        If IsNumeric(Text1(ptxSHIJI_QTY).text) Then
                            Text1(Mode + 1).text = Format(CDbl(CLng(Text1(ptxSHIJI_QTY).text) * CDbl(Text1(Mode).text)), "#0.00")
                            D_Item_Tbl(i).SHIJI_QTY = CDbl(Text1(Mode + 1).text)
                        Else
                            Text1(Mode + 1).text = ""
                            D_Item_Tbl(i).SHIJI_QTY = 0
                        End If
                    
                    End If
                End If
            End If
                                '同梱／構成　備考
        Case ptxD_BIKOU01, ptxD_BIKOU02, ptxD_BIKOU03, ptxD_BIKOU04, ptxD_BIKOU05, ptxD_BIKOU06
            If Trim(Text1(Mode).text) <> "" Then
                If Trim(Text1(Mode - 6).text) = "" Then
                    MsgBox "入力した項目はエラーです。(同梱／構成　備考)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    i = 0
                    j = Mode - ptxD_BIKOU01
                    Do
                        j = j - 7
                        If j < 0 Then
                            Exit Do
                        End If
                        i = i + 1
                    Loop
                
                    D_Item_Tbl(i).BIKOU = Text1(Mode).text
                End If
            End If
    End Select
        
        
    Error_Check_Proc = False
    

End Function

Private Function Item_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   画面表示
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim k           As Integer
Dim g           As Integer
Dim d           As Integer

Dim K_Index     As Integer
Dim G_Index     As Integer
Dim DT_Index    As Integer
Dim DC_Index    As Integer


Dim Mi_Qty      As Long
Dim Sumi_Qty    As Long

Dim wkJgyobu    As String * 1


Dim Wk_LOC      As String   '2013.01.07


    Item_Disp_Proc = True
    
    For i = ptxK_HIN_GAI01 To ptxD_BIKOU06
        Text1(i).text = ""
    Next i
            
    For i = pcmbD_JGYOBU01 To pcmbD_JGYOBU06        '2016.04.07
        Combo2(i).ListIndex = -1                    '2016.04.07
    Next                                            '2016.04.07
            
            
    '出力対象
    Erase K_Item_Tbl
    Erase G_Item_Tbl
    Erase D_Item_Tbl
    
    ReDim K_Item_Tbl(0 To 4)
    ReDim G_Item_Tbl(0 To 2)
    ReDim D_Item_Tbl(0 To 49)

    For i = 0 To UBound(K_Item_Tbl)
        K_Item_Tbl(i).JGYOBU = ""
        K_Item_Tbl(i).NAIGAI = ""
    Next i

    For i = 0 To UBound(G_Item_Tbl)
        G_Item_Tbl(i).JGYOBU = ""
        G_Item_Tbl(i).NAIGAI = ""
    Next i

    For i = 0 To UBound(D_Item_Tbl)
        D_Item_Tbl(i).JGYOBU = ""
        D_Item_Tbl(i).NAIGAI = ""
        D_Item_Tbl(i).HIN_GAI = ""
    
    Next i
    
    Text1(ptxS_CLASS_CODE).text = ""
    Text1(ptxF_CLASS_CODE).text = ""
    Text1(ptxN_CLASS_CODE).text = ""
    
    
    
    '--------------------------------   「親」情報
        
    
    Text1(ptxSHIJI_NO).text = StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode)           '指図票№
                                                                                    
                                                                                    '受注日
    
    If Trim(StrConv(P_SSHIJI_O_REC.ORDER_DT, vbUnicode)) = "" Then
        Text1(ptxORDER_NO).text = ""
    Else
'---------------------------    2012.03.27
'        Text1(ptxORDER_NO).text = Mid(StrConv(P_SSHIJI_O_REC.ORDER_DT, vbUnicode), 1, 4) & "/" & _
'                                    Mid(StrConv(P_SSHIJI_O_REC.ORDER_DT, vbUnicode), 5, 2) & "/" & _
'                                    Mid(StrConv(P_SSHIJI_O_REC.ORDER_DT, vbUnicode), 7, 2)

        Text1(ptxORDER_NO).text = StrConv(P_SSHIJI_O_REC.ORDER_DT, vbUnicode) & StrConv(P_SSHIJI_O_REC.ORDER_DT_SEQ, vbUnicode)
'---------------------------    2012.03.27
    End If
                                                                                    
                                                                                    
                                                                                    
                                                                                    '発行日
    
    If Trim(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode)) = "" Then
        Text1(ptxHAKKO_DT).text = ""
    Else
    
        Text1(ptxHAKKO_DT).text = Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 7, 2)
    
    End If
    
    
    Text1(ptxTANTO_CODE).text = StrConv(P_SSHIJI_O_REC.TANTO_CODE, vbUnicode)       '担当者ｺｰﾄﾞ／名称
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTANTO_CODE).text)

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxTANTO_NAME).text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text1(ptxTANTO_NAME).text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Exit Function
    
    End Select
    
    Text1(ptxSHONIN_CODE).text = StrConv(P_SSHIJI_O_REC.SHONIN_CODE, vbUnicode)     '承認者ｺｰﾄﾞ／名称
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxSHONIN_CODE).text)

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxSHONIN_NAME).text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text1(ptxSHONIN_NAME).text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Exit Function
    
    End Select
    
    For i = 0 To Combo1(pcmbSHIMUKE).ListCount - 1                                  '仕向け先ｺｰﾄﾞ
    
        If StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE).List(i), 4), 1, 2) Then
            Combo1(pcmbSHIMUKE).ListIndex = i
            Exit For
        End If
    
    Next i
    
    
    Text1(ptxHIN_GAI).text = Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))       '品番／品名／標準棚番／未商品／商品化済
        
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxHIN_NAME).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
            Text1(ptxST_LOCATION).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
        
        
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
'            If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI Then
'                wkJgyobu = BUZAI
'            Else
'                'wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)  2012.04.04
'                wkJgyobu = YUKO_JGYOBU                          '2012.04.04
'            End If

            wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
        
            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, wkJgyobu, _
                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), , , , Jyogai_Soko_umu) Then
                Exit Function
            
            End If

            Text1(ptxMI_QTY).text = Format(Mi_Qty, "#0")
            Text1(ptxSUMI_QTY).text = Format(Sumi_Qty, "#0")
        
        
        
        Case BtErrKeyNotFound
            Text1(ptxHIN_NAME).text = ""
            Text1(ptxST_LOCATION).text = ""
            Text1(ptxMI_QTY).text = ""
            Text1(ptxSUMI_QTY).text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    
    End Select
                                                                                    '指示数量
    Text1(ptxSHIJI_QTY).text = Format(CLng(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode)), "#0")
    
    Text1(ptxHIN_GAI).text = Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))       '品番／品名／標準棚番／未商品／商品化済
                                                                                        
    Text1(ptxUKEHARAI_CODE).text = Trim(StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode))   '手配先
    For i = 0 To Combo1(pcmbUKEHARAI).ListCount - 1
    
        If Text1(ptxUKEHARAI_CODE).text = Trim(Right(Combo1(pcmbUKEHARAI).List(i), 5)) Then
            Combo1(pcmbUKEHARAI).ListIndex = i
            Exit For
        End If
    
    Next i
    
    Text1(ptxS_CLASS_CODE).text = Trim(StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode)) '商品化ｸﾗｽ
    Text1(ptxF_CLASS_CODE).text = Trim(StrConv(P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode)) '付加ｸﾗｽ
    Text1(ptxN_CLASS_CODE).text = Trim(StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode)) '内職ｸﾗｽ
    

    If Combo1(pcmbS_TANTO).ListCount = 0 Then                                       '収単／担当者
    Else
        For i = 0 To Combo1(pcmbS_TANTO).ListCount - 1
            If StrConv(P_SSHIJI_O_REC.S_TANTO, vbUnicode) = Right(Combo1(pcmbS_TANTO).List(i), 2) Then
                Combo1(pcmbS_TANTO).ListIndex = i
                Exit For
            End If
        Next i
    End If
    
    
    If StrConv(P_SSHIJI_O_REC.SAMPLE_F, vbUnicode) = P_SAMPLE_F_OFF Then            '見本作成
        Check1(pchkSAMPLE_F).Value = vbUnchecked
    Else
        Check1(pchkSAMPLE_F).Value = vbChecked
    End If
    
    Select Case StrConv(P_SSHIJI_O_REC.SHIJI_F, vbUnicode)                          '通常/ｽﾎﾟｯﾄ/欠品解除
        Case P_SHIJI_F_NORMAL
            Option1(poptSHIJI_NORMAL).Value = True
            Option1(poptSHIJI_SPOT).Value = False
            Option1(poptSHIJI_KEPPIN).Value = False
        Case P_SHIJI_F_SPOT
            Option1(poptSHIJI_NORMAL).Value = False
            Option1(poptSHIJI_SPOT).Value = True
            Option1(poptSHIJI_KEPPIN).Value = False
        Case P_SHIJI_F_KEPPIN
            Option1(poptSHIJI_NORMAL).Value = False
            Option1(poptSHIJI_SPOT).Value = False
            Option1(poptSHIJI_KEPPIN).Value = True
    End Select
    
    
    If StrConv(P_SSHIJI_O_REC.PRI_SHIJI, vbUnicode) = P_PRI_SHIJI_OFF Then          '出力対象　指図票
        Check1(pchkPRI_SHIJI).Value = vbUnchecked
    Else
        Check1(pchkPRI_SHIJI).Value = vbChecked
    End If
    
    If StrConv(P_SSHIJI_O_REC.PRI_PARTS, vbUnicode) = P_PRI_PARTS_OFF Then          '出力対象　ﾊﾟｰﾂﾗﾍﾞﾙ
        Check1(pchkPRI_PARTS).Value = vbUnchecked
    Else
        Check1(pchkPRI_PARTS).Value = vbChecked
    End If
    
    RichTextBox1(prchBIKOU).text = StrConv(P_SSHIJI_O_REC.BIKOU, vbUnicode)         '備考
    
    '--------------------------------   「子」情報
    Erase K_Item_Tbl
    Erase G_Item_Tbl
    Erase D_Item_Tbl
    
    ReDim K_Item_Tbl(0 To 4)
    ReDim G_Item_Tbl(0 To 2)
    ReDim D_Item_Tbl(0 To 49)
    
    
    
    
    k = -1
    g = -1
    d = -1
    
    K_Index = ptxK_HIN_GAI01
    G_Index = ptxG_HIN_GAI01
    DT_Index = ptxD_HIN_GAI01
    DC_Index = pcmbD_SYUBETSU01
    
            
        
    
    Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, Text1(ptxSHIJI_NO).text)
    Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, "")
    Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")
    
    com = BtOpGetGreaterEqual
    
    
    Do
        sts = BTRV(com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_SSHIJI_K_REC.SHIJI_No, vbUnicode) <> Text1(ptxSHIJI_NO).text Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "商品化指図票ﾃﾞｰﾀ(親)")
                Exit Function
        
        End Select
        
        Select Case StrConv(P_SSHIJI_K_REC.DATA_KBN, vbUnicode)
        
            Case P_KOSOU    '個装資材
            
                k = k + 1
                K_Item_Tbl(k).JGYOBU = StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode)
                K_Item_Tbl(k).NAIGAI = StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode)
                            '品番
                Text1(K_Index).text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
'                If K_Item_Tbl(k).JGYOBU = SHIZAI Then
'                    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
'                Else
'                    Call UniCode_Conv(K0_ITEM.JGYOBU, K_Item_Tbl(k).JGYOBU)
'                End If


                Select Case K_Item_Tbl(k).JGYOBU
                    Case SHIZAI
                        Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
                
                    Case SETSUBI
                        Call UniCode_Conv(K0_ITEM.JGYOBU, YUKO_JGYOBU)
                    Case Else
                        Call UniCode_Conv(K0_ITEM.JGYOBU, K_Item_Tbl(k).JGYOBU)
                End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                Call UniCode_Conv(K0_ITEM.NAIGAI, K_Item_Tbl(k).NAIGAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(K_Index).text)
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        '品名
                        Text1(K_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                        '標準棚番
                        If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                            Text1(K_Index + 4) = ""
                        Else
                            Text1(K_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_DAN, vbUnicode)
                        End If
                    
                    Case BtErrKeyNotFound
'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21
                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(K_Index).text)
    
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            
                                '品名
                                Text1(K_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                '標準棚番
                                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                                    Text1(K_Index + 4) = ""
                                Else
                                    Text1(K_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_DAN, vbUnicode)
                                End If
                            
                            Case BtErrKeyNotFound
    
                                Text1(K_Index + 1) = "未登録品番"
                                Text1(K_Index + 4) = ""
                            Case Else
                                Call Input_UnLock             '2008.01.15
                                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                Exit Function
    
                        End Select
'                        Text1(K_Index + 1) = "未登録品番"
'                        Text1(K_Index + 4) = ""
'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "")
                        Exit Function
                
                End Select
            
            
                Text1(K_Index + 2).text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                Text1(K_Index + 3).text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
                            
                K_Index = K_Index + 5
            
            
            
            Case P_GAISOU   '外装資材
                g = g + 1
                G_Item_Tbl(g).JGYOBU = StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode)
                G_Item_Tbl(g).NAIGAI = StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode)
                            '品番
                Text1(G_Index).text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
'                If G_Item_Tbl(g).JGYOBU = SHIZAI Then
'                    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
'                Else
'                    Call UniCode_Conv(K0_ITEM.JGYOBU, G_Item_Tbl(g).JGYOBU)
'                End If

                Select Case G_Item_Tbl(g).JGYOBU
                    Case SHIZAI
                        Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
                    Case SETSUBI
                        Call UniCode_Conv(K0_ITEM.JGYOBU, YUKO_JGYOBU)
                    Case Else
                        Call UniCode_Conv(K0_ITEM.JGYOBU, G_Item_Tbl(g).JGYOBU)
                End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                Call UniCode_Conv(K0_ITEM.NAIGAI, G_Item_Tbl(g).NAIGAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(G_Index).text)
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        '品名
                        Text1(G_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                        '標準棚番
                        If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                            Text1(G_Index + 4) = ""
                        Else
                            Text1(G_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_DAN, vbUnicode)
                        End If
                    
                    Case BtErrKeyNotFound
'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21
                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(G_Index).text)
    
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            
                                '品名
                                Text1(G_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                '標準棚番
                                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                                    Text1(G_Index + 4) = ""
                                Else
                                    Text1(G_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_DAN, vbUnicode)
                                End If
                            
                            Case BtErrKeyNotFound
    
                                Text1(G_Index + 1) = "未登録品番"
                                Text1(G_Index + 4) = ""
                            Case Else
                                Call Input_UnLock             '2008.01.15
                                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                Exit Function
    
                        End Select
'                        Text1(G_Index + 1) = "未登録品番"
'                        Text1(G_Index + 4) = ""
'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "")
                        Exit Function
                
                End Select
            
            
                Text1(G_Index + 2).text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                Text1(G_Index + 3).text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
                            
                G_Index = G_Index + 5
            
            
            Case P_DOUKON   '同梱／構成
            
                d = d + 1
' 2013.01.07 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                D_Item_Tbl(d).JGYOBU = StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode)
'                D_Item_Tbl(d).JGYOBU = BUZAI            '同梱/構成の事業部を「部材」固定に変更
' 2013.01.07 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                D_Item_Tbl(d).NAIGAI = StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode)
                            
                D_Item_Tbl(d).SYUBETSU = StrConv(P_SSHIJI_K_REC.KO_SYUBETSU, vbUnicode)
                D_Item_Tbl(d).HIN_GAI = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                D_Item_Tbl(d).QTY = CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode))
                D_Item_Tbl(d).SHIJI_QTY = CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode))
                D_Item_Tbl(d).BIKOU = StrConv(P_SSHIJI_K_REC.KO_BIKOU, vbUnicode)
                            
                
                            
                If d < 6 Then
                            
                            
                                '種別
                    Combo1(DC_Index).ListIndex = -1
                    For i = 0 To Combo1(DC_Index).ListCount - 1
                    
                        If StrConv(P_SSHIJI_K_REC.KO_SYUBETSU, vbUnicode) = Right(Combo1(DC_Index).List(i), 2) Then
                            Combo1(DC_Index).ListIndex = i
                            Exit For
                        End If
                    
                    Next i
                                
                                
                                '事業部 2016.01.27
                    Combo2(DC_Index).ListIndex = -1
                    For i = 0 To Combo2(DC_Index).ListCount - 1
                    
                        If StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) = Right(Combo2(DC_Index).List(i), 1) Then
                            Combo2(DC_Index).ListIndex = i
                            Exit For
                        End If
                    
                    Next i
                                
                                
                                
                    DC_Index = DC_Index + 1
                                
                                '品番
                    Text1(DT_Index).text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                    
' 2013.01.07 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'                    If D_Item_Tbl(d).JGYOBU = SHIZAI Then
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
'                    Else
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(d).JGYOBU)
'                    End If

'2016.01.27                    Select Case D_Item_Tbl(d).JGYOBU
'2016.01.27                        Case SHIZAI
'2016.01.27                            Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
'2016.01.27                        Case SETSUBI
'2016.01.27                            Call UniCode_Conv(K0_ITEM.JGYOBU, YUKO_JGYOBU)
'2016.01.27                        Case Else
'2016.01.27                            Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(d).JGYOBU)
'2016.01.27                    End Select
' 2013.01.07 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                    Call UniCode_Conv(K0_ITEM.NAIGAI, D_Item_Tbl(d).NAIGAI)
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(DT_Index).text)
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                            '品名
                            Text1(DT_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                            '標準棚番
                            If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                                Text1(DT_Index + 4) = ""
                            Else
                                Text1(DT_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                            End If
                        
                            '>>>>>>>>>>>>>>>>>>>    2013.01.07
                            'If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI Then
                            '    wkJgyobu = BUZAI
                            'Else
                            '    'wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)  2012.04.04
                            '    wkJgyobu = YUKO_JGYOBU                          '2012.04.04
                            'End If
                            
                            Select Case StrConv(ITEMREC.JGYOBU, vbUnicode)
                                Case SHIZAI
                                    wkJgyobu = BUZAI
                                Case BUZAI
                                    wkJgyobu = YUKO_JGYOBU
                                Case Else
                                    wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
                            End Select
                            '>>>>>>>>>>>>>>>>>>>    2013.01.07
                        
                            '2013.01.11 標準棚番を設定
                            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, wkJgyobu, _
                                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                                                    (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)), _
                                                                    , , Jyogai_Soko_umu) Then
                                Exit Function
                            
                            End If
                        
                            Text1(DT_Index + 5) = Format(Sumi_Qty + Mi_Qty, "#0")
                        
                        
                        Case BtErrKeyNotFound
                            
'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21
                            Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(DT_Index).text)
        
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            Select Case sts
                                Case BtNoErr
                                
                                    '品名
                                    Text1(DT_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                    '標準棚番
                                    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                                        Text1(DT_Index + 4) = ""
                                    Else
                                        Text1(DT_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                                    End If
                                
                                
                                
                                    Wk_LOC = StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                             StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)
        
                                    If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                                            StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                                                            Wk_LOC) Then
                                        Exit Function
        
                                    End If
                                
                                    Text1(DT_Index + 5) = Format(Sumi_Qty + Mi_Qty, "#0")
                                
                                Case BtErrKeyNotFound
                                    Text1(DT_Index + 1) = "未登録品番"
                                    Text1(DT_Index + 4) = ""
                                    Text1(DT_Index + 5) = ""
                                Case Else
                                    Call Input_UnLock             '2008.01.15
                                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                    Exit Function
        
                            End Select
'                            Text1(DT_Index + 1) = "未登録品番"
'                            Text1(DT_Index + 4) = ""
'                            Text1(DT_Index + 5) = ""
'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21
                                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "")
                            Exit Function
                    
                    End Select
                
                
                    Text1(DT_Index + 2).text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                    Text1(DT_Index + 3).text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
                    Text1(DT_Index + 6).text = Trim(StrConv(P_SSHIJI_K_REC.KO_BIKOU, vbUnicode))
                                
                    DT_Index = DT_Index + 7
                End If
        
        End Select
        
        
        
        
        com = BtOpGetNext
    
    Loop
    
    Item_Disp_Proc = False

End Function

Private Function Update_Proc(Mode As Integer, Optional HAKKO_F As Integer = 0, Optional MSG As Integer = 0) As Integer
'----------------------------------------------------------------------------
'                   構成マスタ＆商品化指示ﾃﾞｰﾀ出力
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim ans         As Integer
Dim com         As Integer

Dim SEQNO       As Integer

Dim i           As Integer
Dim j           As Integer

Dim k           As Integer              '2012.03.09

Dim SHIJINO     As Long


Dim ORDER_DT    As String * 10          '2012.03.27

    Update_Proc = True
                                        
    Call Input_Lock
                                        
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If

    
    If Text1(ptxSHIJI_NO).text = "" Then
                                        
                                            '管理ファイルより指図票番号の獲得
        Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
        
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound
                
                    If P_KANRI_MAKE_Proc() Then
                        GoTo Abort_Tran
                    End If
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Update_Proc = True
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "管理マスタ")
                    GoTo Abort_Tran
            
            End Select
        
        
        Loop
    
        '指図票№＋１
    
        If CLng(StrConv(P_KANRIREC.SASHIZU_NO, vbUnicode)) = 99999999 Then
            Call UniCode_Conv(P_KANRIREC.SASHIZU_NO, "00000001")
        Else
            Call UniCode_Conv(P_KANRIREC.SASHIZU_NO, Format(CLng(StrConv(P_KANRIREC.SASHIZU_NO, vbUnicode)) + 1, "00000000"))
        End If
    
    
        Do
            
            DoEvents
            
            sts = BTRV(BtOpUpdate, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "管理マスタ")
                        End If
                        GoTo Abort_Tran
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "管理マスタ")
                    GoTo Abort_Tran
            End Select
        Loop

        SHIJINO = CLng(StrConv(P_KANRIREC.SASHIZU_NO, vbUnicode))
    Else
        
        SHIJINO = CLng(Text1(ptxSHIJI_NO).text)
    
    End If
    '---------------------------------------------------    '収単／担当者有りは品目マスタ更新
        
'    If PRI_S_TANTO Then
        
        Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)
                
        
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound
                    MsgBox "品目マスタが他端末で変更されています。更新処理を中止します。"
                    GoTo Abort_Tran
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Update_Proc = True
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                    GoTo Abort_Tran
            
            End Select
        
        
        Loop
                                                                                '収単／担当者ｸﾗｽ
        Call UniCode_Conv(ITEMREC.S_TANTO, Right(Combo1(pcmbS_TANTO).text, 2))
    
        If IsNumeric(StrConv(ITEMREC.L_LABEL, vbUnicode)) Then
            G_Kisyu_F = CInt(StrConv(ITEMREC.L_LABEL, vbUnicode))
            
            If IsNumeric(StrConv(ITEMREC.L_URIKIN1, vbUnicode)) Then
                
                L_URIKIN1 = CDbl(StrConv(ITEMREC.L_URIKIN1, vbUnicode))
            Else
                L_URIKIN1 = 0
            End If
        
            If IsNumeric(StrConv(ITEMREC.L_URIKIN2, vbUnicode)) Then
                
                L_URIKIN2 = CDbl(StrConv(ITEMREC.L_URIKIN2, vbUnicode))
            Else
                L_URIKIN2 = 0
            End If
        
            If IsNumeric(StrConv(ITEMREC.L_URIKIN3, vbUnicode)) Then
                
                L_URIKIN3 = CDbl(StrConv(ITEMREC.L_URIKIN3, vbUnicode))
            Else
                L_URIKIN3 = 0
            End If
        
        
        Else
            G_Kisyu_F = 0
            L_URIKIN1 = 0
            L_URIKIN2 = 0
            L_URIKIN3 = 0
        
        End If
        Do
            
            DoEvents
            
            sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "品目マスタ")
                        End If
                        GoTo Abort_Tran
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "品目マスタ")
                    GoTo Abort_Tran
            End Select
        Loop
    
'    End If
    '---------------------------------------------------    '構成マスタ更新
        
        
        
If PI00015_LOG <> "" Then
    Call LOG_OUT(PI00015_LOG, "ITEM_CODE=" & Text1(ptxHIN_GAI).text & " DEL START" & " Mode =" & Mode)
End If
        
    '該当データ全件削除
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).text)
       
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "")
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
       
    com = BtOpGetGreater
       
    Do
        
        DoEvents
        
        Do
        
            sts = BTRV(com + BtSNoWait, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                
            Select Case sts
                Case BtNoErr
                
                    If StrConv(P_COMPO_O_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                        StrConv(P_COMPO_O_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                        StrConv(P_COMPO_O_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                        Trim(StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHIN_GAI).text) Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "構成マスタ")
                            GoTo Abort_Tran
                        End If
                        sts = BtErrEOF
                    End If
                    Exit Do
                Case BtErrEOF
                    Exit Do
                
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_CMOPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        GoTo Abort_Tran
                    End If
                
                
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "構成マスタ")
                    GoTo Abort_Tran
            End Select
    
        Loop
            
        If sts = BtErrEOF Then
If PI00015_LOG <> "" Then
    Call LOG_OUT(PI00015_LOG, "ITEM_CODE=" & Text1(ptxHIN_GAI).text & " DEL NORMAL END" & " Mode =" & Mode)
End If
            Exit Do
        End If


        Do
            sts = BTRV(BtOpDelete, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_CMOPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "構成マスタ")
                        End If
                        GoTo Abort_Tran
                    End If
                
                
                Case Else
                    Call File_Error(sts, BtOpDelete, "構成マスタ")
                    GoTo Abort_Tran
            End Select
        Loop
    
        com = BtOpGetNext
    
    Loop
        
    '構成マスタ(ﾍｯﾀﾞｰ)出力
If PI00015_LOG <> "" Then
    Call LOG_OUT(PI00015_LOG, "ITEM_CODE=" & Text1(ptxHIN_GAI).text & " HEAD INSERT" & " Mode =" & Mode)
End If
                                                                                '仕向け先ｺｰﾄﾞ
    Call UniCode_Conv(P_COMPO_O_REC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 1, 2))
                                                                                '事業部
    Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                                                                                '国内外
    Call UniCode_Conv(P_COMPO_O_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
    Call UniCode_Conv(P_COMPO_O_REC.HIN_GAI, Text1(ptxHIN_GAI).text)
    Call UniCode_Conv(P_COMPO_O_REC.DATA_KBN, P_HEAD)
    Call UniCode_Conv(P_COMPO_O_REC.SEQNO, "000")

    Call UniCode_Conv(P_COMPO_O_REC.CLASS_CODE, Text1(ptxS_CLASS_CODE).text)    'ｸﾗｽｺｰﾄﾞ
    Call UniCode_Conv(P_COMPO_O_REC.BIKOU, RichTextBox1(prchBIKOU).text)        '備考
    
    Call UniCode_Conv(P_COMPO_O_REC.F_CLASS_CODE, Text1(ptxF_CLASS_CODE).text)  '付加ｺｰﾄﾞ
    
    Call UniCode_Conv(P_COMPO_O_REC.N_CLASS_CODE, Text1(ptxN_CLASS_CODE).text)  '内職ｺｰﾄﾞ


    Call UniCode_Conv(P_COMPO_O_REC.FILLER, "")

    Call UniCode_Conv(P_COMPO_O_REC.UPD_TANTO, Text1(ptxTANTO_CODE))            '更新担当者ｺｰﾄﾞ
                                                                                '更新日時
    Call UniCode_Conv(P_COMPO_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))


    Do
        
        DoEvents
        
        sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "構成マスタ")
                GoTo Abort_Tran
        End Select
    
    Loop
If PI00015_LOG <> "" Then
    Call LOG_OUT(PI00015_LOG, "ITEM_CODE=" & Text1(ptxHIN_GAI).text & " HEAD INSERT" & " Mode =" & Mode)
End If

    '構成マスタ(ﾎﾞﾃﾞｨ)出力


    '個装資材分
    SEQNO = 0
    j = 0
    For i = ptxK_HIN_GAI01 To ptxK_HIN_GAI05 Step 5
    
        If Trim(Text1(i).text) = "" Then
        Else
                                                                                        
            SEQNO = SEQNO + 10
                                                                                        
                                                                                        '仕向け先ｺｰﾄﾞ
            Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 1, 2))
                                                                                        '事業部
            Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                                                                                        '国内外
            Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
            Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, Text1(ptxHIN_GAI).text)
            Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_KOSOU)                          'データ区分
            Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(SEQNO, "000"))                '追番
                        
            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, "")                            '種別
            Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, K_Item_Tbl(j).JGYOBU)            '事業部
            Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, K_Item_Tbl(j).NAIGAI)            '国内外
            Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, Text1(i).text)                  '品番
                                                                                        '員数
            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(CDbl(Text1(i + 2).text), "000.00"))
            Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, "")                               '備考
        
            Call UniCode_Conv(P_COMPO_K_REC.FILLER, "")
        
            Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, Text1(ptxTANTO_CODE).text)       '更新担当者ｺｰﾄﾞ
                                                                                        '更新日時
            Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
        
            Do
                
                DoEvents
                
                sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpInsert, "構成マスタ")
                        GoTo Abort_Tran
                End Select
            
            Loop
        
        
        End If
        
        j = j + 1
    
    
    Next i

    '外装資材分
    SEQNO = 0
    j = 0
    For i = ptxG_HIN_GAI01 To ptxG_HIN_GAI03 Step 5
    
        If Trim(Text1(i).text) = "" Then
        Else
            
            SEQNO = SEQNO + 10
                                                                                        
                                                                                        '仕向け先ｺｰﾄﾞ
            Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 1, 2))
                                                                                        '事業部
            Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                                                                                        '国内外
            Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
            Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, Text1(ptxHIN_GAI).text)
            Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_GAISOU)                         'データ区分
            Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(SEQNO, "000"))                '追番
                        
            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, "")                            '種別
            Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, G_Item_Tbl(j).JGYOBU)            '事業部
            Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, G_Item_Tbl(j).NAIGAI)            '国内外
            Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, Text1(i).text)                  '品番
                                                                                        '員数
            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(CDbl(Text1(i + 2).text), "000.00"))
            Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, "")                               '備考
        
            Call UniCode_Conv(P_COMPO_K_REC.FILLER, "")
        
            Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, Text1(ptxTANTO_CODE).text)       '更新担当者ｺｰﾄﾞ
                                                                                        '更新日時
            Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
            Do
                
                DoEvents
                
                sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpInsert, "構成マスタ")
                        GoTo Abort_Tran
                End Select
            
            Loop
        
        
        
        End If
        
        j = j + 1
    
    
    Next i


    '同梱／構成分
If PI00015_LOG <> "" Then
    Call LOG_OUT(PI00015_LOG, "ITEM_CODE=" & Text1(ptxHIN_GAI).text & " BODY INSERT START" & " Mode =" & Mode)
End If
    SEQNO = 0
    For i = 0 To 49
    
        If D_Item_Tbl(i).JGYOBU = vbNullChar Or _
            Trim(D_Item_Tbl(i).JGYOBU) = "" Then
        Else
            SEQNO = SEQNO + 10
                                                                                        
                                                                                        '仕向け先ｺｰﾄﾞ
            Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 1, 2))
                                                                                        '事業部
            Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                                                                                        '国内外
            Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
            Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, Text1(ptxHIN_GAI).text)
            Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_DOUKON)                         'データ区分
            Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(SEQNO, "000"))                '追番
                        
            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, D_Item_Tbl(i).SYUBETSU)        '種別
            Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, D_Item_Tbl(i).JGYOBU)            '事業部
            Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, D_Item_Tbl(i).NAIGAI)            '国内外
            Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, D_Item_Tbl(i).HIN_GAI)          '品番
                                                                                        '員数
            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(D_Item_Tbl(i).QTY, "000.00"))
            Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, D_Item_Tbl(i).BIKOU)              '備考
        
            Call UniCode_Conv(P_COMPO_K_REC.FILLER, "")
        
            Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, Text1(ptxTANTO_CODE).text)       '更新担当者ｺｰﾄﾞ
                                                                                        '更新日時
            Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
        
            Do
                
                DoEvents
                
                sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpInsert, "構成マスタ")
                        GoTo Abort_Tran
                End Select
            
            Loop
        
If PI00015_LOG <> "" Then
    Call LOG_OUT(PI00015_LOG, "ITEM_CODE=" & Text1(ptxHIN_GAI).text & " BODY INSERT" & " Mode =" & Mode & "KO_ITEM_CODE = " & D_Item_Tbl(i).HIN_GAI)
End If
        
        
        
        End If
    
    Next i
If PI00015_LOG <> "" Then
    Call LOG_OUT(PI00015_LOG, "ITEM_CODE=" & Text1(ptxHIN_GAI).text & " BODY INSERT NORMAL END" & " Mode =" & Mode)
End If


    If Mode = 1 Then
        GoTo End_Tran
    End If
    
    '---------------------------------------------------    '指図票データ更新
    
    '指図票データ(ﾍｯﾀﾞｰ)処理
    
    
    Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, Format(SHIJINO, "00000000"))
    
    Do
    
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            
        Select Case sts
            Case BtNoErr
            
                com = BtOpUpdate
                
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_SSHIJI_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            
            
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "商品化指図票ﾃﾞｰﾀ(親)")
                GoTo Abort_Tran
        End Select

    Loop
    
    
    If com = BtOpInsert Then
        '新規作成
        Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_No, Format(SHIJINO, "00000000"))    '指図票№
        
        Call UniCode_Conv(P_SSHIJI_O_REC.HAKKO_DT, "")
        
        
        Call UniCode_Conv(P_SSHIJI_O_REC.Print_datetime, "")                    '発行日時
        
        Call UniCode_Conv(P_SSHIJI_O_REC.KAN_F, P_KAN_OFF)                      '完了F
        Call UniCode_Conv(P_SSHIJI_O_REC.KAN_DT, "")                            '完了日
        Call UniCode_Conv(P_SSHIJI_O_REC.BUNNOU_CNT, "00")                      '分納回数
        Call UniCode_Conv(P_SSHIJI_O_REC.UKEIRE_QTY, "00000000")                '受入回数
    
        For i = 0 To 9                                                          '原価管理
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(i).NIN, "0.0")           '人数
            Call UniCode_Conv(P_SSHIJI_O_REC.GENKA_TBL(i).TIMES, "000.00")      '時間
        Next i
    
            
        Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_NAME, "")                       '自責要因名
        Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_NIN, "0.0")                     '        人
        Call UniCode_Conv(P_SSHIJI_O_REC.JISEKI_TIMES, "000.00")                '        分
            
        Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_NAME, "")                       '他責要因名
        Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_NIN, "0.0")                     '        人
        Call UniCode_Conv(P_SSHIJI_O_REC.TASEKI_TIMES, "000.00")                '        分
            
            
    
    
        Call UniCode_Conv(P_SSHIJI_O_REC.CANCEL_F, P_CANCEL_OFF)                'ｷｬﾝｾﾙﾌﾗｸﾞ
        Call UniCode_Conv(P_SSHIJI_O_REC.CANCEL_DATETIME, "")                   'ｷｬﾝｾﾙ日時
        
        Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_TANTO, "")                   '品番ﾁｪｯｸ担当者ｺｰﾄﾞ 2013.08.21
        Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_DATETIME, "")                '品番ﾁｪｯｸ日時       2013.08.21
        
        
        
        Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_LABEL_CNT, "000")            '品番ﾁｪｯｸﾗﾍﾞﾙ件数   2010.09.03
        Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_GENPIN_CNT, "000")           '品番ﾁｪｯｸﾗﾍﾞﾙ件数   2010.09.03
        
        Call UniCode_Conv(P_SSHIJI_O_REC.ORDER_DT_SEQ, "")                      '受注日(注文№)枝番 2012.03.27
        Call UniCode_Conv(P_SSHIJI_O_REC.COMPO_END_F, "")                       '構成ﾁｪｯｸ完了F(大阪PC) 9:完了 2012.04.13
            
            
        
'        Call UniCode_Conv(P_SSHIJI_O_REC.FILLER, "")                           '2016.01.13
        Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_GAISOU_CNT, "")              '2016.01.13

    End If
                                                                                '発行日
    If HAKKO_F = 1 Then
        Call UniCode_Conv(P_SSHIJI_O_REC.HAKKO_DT, Format(Now, "YYYYMMDD"))
        Call UniCode_Conv(P_SSHIJI_O_REC.Print_datetime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    End If
                                                                                '担当者ｺｰﾄﾞ
    Call UniCode_Conv(P_SSHIJI_O_REC.TANTO_CODE, Text1(ptxTANTO_CODE).text)
                                                                                '承認者ｺｰﾄﾞ
    Call UniCode_Conv(P_SSHIJI_O_REC.SHONIN_CODE, Text1(ptxSHONIN_CODE).text)
                                                                                '仕向け先ｺｰﾄﾞ
    Call UniCode_Conv(P_SSHIJI_O_REC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 1, 2))
                                                                                '事業部
    Call UniCode_Conv(P_SSHIJI_O_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                                                                                '国内外
    Call UniCode_Conv(P_SSHIJI_O_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
                                                                                '品番
    Call UniCode_Conv(P_SSHIJI_O_REC.HIN_GAI, Text1(ptxHIN_GAI).text)
                                                                                '数量
    Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_QTY, Format(CDbl(Text1(ptxSHIJI_QTY).text), "00000000.00"))
                                                                                '手配先
    Call UniCode_Conv(P_SSHIJI_O_REC.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE).text)
                                                                                '取引先区分
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE).text)
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrEOF
            MsgBox "手配先情報が他で変更されました。更新処理を中止します。"
            GoTo Abort_Tran
        Case Else
            Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
            Exit Function
    
    End Select
    Call UniCode_Conv(P_SSHIJI_O_REC.TORI_KBN, StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode))
                                                                                '商品化ｸﾗｽ
    Call UniCode_Conv(P_SSHIJI_O_REC.S_CLASS_CODE, Text1(ptxS_CLASS_CODE).text)
                                                                                '付加ｸﾗｽ
    Call UniCode_Conv(P_SSHIJI_O_REC.F_CLASS_CODE, Text1(ptxF_CLASS_CODE).text)
                                                                                '内職ｸﾗｽ
    Call UniCode_Conv(P_SSHIJI_O_REC.N_CLASS_CODE, Text1(ptxN_CLASS_CODE).text)
                                                                                '収単／担当者ｸﾗｽ
    Call UniCode_Conv(P_SSHIJI_O_REC.S_TANTO, Right(Combo1(pcmbS_TANTO).text, 2))
                                                                                                                                                        
    If Check1(pchkSAMPLE_F).Value = vbChecked Then                              '見本作成
        Call UniCode_Conv(P_SSHIJI_O_REC.SAMPLE_F, P_SAMPLE_F_ON)
    Else
        Call UniCode_Conv(P_SSHIJI_O_REC.SAMPLE_F, P_SAMPLE_F_OFF)
    End If

    If Option1(poptSHIJI_NORMAL).Value Then
        Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_F, P_SHIJI_F_NORMAL)             '通常
    Else
        If Option1(poptSHIJI_SPOT).Value Then
            Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_F, P_SHIJI_F_SPOT)           'ｽﾎﾟｯﾄ
        Else
            If Option1(poptSHIJI_KEPPIN).Value Then
                Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_F, P_SHIJI_F_KEPPIN)     '欠品解除
            
            Else
                If Option1(poptSHIJI_SAIKON).Value Then
                    Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_F, P_SHIJI_F_SAIKON) '再梱包 2007.11.09
            
                End If
            End If
        End If
    End If
    

    If Check1(pchkPRI_SHIJI).Value = vbChecked Then                             '出力対象　指図票
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_SHIJI, P_PRI_SHIJI_ON)
'''        Call UniCode_Conv(P_SSHIJI_O_REC.Print_datetime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    Else
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_SHIJI, P_PRI_SHIJI_OFF)
    End If

    If Check1(pchkPRI_PARTS).Value = vbChecked Then                             '出力対象　ﾊﾟｰﾂﾗﾍﾞﾙ
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_PARTS, P_PRI_PARTS_ON)
    Else
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_PARTS, P_PRI_PARTS_OFF)
    End If

    If Check1(pchkPRI_GAISOU).Value = vbChecked Then                            '出力対象　外装ﾗﾍﾞﾙ
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_GAISOU, P_PRI_GAISOU_ON)
    Else
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_GAISOU, P_PRI_GAISOU_OFF)
    End If

    If Check1(pchkPRI_KISHU).Value = vbChecked Then                             '出力対象　外装ﾗﾍﾞﾙ
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_KISHU, P_PRI_KISHU_ON)
    Else
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_KISHU, P_PRI_KISHU_OFF)
    End If

    Call UniCode_Conv(P_SSHIJI_O_REC.BIKOU, RichTextBox1(prchBIKOU).text)       '備考
    
                                                                                '更新日時
    Call UniCode_Conv(P_SSHIJI_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                                                                                
                                                                                
'-----------------------------------------------------------------------------  2012.03.18　受注日--＞注文№
                                                                                '受注日
    If Trim(Text1(ptxORDER_NO).text) = "" Then
        Call UniCode_Conv(P_SSHIJI_O_REC.ORDER_DT, "")
        Call UniCode_Conv(P_SSHIJI_O_REC.ORDER_DT_SEQ, "")
    Else
'2012.03.17        Call UniCode_Conv(P_SSHIJI_O_REC.ORDER_DT, Format(Text1(ptxORDER_DT).text, "YYYYMMDD"))
                
        ORDER_DT = Text1(ptxORDER_NO).text
        Call UniCode_Conv(P_SSHIJI_O_REC.ORDER_DT, Mid(ORDER_DT, 1, 8))
        Call UniCode_Conv(P_SSHIJI_O_REC.ORDER_DT_SEQ, Mid(ORDER_DT, 9, 2))
    End If
'-----------------------------------------------------------------------------  2012.03.18　受注日--＞注文№
    
    Do
        
        DoEvents
        
        sts = BTRV(com, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    If com = BtOpUpdate Then
                        sts = BTRV(BtOpUnlock, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "商品化指図ﾃﾞｰﾀ(親)")
                        End If
                    End If
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, com, "商品化指図ﾃﾞｰﾀ(親)")
                GoTo Abort_Tran
        End Select
    
    Loop
    
    If com = BtOpUpdate Then
        
        
'--------------------------------------------------- 大阪  部材対応　2012.03.09
        
                
        
        For k = 0 To UBound(ZAIKO_FUSOKU)
        
            ZAIKO_FUSOKU(k).IDO_SUMI = ""
            ZAIKO_FUSOKU(k).HIKIATE_QTY = 0
            ZAIKO_FUSOKU(k).IDO_SUMI_QTY = 0        '2012.04.13
        
        Next k
'--------------------------------------------------- 大阪  部材対応　2012.03.09
        
        
        '対象の子を削除する
        Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, Format(SHIJINO, "00000000"))
        Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, "")
        Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")
    
        com = BtOpGetGreater
    
        Do
            
            DoEvents
            
            Do
            
                sts = BTRV(com + BtSNoWait, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                    
                Select Case sts
                    Case BtNoErr
                    
                        If StrConv(P_SSHIJI_K_REC.SHIJI_No, vbUnicode) <> Format(SHIJINO, "00000000") Then
                            sts = BTRV(BtOpUnlock, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                            If sts Then
                                Call File_Error(sts, BtOpUnlock, "商品化指図ﾃﾞｰﾀ(子)")
                                GoTo Abort_Tran
                            End If
                            sts = BtErrEOF
                        End If
                        Exit Do
                    Case BtErrEOF
                        Exit Do
                    
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<P_SSHIJI_K.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            GoTo Abort_Tran
                        End If
                    
                    
                    Case Else
                        Call File_Error(sts, com + BtSNoWait, "商品化指図ﾃﾞｰﾀ(子)")
                        GoTo Abort_Tran
                End Select
        
            Loop
                
            If sts = BtErrEOF Then
                Exit Do
            End If
    
    
'--------------------------------------------------- 大阪  部材対応　2012.03.09
            
            
            Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI, "")                  '2013.02.13
            Call UniCode_Conv(P_SSHIJI_K_REC.HIKIATE_QTY, "")               '2013.02.13
            Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI_QTY, "")              '2013.02.13
            
            
            
            If Trim(Text1(ptxORDER_NO).text) <> "" Then
                For k = 0 To UBound(ZAIKO_FUSOKU)
                
                
                    If ZAIKO_FUSOKU(k).JGYOBU = StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) And _
                        ZAIKO_FUSOKU(k).NAIGAI = StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode) And _
                        ZAIKO_FUSOKU(k).HIN_GAI = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode) Then
                
                        ZAIKO_FUSOKU(k).IDO_SUMI = StrConv(P_SSHIJI_K_REC.IDO_SUMI, vbUnicode)
                        ZAIKO_FUSOKU(k).HIKIATE_QTY = Val(StrConv(P_SSHIJI_K_REC.HIKIATE_QTY, vbUnicode))
                        ZAIKO_FUSOKU(k).IDO_SUMI_QTY = Val(StrConv(P_SSHIJI_K_REC.IDO_SUMI_QTY, vbUnicode))     '2012.04.13
                        
                
                        Exit For
                
                    End If
                Next k
            End If
'--------------------------------------------------- 大阪  部材対応　2012.03.09
            Do
                sts = BTRV(BtOpDelete, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                Select Case sts
                    Case BtNoErr
                    
                        Exit Do
                    Case BtErrEOF
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<P_CMOPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            sts = BTRV(BtOpUnlock, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                            If sts Then
                                Call File_Error(sts, BtOpUnlock, "商品化指図ﾃﾞｰﾀ(子)")
                                GoTo Abort_Tran
                            End If
                            GoTo Abort_Tran
                        End If
                    
                    
                    Case Else
                        Call File_Error(sts, BtOpDelete, "商品化指図ﾃﾞｰﾀ(子)")
                        GoTo Abort_Tran
                End Select
            Loop
        
            com = BtOpGetNext
        
        Loop
    End If
    
    
    '商品化指図票ﾃﾞｰﾀ(ﾎﾞﾃﾞｨ)出力


    '個装資材分
    SEQNO = 0
    j = 0
    For i = ptxK_HIN_GAI01 To ptxK_HIN_GAI05 Step 5
    
        If Trim(Text1(i).text) = "" Then
        Else
            Call UniCode_Conv(P_SSHIJI_K_REC.SHIJI_No, Format(SHIJINO, "00000000"))        '指図票№
                                                                                        
                                                                                        
                                                                                        
            SEQNO = SEQNO + 10
            Call UniCode_Conv(P_SSHIJI_K_REC.DATA_KBN, P_KOSOU)                         'データ区分
            
            Call UniCode_Conv(P_SSHIJI_K_REC.SEQNO, Format(SEQNO, "000"))               '追番
                        
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_SYUBETSU, "")                           '種別
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_JGYOBU, K_Item_Tbl(j).JGYOBU)           '事業部
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_NAIGAI, K_Item_Tbl(j).NAIGAI)           '国内外
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_HIN_GAI, Text1(i).text)                 '品番
                                                                                        '員数
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_QTY, Format(CDbl(Text1(i + 2).text), "000.00"))
                                                                                        '数量
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_SHIJI_QTY, Format(CDbl(Text1(i + 3).text), "000000000.00"))
            
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_BIKOU, "")                              '備考
        
            Call UniCode_Conv(P_SSHIJI_K_REC.CALCEL_F, P_CANCEL_OFF)                    'ｷｬﾝｾﾙﾌﾗｸﾞ
            Call UniCode_Conv(P_SSHIJI_K_REC.CANCEL_DATETIME, "")                       'ｷｬﾝｾﾙ日時
        
        
        
 '--------------------------------------------------- 大阪  部材対応　2012.03.18
            If K_Item_Tbl(j).JGYOBU = SHIZAI Then           '2013.03.31
            
            
                Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)    '2013.03.31
            Else                                            '2013.03.31
                Call UniCode_Conv(K0_ITEM.JGYOBU, K_Item_Tbl(j).JGYOBU)
            
            End If                                          '2013.03.31
            
            
            Call UniCode_Conv(K0_ITEM.NAIGAI, K_Item_Tbl(j).NAIGAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(i).text)
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Call UniCode_Conv(P_SSHIJI_K_REC.ST_TANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_DAN, vbUnicode))
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_SSHIJI_K_REC.ST_TANABAN, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    GoTo Abort_Tran
            End Select
'--------------------------------------------------- 大阪  部材対応　2012.03.18
       
        
        
        
            Call UniCode_Conv(P_SSHIJI_K_REC.FILLER, "")
                                                                                        '更新日時
            Call UniCode_Conv(P_SSHIJI_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
                                                                                        '出荷予定ＩＤ
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_ID_NO, "")
        
        
                    
'--------------------------------------------------- 大阪  部材対応　2012.03.09
            Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI, "")                  '2013.02.13
            Call UniCode_Conv(P_SSHIJI_K_REC.HIKIATE_QTY, "")               '2013.02.13
            Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI_QTY, "")              '2013.02.13
            
            For k = 0 To UBound(ZAIKO_FUSOKU)
            
            
                If ZAIKO_FUSOKU(k).JGYOBU = StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) And _
                    ZAIKO_FUSOKU(k).NAIGAI = StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode) And _
                    ZAIKO_FUSOKU(k).HIN_GAI = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode) Then
            
                    Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI, ZAIKO_FUSOKU(k).IDO_SUMI)
                    Call UniCode_Conv(P_SSHIJI_K_REC.HIKIATE_QTY, Format(ZAIKO_FUSOKU(k).HIKIATE_QTY, "00000000.00"))
                    
                    Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI_QTY, Format(ZAIKO_FUSOKU(k).IDO_SUMI_QTY, "00000000.00")) '2012.04.13
            
                    Exit For
            
                End If
            Next k
'--------------------------------------------------- 大阪  部材対応　2012.03.09
        
        
        
        
            Do
                
                DoEvents
                
                sts = BTRV(BtOpInsert, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<P_SSHIJI_K.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpInsert, "商品化指図ﾃﾞｰﾀ(子)")
                        GoTo Abort_Tran
                End Select
            
            Loop
        
        
        End If
        
        j = j + 1
    
    
    Next i

    '外装資材分
    SEQNO = 0
    j = 0
    For i = ptxG_HIN_GAI01 To ptxG_HIN_GAI03 Step 5
    
        If Trim(Text1(i).text) = "" Then
        Else
            
            Call UniCode_Conv(P_SSHIJI_K_REC.SHIJI_No, Format(SHIJINO, "00000000"))        '指図票№
            SEQNO = SEQNO + 10
            Call UniCode_Conv(P_SSHIJI_K_REC.DATA_KBN, P_GAISOU)                        'データ区分
            Call UniCode_Conv(P_SSHIJI_K_REC.SEQNO, Format(SEQNO, "000"))               '追番
                        
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_SYUBETSU, "")                           '種別
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_JGYOBU, G_Item_Tbl(j).JGYOBU)           '事業部
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_NAIGAI, G_Item_Tbl(j).NAIGAI)           '国内外
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_HIN_GAI, Text1(i).text)                 '品番
                                                                                        '員数
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_QTY, Format(CDbl(Text1(i + 2).text), "000.00"))
                                                                                        '数量
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_SHIJI_QTY, Format(CDbl(Text1(i + 3).text), "00000000.00"))
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_BIKOU, "")                               '備考
            
                        
            
            Call UniCode_Conv(P_SSHIJI_K_REC.CALCEL_F, P_CANCEL_OFF)                    'ｷｬﾝｾﾙﾌﾗｸﾞ
            Call UniCode_Conv(P_SSHIJI_K_REC.CANCEL_DATETIME, "")                       'ｷｬﾝｾﾙ日時
        
        
        
 '--------------------------------------------------- 大阪  部材対応　2012.03.18
            
            
            If G_Item_Tbl(j).JGYOBU = SHIZAI Then                       '2013.03.31
                Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)                '2013.03.31
            Else                                                        '2013.03.31
                Call UniCode_Conv(K0_ITEM.JGYOBU, G_Item_Tbl(j).JGYOBU)
            End If                                                      '2013.03.31
            
            Call UniCode_Conv(K0_ITEM.NAIGAI, G_Item_Tbl(j).NAIGAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(i).text)
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Call UniCode_Conv(P_SSHIJI_K_REC.ST_TANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_DAN, vbUnicode))
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_SSHIJI_K_REC.ST_TANABAN, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    GoTo Abort_Tran
            End Select
'--------------------------------------------------- 大阪  部材対応　2012.03.18
        
        
        
            Call UniCode_Conv(P_SSHIJI_K_REC.FILLER, "")
                                                                                        '更新日時
            Call UniCode_Conv(P_SSHIJI_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
        
                                                                                        '出荷予定ＩＤ
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_ID_NO, "")
'--------------------------------------------------- 大阪  部材対応　2012.03.09
            For k = 0 To UBound(ZAIKO_FUSOKU)
            
            
                If ZAIKO_FUSOKU(k).JGYOBU = StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) And _
                    ZAIKO_FUSOKU(k).NAIGAI = StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode) And _
                    ZAIKO_FUSOKU(k).HIN_GAI = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode) Then
            
                    Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI, ZAIKO_FUSOKU(k).IDO_SUMI)
                    Call UniCode_Conv(P_SSHIJI_K_REC.HIKIATE_QTY, Format(ZAIKO_FUSOKU(k).HIKIATE_QTY, "00000000.00"))
                    
                    Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI_QTY, Format(ZAIKO_FUSOKU(k).IDO_SUMI_QTY, "00000000.00")) '2012.04.13
            
                    Exit For
            
                End If
            Next k
'--------------------------------------------------- 大阪  部材対応　2012.03.09
        
            Do
                
                DoEvents
                
                sts = BTRV(BtOpInsert, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<P_SSHIJI_K.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpInsert, "商品化指図ﾃﾞｰﾀ(子)")
                        GoTo Abort_Tran
                End Select
            
            Loop
        
        End If
        
        j = j + 1
    
    
    Next i


    '同梱／構成分
    SEQNO = 0
    For i = 0 To 49
    
        If D_Item_Tbl(i).JGYOBU = vbNullChar Or Trim(D_Item_Tbl(i).JGYOBU) = "" Then
        Else
            Call UniCode_Conv(P_SSHIJI_K_REC.SHIJI_No, Format(SHIJINO, "00000000"))        '指図票№
            
            SEQNO = SEQNO + 10
                                                                                        
            Call UniCode_Conv(P_SSHIJI_K_REC.DATA_KBN, P_DOUKON)                        'データ区分
            Call UniCode_Conv(P_SSHIJI_K_REC.SEQNO, Format(SEQNO, "000"))               '追番
                        
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_SYUBETSU, D_Item_Tbl(i).SYUBETSU)       '種別
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_JGYOBU, D_Item_Tbl(i).JGYOBU)           '事業部
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_NAIGAI, D_Item_Tbl(i).NAIGAI)           '国内外
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_HIN_GAI, D_Item_Tbl(i).HIN_GAI)         '品番
                                                                                        '員数
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_QTY, Format(D_Item_Tbl(i).QTY, "000.00"))
                                                                                        '数量
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_SHIJI_QTY, Format(D_Item_Tbl(i).SHIJI_QTY, "00000000.00"))
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_BIKOU, D_Item_Tbl(i).BIKOU)             '備考
        
            Call UniCode_Conv(P_SSHIJI_K_REC.CALCEL_F, P_CANCEL_OFF)                    'ｷｬﾝｾﾙﾌﾗｸﾞ
            Call UniCode_Conv(P_SSHIJI_K_REC.CANCEL_DATETIME, "")                       'ｷｬﾝｾﾙ日時
            
            
 '--------------------------------------------------- 大阪  部材対応　2012.03.18
            If D_Item_Tbl(i).JGYOBU = SHIZAI Then                       '2013.03.31
                Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)                '2013.03.31
            Else                                                        '2013.03.31
                Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(i).JGYOBU)
            End If                                                      '2013.03.31
            Call UniCode_Conv(K0_ITEM.NAIGAI, D_Item_Tbl(i).NAIGAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, D_Item_Tbl(i).HIN_GAI)
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Call UniCode_Conv(P_SSHIJI_K_REC.ST_TANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_DAN, vbUnicode))
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_SSHIJI_K_REC.ST_TANABAN, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    GoTo Abort_Tran
            End Select
'--------------------------------------------------- 大阪  部材対応　2012.03.18
            
            
            Call UniCode_Conv(P_SSHIJI_K_REC.FILLER, "")
                                                                                        '更新日時
            Call UniCode_Conv(P_SSHIJI_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
        
            If POS_UMU Then
                '出荷指示の作成
'''2007.03.08                If Y_SYUKA_Make_Proc(i) Then
'''2007.03.08                    GoTo Abort_Tran
'''2007.03.08                End If
            End If
        
                                                                                        '出荷予定ＩＤ
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_ID_NO, D_Item_Tbl(i).ID_NO)
        
        
'--------------------------------------------------- 大阪  部材対応　2012.03.09
            Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI, "")                  '2013.02.13
            Call UniCode_Conv(P_SSHIJI_K_REC.HIKIATE_QTY, "")               '2013.02.13
            Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI_QTY, "")              '2013.02.13
            
            
            If Trim(Text1(ptxORDER_NO).text) <> "" Then
                For k = 0 To UBound(ZAIKO_FUSOKU)
                
                
                    If ZAIKO_FUSOKU(k).JGYOBU = StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) And _
                        ZAIKO_FUSOKU(k).NAIGAI = StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode) And _
                        ZAIKO_FUSOKU(k).HIN_GAI = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode) Then
                
                        Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI, ZAIKO_FUSOKU(k).IDO_SUMI)
                        Call UniCode_Conv(P_SSHIJI_K_REC.HIKIATE_QTY, Format(ZAIKO_FUSOKU(k).HIKIATE_QTY, "00000000.00"))
                        
                        Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI_QTY, Format(ZAIKO_FUSOKU(k).IDO_SUMI_QTY, "00000000.00")) '2012.04.13
                
                        Exit For
                
                    End If
                Next k
            End If
'--------------------------------------------------- 大阪  部材対応　2012.03.09
        
        
        
            Do
                
                DoEvents
                
                sts = BTRV(BtOpInsert, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<P_SSHIJI_K.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpInsert, "商品化指図ﾃﾞｰﾀ(子)")
                        GoTo Abort_Tran
                End Select
            
            Loop
        
        End If
    
    Next i
    
    
'--------------------------------------------------- 大阪  部材対応　2012.04.13
    'If Trim(Text1(ptxORDER_NO).text) = "" Then         '2013.05.22 DEL
    If Trim(Text1(ptxORDER_NO).text) <> "" Then         '2013.05.22 INS
        Call UniCode_Conv(K6_ODR_ORDER.ORDER_NO, Text1(ptxORDER_NO).text)

        sts = BTRV(BtOpGetEqual, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K6_ODR_ORDER, Len(K6_ODR_ORDER), 6)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Beep
                MsgBox "注文データが変更されています。発注検討画面で確認してください。"
                GoTo Abort_Tran
            Case Else
                Call File_Error(sts, BtOpGetEqual, "注文データ")
                GoTo Abort_Tran
        End Select

        Call UniCode_Conv(ODR_ORDER_REC.PRT_FLG, "F")
        sts = BTRV(BtOpUpdate, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K6_ODR_ORDER, Len(K6_ODR_ORDER), 6)
        Select Case sts
            Case BtNoErr
            Case Else
                Call File_Error(sts, BtOpUpdate, "注文データ")
                GoTo Abort_Tran
        End Select
    End If
'--------------------------------------------------- 大阪  部材対応　2012.04.13



End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
'2007.11.21    If Mode = 0 Then
    If MSG = 0 Then     '2007.11.21
        If Text1(ptxSHIJI_NO).text = "" Then
            MsgBox "指図票№：" & Format(SHIJINO, "00000000") & "を作成しました。"
        End If
    End If
    
    Call Input_UnLock
                                        '印刷に対象指図票№を通知
    Taget_Key = Format(SHIJINO, "00000000")
    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    
    Call Input_UnLock

End Function

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    Select Case Index
        Case pcmbSHIMUKE        '仕向け先


        Case pcmbUKEHARAI       '手配先
            Text1(ptxUKEHARAI_CODE).text = Trim(Right(Combo1(pcmbUKEHARAI).text, 5))

        Case pcmbS_TANTO        '収単／担当者

                                '同梱／構成　種別
        Case pcmbD_SYUBETSU01, pcmbD_SYUBETSU02, pcmbD_SYUBETSU03, pcmbD_SYUBETSU04, pcmbD_SYUBETSU05, pcmbD_SYUBETSU06

            D_Item_Tbl(Index - pcmbD_SYUBETSU01).SYUBETSU = Right(Combo1(Index).text, 2)

    End Select

    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Sub Combo1_LostFocus(Index As Integer)

    Select Case Index
        Case pcmbSHIMUKE        '仕向け先

        Case pcmbUKEHARAI       '手配先
            Text1(ptxUKEHARAI_CODE).text = Trim(Right(Combo1(pcmbUKEHARAI).text, 5))

        Case pcmbS_TANTO        '収単／担当者

                                '同梱／構成　種別
        Case pcmbD_SYUBETSU01, pcmbD_SYUBETSU02, pcmbD_SYUBETSU03, pcmbD_SYUBETSU04, pcmbD_SYUBETSU05, pcmbD_SYUBETSU06

            D_Item_Tbl(Index - pcmbD_SYUBETSU01).SYUBETSU = Right(Combo1(Index).text, 2)

    End Select

End Sub

Private Sub Combo2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Call Tab_Ctrl(Shift)        '移動


End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans             As Integer
Dim i               As Integer

Dim rpt             As New PI00015F1
Dim f               As New PI000153

Dim com             As Integer
Dim sts             As Integer


Dim Parts_F         As Integer
Dim Gaisou_F        As Integer
Dim Kishu_F         As Integer

Dim objAccess       As Access.Application
Dim strAccessPath   As String

Dim GAISOU_QTY          As Long
Dim GAISOU_SHIJI_QYU    As Long

Dim L_print_Flg     As Boolean

'--------------------------------------------------- 大阪  部材対応　2012.03.08
Dim Order_QTY       As Long
Dim SHIJI_QTY       As Long

Dim ZAIKO_F         As Boolean
Dim wkMSG           As String
'--------------------------------------------------- 大阪  部材対応　2012.03.08


    Select Case Index
        Case P_CMD_Upd        '更新
            
            
            For i = ptxSHIJI_NO To ptxD_BIKOU06
            
                If Error_Check_Proc(i, 0, 1) Then   'エラーチェック
                    Exit Sub
                End If
            
            Next i
            
            
'--------------------------------------------------- 大阪  部材対応　2012.03.08
'            If ORDER_Check_Proc(Order_QTY, SHIJI_QTY) Then
'                Unload Me
'            End If
'            If Order_QTY < SHIJI_QTY Then
'
'                wkMSG = "親品番　注文情報と異なります。処理を継続しますか？" & Chr(13) & Chr(10)
'                wkMSG = wkMSG & "親品番　注文数:" & Format(Order_QTY, "#0") & Chr(13) & Chr(10)
'                wkMSG = wkMSG & "　　　　　 指示数:" & Format(SHIJI_QTY, "#0")
'
'
'                ans = MsgBox(wkMSG, vbYesNo + vbDefaultButton2, "確認入力")
'                If ans = vbNo Then
'                    Exit Sub
'                End If
'            End If

            
            
            If Trim(Text1(ptxORDER_NO).text) <> "" Then
                If Zaiko_Check_Proc(ZAIKO_F) Then
                    Unload Me
                End If
    
    
                If ZAIKO_F Then
                    wkMSG = "在庫不足が発生しています。処理を継続しますか？" & Chr(13) & Chr(10)
                    For i = 0 To UBound(ZAIKO_FUSOKU)
                        If ZAIKO_FUSOKU(i).SAI_QTY < 0 Then
                            wkMSG = wkMSG & RTrim(ZAIKO_FUSOKU(i).HIN_GAI) & Chr(13) & Chr(10)
                        End If
                    Next i
                    ans = MsgBox(wkMSG, vbYesNo + vbDefaultButton2, "確認入力")
                    If ans = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
'--------------------------------------------------- 大阪  部材対応　2012.03.08
            
            Beep
            ans = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Update_Proc(0, , 0) Then  '2007.11.21 引数変更
                    Unload Me
                End If
                
                If Init_Proc() Then
                    Unload Me
                End If
            
                Text1(ptxSHIJI_NO).SetFocus
            
            
            Else
                Text1(ptxORDER_NO).SetFocus
            End If


'        Case P_CMD_DEL                      '削除
        Case cmdMUPDATE                     'ﾏｽﾀ更新
        
            For i = ptxSHIJI_NO To ptxD_BIKOU06
            
                If Error_Check_Proc(i, 1, 1) Then   'エラーチェック
                    Exit Sub
                End If
            
            Next i
            
            Beep
            ans = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Update_Proc(1, , 1) Then  '2007.11.21引数変更
                    Unload Me
                End If
                
'                If Init_Proc() Then
'                    Unload Me
'                End If
            
                Call UniCode_Conv(ITEMREC.JGYOBU, "")
                Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, "")
            
            
                If Text1(ptxSHIJI_NO).Locked Then
                    Text1(ptxORDER_NO).SetFocus
                Else
                    Text1(ptxSHIJI_NO).SetFocus
                End If
            
            Else
                Text1(ptxORDER_NO).SetFocus
            End If
        
        
        Case P_CMD_DSP                      '検索/表示
        Case cmdNext                        '構成部品画面へ
        
            Doukon_Start = 1
            PI000152.Show vbModal           '部品詳細フォーム表示
            If G_SCREEN_FLG = SYS_ERR Then
                Unload Me
            End If
        
            'ﾃｰﾌﾞﾙより構成／同梱を表示
            If Tbl_To_Disp_Proc() Then
                Unload Me
            End If
        
        
        
        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        Case P_CMD_PRT                      '印刷
            
            
            
            For i = ptxSHIJI_NO To ptxD_BIKOU06
            
                If Error_Check_Proc(i, 0, 1) Then   'エラーチェック
                    Exit Sub
                End If
            
            Next i
'--------------------------------------------------- 大阪  部材対応　2012.03.08
'            If ORDER_Check_Proc(Order_QTY, SHIJI_QTY) Then
'                Unload Me
'            End If
'            If Order_QTY < SHIJI_QTY Then
'
'                wkMSG = "親品番　注文情報と異なります。処理を継続しますか？" & Chr(13) & Chr(10)
'                wkMSG = wkMSG & "親品番　注文数:" & Format(Order_QTY, "#0") & Chr(13) & Chr(10)
'                wkMSG = wkMSG & "　　　　　 指示数:" & Format(SHIJI_QTY, "#0")
'
'
'                ans = MsgBox(wkMSG, vbYesNo + vbDefaultButton2, "確認入力")
'                If ans = vbNo Then
'                    Exit Sub
'                End If
'            End If

            
            
            If Trim(Text1(ptxORDER_NO).text) <> "" Then
                If Zaiko_Check_Proc(ZAIKO_F) Then
                    Unload Me
                End If
    
    
                If ZAIKO_F Then
                    wkMSG = "在庫不足が発生しています。処理を継続しますか？" & Chr(13) & Chr(10)
                    For i = 0 To UBound(ZAIKO_FUSOKU)
                        If ZAIKO_FUSOKU(i).SAI_QTY < 0 Then
                            wkMSG = wkMSG & RTrim(ZAIKO_FUSOKU(i).HIN_GAI) & Chr(13) & Chr(10)
                        End If
                    Next i
                    ans = MsgBox(wkMSG, vbYesNo + vbDefaultButton2, "確認入力")
                    If ans = vbNo Then
                        Exit Sub
                    End If
                End If
            Else                                    '2012.12.20
                Erase ZAIKO_FUSOKU                  '2012.12.20
                ReDim ZAIKO_FUSOKU(0 To 0)          '2012.12.20
                ZAIKO_FUSOKU(0).JGYOBU = ""         '2012.12.20
                ZAIKO_FUSOKU(0).NAIGAI = ""         '2012.12.20
                ZAIKO_FUSOKU(0).HIN_GAI = ""        '2012.12.20
                
            
            End If
'--------------------------------------------------- 大阪  部材対応　2012.03.08
            
            Beep
            ans = MsgBox("印刷／更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Update_Proc(0, 1, 1) Then
                    Unload Me
                End If
                
                
                If Check1(pchkPRI_SHIJI).Value = vbChecked Then
                
                    Set rpt = New PI00015F1
                
                    'レポートを印刷します。（true：印刷ダイアログあり false：なし）
                    rpt.PrintReport False
                
                    Set rpt = Nothing


'                    f.RunReport rpt
'                    f.Show
                
                End If
                
                
                
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2013.11.20

                'ﾗﾍﾞﾙｼｽﾃﾑ印刷要求




                If Check1(pchkPRI_PARTS).Value = vbChecked Or _
                    Check1(pchkPRI_GAISOU).Value = vbChecked Then

                    L_print_Flg = True

                    
'>>>>>>>>>>>>>>>>   2016.01.13
'                    If L_URIKIN1 = 0 And L_URIKIN2 = 0 And L_URIKIN3 = 0 Then
'
'                        Beep
'                        ans = MsgBox("単価未設定です。ラベル印刷しますか？", vbYesNo + vbQuestion, "確認入力")
'                        If ans = vbYes Then
'                        Else
'                            L_print_Flg = False
'                        End If
'                    Else
'                    End If
'>>>>>>>>>>>>>>>>   2016.01.13
                    
                    If L_print_Flg Then


                        On Error Resume Next
                        Set objAccess = GetObject(, "Access.Application")
                        If Err().Number <> 0 Then
    '2016.01.13                        MsgBox "この端末では商品ラベル発行は行えません。"
    '                        MsgBox "GetObject(Access.Application)" & Err().Number & " " & Err().Description
                        Else
    '                        MsgBox Err.Number

                            strAccessPath = App.Path
                            If Right(strAccessPath, 1) <> "\" Then
                                strAccessPath = strAccessPath & "\"
                            End If

                            strAccessPath = strAccessPath & "litem.mdb"
                            Set objAccess = GetObject(strAccessPath)



                            If Check1(pchkPRI_PARTS).Value = vbChecked Then
                                Parts_F = 1
                            Else
                                Parts_F = 0
                            End If


                            If Check1(pchkPRI_GAISOU).Value = vbChecked Then
                                Gaisou_F = 1
                            Else
                                Gaisou_F = 0
                            End If

                            If G_Kisyu_F = L_LABEL_ON Then
                                Kishu_F = 1
                                Parts_F = 0
                            Else
                                Kishu_F = 0
                            End If

                            If IsNumeric(Text1(ptxG_QTY01).text) Then
                                GAISOU_QTY = CLng(Text1(ptxG_QTY01).text)
                            Else
                                GAISOU_QTY = 0
                            End If
                            If IsNumeric(Text1(ptxG_SHIJI_QTY01).text) Then
                                GAISOU_SHIJI_QYU = CLng(Text1(ptxG_SHIJI_QTY01).text)
                            Else
                                GAISOU_SHIJI_QYU = 0
                            End If

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

                            Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                            Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)


                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            Select Case sts
                                Case BtNoErr


                                    Call UniCode_Conv(ITEMREC.L_IRI_QTY, Format(GAISOU_QTY, "00000000"))



                                    '再梱包ﾏｰｸ追加  2007.11.09

                                    If Option1(poptSHIJI_SAIKON).Value Then
                                        Call UniCode_Conv(ITEMREC.L_MARK, SAIKON_F)

                                    End If





                                    sts = BTRV(BtOpInsert, L_ITEM_POS, ITEMREC, Len(ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)

                                    Select Case sts
                                        Case BtNoErr

                                            objAccess.Run "PosPrintLabel", Trim(Text1(ptxHIN_GAI).text), CLng(Text1(ptxSHIJI_QTY).text), Parts_F, Gaisou_F, Kishu_F, GAISOU_QTY, GAISOU_SHIJI_QYU, 0



                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                            Exit Sub


                                    End Select

                                Case BtErrKeyNotFound

                                Case Else
                                   Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                    Exit Sub

                            End Select





                            Set objAccess = Nothing
                        End If



                    End If
                End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2013.11.20
                
                If Init_Proc() Then
                    Unload Me
                End If
            
'2007.11.21                Text1(ptxSHIJI_NO).SetFocus
                
'--------------------------------------------------- 大阪  部材対応　2012.03.18
'                Text1(ptxHIN_GAI).SetFocus  '2007.11.21
                Text1(ptxORDER_NO).SetFocus
'--------------------------------------------------- 大阪  部材対応　2012.03.18
            Else
                Text1(ptxORDER_NO).SetFocus
            End If
            
            
            
        Case cmdCen                         '取り消し
            If Init_Proc() Then
                Unload Me
            End If
            Text1(ptxSHIJI_NO).SetFocus
        Case P_CMD_End                      '終了
            Unload Me
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
            Command1(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

Dim c           As String * 128
Dim sts         As Integer
Dim i           As Integer

Dim MUKE_CODE   As Variant

Dim YOMI_JGYOBU_w As Variant    '2014.03.24
Dim j           As Integer      '2014.03.24


    If App.PrevInstance Then
        MsgBox "同一プログラム実行中です。"
        End
    End If
                                
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止します。"
        End
    End If
    LOG_F = RTrim(c)
                                '出荷ログファイル名取り込み
    If SYUKA_LOGF_GET_PROC() Then
        Beep
        MsgBox "出荷ログファイル名の獲得に失敗しました。処理を中止します。"
        End
    End If
                                

    PI000151.Caption = PI000151.Caption & LAST_UPDATE_DAY       '2017.10.17

                                
                                
                                
                                
                                
                                
                                
                                '調査用ログファイル名取り込み   2016.03.30
    If GetIni(App.EXEName, "PI00015_LOG", App.EXEName, c) Then
        PI00015_LOG = ""
    Else
        PI00015_LOG = Trim(c)
    End If
                                
                                
                                
                                
                                '事業部の獲得       2016.01.27
    If JGYOB_TB_Set() Then
        MsgBox "事業部の獲得に失敗しました。"
        End
    End If
                                
                                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    P_SYS.INI　--＞　PI00015.INI    2016.01.13
                                
                                
                                '手配先取り込み
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "TEHAI", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "TEHAI", App.EXEName, c) Then
    Else
        TEHAI = RTrim(c)
    End If
                                
                                'POSｼｽﾃﾑ有無の取り込み
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "POS_UMU", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "POS_UMU", App.EXEName, c) Then
        POS_UMU = False
    Else
        If RTrim(c) = "0" Then
            POS_UMU = False
        Else
            POS_UMU = True
        End If
    End If
                                'ﾊﾞｰｺｰﾄﾞ印字
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "BCR", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "BCR", App.EXEName, c) Then
        PRI_MAIN_BCR = False
    Else
        If RTrim(c) = "0" Then
            PRI_MAIN_BCR = False
        Else
            If Not POS_UMU Then
                MsgBox "ＰＯＳｼｽﾃﾑが未設定です。処理を中止します。"
                End
            End If
            PRI_MAIN_BCR = True
        End If
    End If
                                    '明細備考印字内容
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "DET_BIKOU", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "DET_BIKOU", App.EXEName, c) Then
        PRI_BIKOU_BCR = 0
    Else
        
        If Not IsNumeric(RTrim(c)) Then
            PRI_BIKOU_BCR = 0
        Else
            Select Case RTrim(c)
                Case "0", "2"
                    PRI_BIKOU_BCR = CInt(RTrim(c))
                Case "1"
                    If Not POS_UMU Then
                        MsgBox "ＰＯＳｼｽﾃﾑが未設定です。処理を中止します。"
                        End
                    Else
                        PRI_BIKOU_BCR = CInt(RTrim(c))
                    End If
                Case Else
                    PRI_BIKOU_BCR = 0
            End Select
        
        End If
    End If
                                '収単／担当者の取り込み
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "S_TANTO", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "S_TANTO", App.EXEName, c) Then
        PRI_S_TANTO = False
    Else
        If RTrim(c) = "0" Then
            PRI_S_TANTO = False
        Else
            PRI_S_TANTO = True
        End If
    End If
                                
                                '作業日／数量／担当 2007.05.22
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "SAGYO_DAY", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "SAGYO_DAY", App.EXEName, c) Then
        PRI_SAGYO_DAY = False
    Else
        If RTrim(c) = "0" Then
            PRI_SAGYO_DAY = False
        Else
            PRI_SAGYO_DAY = True
        End If
    End If
                                
                                
                                
                                '商品検査　同梱の取り込み
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "DOUKON", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "DOUKON", App.EXEName, c) Then
        PRI_DOUKON = False
    Else
        If RTrim(c) = "0" Then
            PRI_DOUKON = False
        Else
            PRI_DOUKON = True
        End If
    End If
                                '入庫完了印の取り込み
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "NYUKO_IN", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "NYUKO_IN", App.EXEName, c) Then
        PRI_NYUKO_IN = False
    Else
        If RTrim(c) = "0" Then
            PRI_NYUKO_IN = False
        Else
            PRI_NYUKO_IN = True
        End If
    End If
                                '入力完了印の取り込み
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "INPUT_IN", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "INPUT_IN", App.EXEName, c) Then
        PRI_INPUT_IN = False
    Else
        If RTrim(c) = "0" Then
            PRI_INPUT_IN = False
        Else
            PRI_INPUT_IN = True
        End If
    End If
    
    '下部　品番／№／数量   2007.05.22
    If PRI_NYUKO_IN Or PRI_INPUT_IN Then
    Else
'        If GetIni(StrConv(App.EXEName, vbUpperCase), "HINBAN_BIKOU", "P_SYS", c) Then
        If GetIni(StrConv(App.EXEName, vbUpperCase), "HINBAN_BIKOU", App.EXEName, c) Then
            PRI_HINBAN_BIKOU = False
        Else
            If RTrim(c) = "0" Then
                PRI_HINBAN_BIKOU = False
            Else
                PRI_HINBAN_BIKOU = True
            End If
        End If
    End If
    
    
                                '自責
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "JISEKI", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "JISEKI", App.EXEName, c) Then
        JISEKI_TITLE = ""
    Else
        JISEKI_TITLE = Split(Trim(c), ",", -1)
    End If
    
                                '他責
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "TASEKI", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "TASEKI", App.EXEName, c) Then
        TASEKI_TITLE = ""
    Else
        TASEKI_TITLE = Split(Trim(c), ",", -1)
    End If
    
                                '未登録品番の可否
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "HIN_INV", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "HIN_INV", App.EXEName, c) Then
        HIN_INV = False
    Else
        If Trim(c) = "0" Then
            HIN_INV = False
        Else
            HIN_INV = True
        End If
    
    End If
    
    
    
    If PRI_BIKOU_BCR = 1 Then
                                    '向け先
'        If GetIni(StrConv(App.EXEName, vbUpperCase), "MTSSS", "P_SYS", c) Then
        If GetIni(StrConv(App.EXEName, vbUpperCase), "MTSSS", App.EXEName, c) Then
            MTS_CODE = ""
            SS_CODE = ""
        Else
        
            MUKE_CODE = Split(Trim(c), ",", -1)
        
            Select Case UBound(MUKE_CODE)
                Case 0
                
                    MTS_CODE = CStr(MUKE_CODE(0))
                    SS_CODE = ""
                Case 1
                    MTS_CODE = CStr(MUKE_CODE(0))
                    SS_CODE = CStr(MUKE_CODE(1))
                Case Else
                    MTS_CODE = ""
                    SS_CODE = ""
            End Select
        End If
                                    '向け先管理マスタのチェック
        If MTS_Open(BtOpenNomal) Then
            Unload Me
        End If
                                    
        Call UniCode_Conv(K0_MTS.MUKE_CODE, MTS_CODE)
        Call UniCode_Conv(K0_MTS.SS_CODE, SS_CODE)
                                            
        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                MsgBox "向け先が未設定です。処理を中止します。"
                                            '向け先管理マスタＣＬＯＳＥ
                sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                If sts Then
                    If sts <> BtErrNoOpen Then
                        Call File_Error(sts, BtOpClose, "向け先管理マスタ")
                    End If
                End If
                End
            
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                                            '向け先管理マスタＣＬＯＳＥ
                sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                If sts Then
                    If sts <> BtErrNoOpen Then
                        Call File_Error(sts, BtOpClose, "向け先管理マスタ")
                    End If
                End If
                End
        End Select
                                            '向け先管理マスタＣＬＯＳＥ
        sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        If sts Then
            If sts <> BtErrNoOpen Then
                Call File_Error(sts, BtOpClose, "向け先管理マスタ")
            End If
        End If
                                            '注文区分の獲得
'        If GetIni(StrConv(App.EXEName, vbUpperCase), "CYU_KBN", "P_SYS", c) Then
        If GetIni(StrConv(App.EXEName, vbUpperCase), "CYU_KBN", App.EXEName, c) Then
            CYU_KBN = ""
        Else
            CYU_KBN = Trim(c)
        End If
        
        
        
        Select Case CYU_KBN
            Case CYU_KBN_TUK            '月切
                CYU_KBN_N = CYU_KBN_1
            Case CYU_KBN_SPO
                CYU_KBN_N = CYU_KBN_2   '緊急
            Case CYU_KBN_HJU
                CYU_KBN_N = CYU_KBN_3   '補充
            Case CYU_KBN_TOK
                CYU_KBN_N = CYU_KBN_4   '特売
            Case CYU_KBN_BOU
                CYU_KBN_N = CYU_KBN_E   '貿易
            Case Else
            MsgBox "注文区分が未設定です。処理を中止します。"
            End
        End Select
    Else
        MTS_CODE = ""
        SS_CODE = ""
        CYU_KBN = ""
    End If
                                
                                
                                        '再梱包の獲得   2007.11.09
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "SAIKON_F", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "SAIKON_F", App.EXEName, c) Then
        SAIKON_F = ""
    Else
        SAIKON_F = Trim(c)
    End If
                                
                                
                                
'--------------------------------------------------- 大阪  部材対応　2012.03.20
    Jyogai_Soko_umu = False
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "JYOGAI_SOKO", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "JYOGAI_SOKO", App.EXEName, c) Then
    Else
        Jyogai_Soko_umu = True
        Zaiko_Syukei_Jyogai_Soko_No = Split(Trim(c), ",", -1)
    End If



'    If GetIni(StrConv(App.EXEName, vbUpperCase), "B", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "B", App.EXEName, c) Then
        YUKO_JGYOBU = "B"
    Else
        YUKO_JGYOBU = Trim(c)
    End If



'--------------------------------------------------- 大阪  部材対応　2012.03.20
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "YOMI_JGYOBU", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "YOMI_JGYOBU", App.EXEName, c) Then
        Call DEF_JGYOBU_PROC
    Else
        YOMI_JGYOBU_w = Split(Trim(c), ",", -1)
        For j = 0 To UBound(YOMI_JGYOBU_w)
            ReDim Preserve YOMI_JGYOBU(j)
            YOMI_JGYOBU(j) = YOMI_JGYOBU_w(j)
        Next j
    End If
'--------------------------------------------------- 事業部読み替え順 2014.03.24




                                
                                
                                
                                '発番マスタＯＰＥＮ
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '商品ﾗﾍﾞﾙ用品目マスタＯＰＥＮ
    If L_ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                'クラスマスタＯＰＥＮ
    If P_Class_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '構成マスタＯＰＥＮ
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '管理マスタＯＰＥＮ
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '商品化指図（子）ﾃﾞｰﾀＯＰＥＮ
    If P_SSHIJI_K_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '商品化指図（親）ﾃﾞｰﾀＯＰＥＮ
    If P_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定ﾃﾞｰﾀＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '受払先マスタＯＰＥＮ
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫ﾃﾞｰﾀＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '商品化指図（親）ﾜｰｸＯＰＥＮ
    If wP_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
                                '発注検討　親品番注文ﾌｧｲﾙＯＰＥＮ   2012.03.08
    If ODR_ORDER_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    
    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc
    
    
    
    
    
    
    
    
    Load PI000152
    Load PI000153
    
    
    
    '管理マスタの読み込み
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)

    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            If P_KANRI_MAKE_Proc() Then
                Unload Me
            End If
        Case Else
            Call File_Error(sts, BtOpGetEqual, "管理マスタ")
            Unload Me
    End Select
        
    
    
    '仕向け先のセット
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD, 0) Then
        Unload Me
    End If
    '収単／担当者のセット
    If Code_Set_Proc(pcmbS_TANTO, P_KBN05_CD, 0) Then
        Unload Me
    End If
    
    '受払先
    If Ukeharai_Set_Proc() Then
        Unload Me
    End If
    
    
    Doukon_Tbl_No(0) = "①"
    Doukon_Tbl_No(1) = "②"
    Doukon_Tbl_No(2) = "③"
    Doukon_Tbl_No(3) = "④"
    Doukon_Tbl_No(4) = "⑤"
    Doukon_Tbl_No(5) = "⑥"
    Doukon_Tbl_No(6) = "⑦"
    Doukon_Tbl_No(7) = "⑧"
    Doukon_Tbl_No(8) = "⑨"
    Doukon_Tbl_No(9) = "⑩"
    Doukon_Tbl_No(10) = "⑪"
    Doukon_Tbl_No(11) = "⑫"
    Doukon_Tbl_No(12) = "⑬"
    Doukon_Tbl_No(13) = "⑭"
    Doukon_Tbl_No(14) = "⑮"
    Doukon_Tbl_No(15) = "⑯"
    Doukon_Tbl_No(16) = "⑰"
    Doukon_Tbl_No(17) = "⑱"
    Doukon_Tbl_No(18) = "⑲"
    Doukon_Tbl_No(19) = "⑳"
    
    
    
    '種別のセット
    For i = pcmbD_SYUBETSU01 To pcmbD_SYUBETSU06
        If Code_Set_Proc(i, P_KBN06_CD, 1) Then
            Unload Me
        End If
    Next i
    
    '画面初期設定
    If Init_Proc() Then
        Unload Me
    End If
    Combo1(pcmbSHIMUKE).ListIndex = 0                   '2007.11.01


    '指示形態       2007.11.01
    Option1(poptSHIJI_NORMAL).Value = True
    Option1(poptSHIJI_SPOT).Value = False
    Option1(poptSHIJI_KEPPIN).Value = False




    




End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts     As Integer
Dim ans     As Integer      '2012.03.09


    ans = MsgBox("処理を終了しますか？", vbYesNo + vbDefaultButton1, "確認入力")
    If ans = vbNo Then
        Cancel = True
        Exit Sub
    End If
                                            
                                            '発番マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "発番マスタ")
        End If
    End If
                                            
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '商品ﾗﾍﾞﾙ用品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品ﾗﾍﾞﾙ用品目マスタ")
        End If
    End If
    
                                            'クラスマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "クラスマスタ")
        End If
    End If
    
                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "コードマスタ")
        End If
    End If
    
                                            '構成マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "構成マスタ")
        End If
    End If
                                            '管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "管理マスタ")
        End If
    End If
                                            '商品化指図ﾃﾞｰﾀ(親)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化指図ﾃﾞｰﾀ(親)")
        End If
    End If
                                            '商品化指図ﾃﾞｰﾀ(子)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化指図ﾃﾞｰﾀ(子)")
        End If
    End If
    
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
    
                                            '出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定データ")
        End If
    End If
                                            '受払先マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "受払先マスタ")
        End If
    End If
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
    
                                            '商品化指図ﾜｰｸ(親)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, wP_SSHIJI_O_POS, wP_SSHIJI_O_REC, Len(wP_SSHIJI_O_REC), K0_wP_SSHIJI_O, Len(K0_wP_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化指図(親)ﾜｰｸ")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PI000151 = Nothing
    Set PI000152 = Nothing
    Set PI000153 = Nothing

    End
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).text)
    End If

    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2013.12.28
    Select Case Index
        Case ptxHIN_GAI
            svHin_Gai = Text1(Index).text
    End Select
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2013.12.28

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
'--------------------------------------------------- 大阪  部材対応　2012.03.09
    Select Case Index
        Case ptxHIN_GAI, ptxK_HIN_GAI01, ptxK_HIN_GAI02, ptxK_HIN_GAI03, ptxK_HIN_GAI04, ptxK_HIN_GAI05, _
                ptxG_HIN_GAI01, ptxG_HIN_GAI02, ptxG_HIN_GAI03, _
                ptxD_HIN_GAI01, ptxD_HIN_GAI02, ptxD_HIN_GAI03, ptxD_HIN_GAI04, ptxD_HIN_GAI05, ptxD_HIN_GAI06 _

            Text1(Index).text = StrConv(Text1(Index).text, vbUpperCase)
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2013.12.28
            If Index = ptxHIN_GAI Then
                svHin_Gai = Text1(Index).text
            End If
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2013.12.28
    
    End Select
'--------------------------------------------------- 大阪  部材対応　2012.03.09
        
        
        
        
        
        
    If Error_Check_Proc(Index, 0, 0) Then   'エラーチェック
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動
End Sub

Private Function Init_Proc() As Integer
'----------------------------------------------------------------------------
'                   入力画面の初期設定
'----------------------------------------------------------------------------
Dim i           As Integer
Dim j           As Integer      '2016.01.27
Dim sts         As Integer

Dim TANTO_CODE  As String
Dim TANTO_NAME  As String

Dim SHONIN_CODE     As String   '2007.11.21
Dim SHONIN_NAME     As String   '2007.11.21


    Init_Proc = True
    
    Text1(ptxSHIJI_NO).BackColor = G_INPUT_OK
    Text1(ptxSHIJI_NO).Locked = False
    Text1(ptxSHIJI_NO).TabStop = True
    
    
    Combo1(pcmbS_TANTO).Enabled = PRI_S_TANTO
    
    TANTO_CODE = Text1(ptxTANTO_CODE).text
    TANTO_NAME = Text1(ptxTANTO_NAME).text
    
    SHONIN_CODE = Text1(ptxSHONIN_CODE).text    '2007.11.21
    SHONIN_NAME = Text1(ptxSHONIN_NAME).text    '2007.11.21
    
    
    
    For i = ptxSHIJI_NO To ptxD_BIKOU06
        
        If i = ptxUKEHARAI_CODE Then '2007.11.01
        Else
            Text1(i).text = ""
        End If
    Next i
    Text1(ptxTANTO_CODE).text = TANTO_CODE
    Text1(ptxTANTO_NAME).text = TANTO_NAME




    RichTextBox1(prchBIKOU).text = ""

    For i = pchkSAMPLE_F To pchkPRI_KISHU              '2013.11.20
        Check1(i).Value = vbChecked                    '2013.11.20
    Next i                                             '2013.11.20
    
'    For i = pchkSAMPLE_F To pchkPRI_SHIJI               '2013.11.20
'        Check1(i).Value = vbChecked                     '2013.11.20
'    Next i                                              '2013.11.20
    
    
    
    
    

    
    Check1(pchkSAMPLE_F).Value = vbUnchecked
    Check1(pchkPRI_KISHU).Value = vbUnchecked

    For i = pcmbSHIMUKE To pcmbD_SYUBETSU06
        
        If i = pcmbSHIMUKE Or i = pcmbUKEHARAI Then     '2007.11.01
        Else
            Combo1(i).ListIndex = -1
        End If
    Next i
'    Combo1(pcmbSHIMUKE).ListIndex = 0                  2007.11.01
    

'--------------------------------------------------- 大阪  部材対応　2012.03.18
    '受注日
'    Text1(ptxORDER_NO).text = Format(Now, "YYYY/MM/DD")
'--------------------------------------------------- 大阪  部材対応　2012.03.18


    '発行日
'''    Text1(ptxHAKKO_DT).text = Format(Now, "YYYY/MM/DD")
    Text1(ptxHAKKO_DT).text = ""


    '承認者設定
    If Trim(SHONIN_CODE) = "" Then      '2007.11.21
    
    
        Text1(ptxSHONIN_CODE).text = StrConv(P_KANRIREC.SHONIN_CODE, vbUnicode)
    
        Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(P_KANRIREC.SHONIN_CODE, vbUnicode))
    
        sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
                Text1(ptxSHONIN_NAME).text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
            Case BtErrKeyNotFound
                Text1(ptxSHONIN_NAME).text = ""
        
            Case Else
                Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                Exit Function
        End Select
    Else
        Text1(ptxSHONIN_CODE).text = SHONIN_CODE    '2007.11.21
        Text1(ptxSHONIN_NAME).text = SHONIN_NAME    '2007.11.21
    End If
    '手配先
    If Trim(Text1(ptxUKEHARAI_CODE).text) = "" Then '2007.11.01
        Text1(ptxUKEHARAI_CODE).text = TEHAI
    End If

    '指示形態
'2007.11.01    Option1(poptSHIJI_NORMAL).Value = True
'2007.11.01    Option1(poptSHIJI_SPOT).Value = False
'2007.11.01    Option1(poptSHIJI_KEPPIN).Value = False

    '出力対象
    Erase K_Item_Tbl
    Erase G_Item_Tbl
    Erase D_Item_Tbl
    
    ReDim K_Item_Tbl(0 To 4)
    ReDim G_Item_Tbl(0 To 2)
    ReDim D_Item_Tbl(0 To 49)

    For i = 0 To UBound(K_Item_Tbl)
        K_Item_Tbl(i).JGYOBU = ""
        K_Item_Tbl(i).NAIGAI = ""
    Next i

    For i = 0 To UBound(G_Item_Tbl)
        G_Item_Tbl(i).JGYOBU = ""
        G_Item_Tbl(i).NAIGAI = ""
    Next i

    For i = 0 To UBound(D_Item_Tbl)
        D_Item_Tbl(i).JGYOBU = ""
        D_Item_Tbl(i).NAIGAI = ""
        D_Item_Tbl(i).HIN_GAI = ""
    
    Next i


    Call UniCode_Conv(ITEMREC.JGYOBU, "")
    Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, "")



'--------------------------------------------------- 事業部　セット 2016.01.27
    For i = 3 To 8
    
        Combo2(i).Clear
    
        For j = 0 To UBound(JGYOBU_T)
        
            Combo2(i).AddItem JGYOBU_T(j).NAME & Space(10) & JGYOBU_T(j).CODE
        
        Next j
    
    
    Next i
'--------------------------------------------------- 事業部　セット 2016.01.27

'--------------------------------------------------- 大阪  部材対応　2012.03.18
    Call Disp_Lock_Proc(False)
'--------------------------------------------------- 大阪  部材対応　2012.03.18

    Init_Proc = False

End Function
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

Private Function Ukeharai_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   受払先マスタをコンボにセットする。
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer




Dim i           As Integer
    
    Ukeharai_Set_Proc = True
    
    Combo1(pcmbUKEHARAI).Clear
    
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, "")

    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "受払先マスタ")
                Exit Function
        
        End Select

        
        
        Combo1(pcmbUKEHARAI).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode) & "            " & _
                                StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        
        com = BtOpGetNext
    
    Loop

    Ukeharai_Set_Proc = False
    



End Function



Private Function P_SSHIJI_Read_Proc() As Integer
'----------------------------------------------------------------------------
'                   指図データの読み込み
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
    
    
    P_SSHIJI_Read_Proc = True
    
    
    '指図票ﾃﾞｰﾀ（親）
    Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, Text1(ptxSHIJI_NO))
    sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case Else
            P_SSHIJI_Read_Proc = sts
            Exit Function
    
    End Select
    
    
    If Item_Disp_Proc() Then
        Exit Function
    End If
    
    P_SSHIJI_Read_Proc = False
        
    

End Function


Private Function P_COMPO_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   構成マスタの読み込み＆表示
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim k           As Integer
Dim g           As Integer
Dim d           As Integer

Dim K_Index     As Integer
Dim G_Index     As Integer
Dim DT_Index    As Integer
Dim DC_Index    As Integer


Dim Mi_Qty      As Long
Dim Sumi_Qty    As Long
    
Dim wkJgyobu    As String * 1
    
    
    P_COMPO_Disp_Proc = True
    
    For i = ptxK_HIN_GAI01 To ptxD_BIKOU06
        Text1(i).text = ""
    Next i
            
            
    For i = pcmbD_JGYOBU01 To pcmbD_JGYOBU06        '2016.04.07
        Combo2(i).ListIndex = -1                    '2016.04.07
    Next                                            '2016.04.07
            
            
    '出力対象
    Erase K_Item_Tbl
    Erase G_Item_Tbl
    Erase D_Item_Tbl
    
    ReDim K_Item_Tbl(0 To 4)
    ReDim G_Item_Tbl(0 To 2)
    ReDim D_Item_Tbl(0 To 49)

    For i = 0 To UBound(K_Item_Tbl)
        K_Item_Tbl(i).JGYOBU = ""
        K_Item_Tbl(i).NAIGAI = ""
    Next i

    For i = 0 To UBound(G_Item_Tbl)
        G_Item_Tbl(i).JGYOBU = ""
        G_Item_Tbl(i).NAIGAI = ""
    Next i

    For i = 0 To UBound(D_Item_Tbl)
        D_Item_Tbl(i).JGYOBU = ""
        D_Item_Tbl(i).NAIGAI = ""
        D_Item_Tbl(i).HIN_GAI = ""
    
    Next i
    
    
    
    
    
    Text1(ptxS_CLASS_CODE).text = ""
    Text1(ptxF_CLASS_CODE).text = ""
    Text1(ptxN_CLASS_CODE).text = ""

    
    
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).text)
       
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_HEAD)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
       
    sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case Else
            
            For i = ptxK_HIN_GAI01 To ptxD_BIKOU06
                Text1(i).text = ""
            Next i
            
            '出力対象
            Erase K_Item_Tbl
            Erase G_Item_Tbl
            Erase D_Item_Tbl
            
            ReDim K_Item_Tbl(0 To 4)
            ReDim G_Item_Tbl(0 To 2)
            ReDim D_Item_Tbl(0 To 49)
        
            For i = 0 To UBound(K_Item_Tbl)
                K_Item_Tbl(i).JGYOBU = ""
                K_Item_Tbl(i).NAIGAI = ""
            Next i
        
            For i = 0 To UBound(G_Item_Tbl)
                G_Item_Tbl(i).JGYOBU = ""
                G_Item_Tbl(i).NAIGAI = ""
            Next i
        
            For i = 0 To UBound(D_Item_Tbl)
                D_Item_Tbl(i).JGYOBU = ""
                D_Item_Tbl(i).NAIGAI = ""
                D_Item_Tbl(i).HIN_GAI = ""
            
            Next i
            
            Text1(ptxS_CLASS_CODE).text = ""
            Text1(ptxF_CLASS_CODE).text = ""
            Text1(ptxN_CLASS_CODE).text = ""
            
            
            P_COMPO_Disp_Proc = sts
            Exit Function
    End Select

    '商品ｸﾗｽ
    Text1(ptxS_CLASS_CODE).text = Trim(StrConv(P_COMPO_O_REC.CLASS_CODE, vbUnicode))
    '付加ｸﾗｽ
    Text1(ptxF_CLASS_CODE).text = Trim(StrConv(P_COMPO_O_REC.F_CLASS_CODE, vbUnicode))
    '内職ｸﾗｽ
    Text1(ptxN_CLASS_CODE).text = Trim(StrConv(P_COMPO_O_REC.N_CLASS_CODE, vbUnicode))
    RichTextBox1(prchBIKOU) = Trim(StrConv(P_COMPO_O_REC.BIKOU, vbUnicode))
    '--------------------------------   「子」情報
    Erase K_Item_Tbl
    Erase G_Item_Tbl
    Erase D_Item_Tbl
    
    ReDim K_Item_Tbl(0 To 4)
    ReDim G_Item_Tbl(0 To 2)
    ReDim D_Item_Tbl(0 To 49)
    
'--------------------------------------------------- 大阪  部材対応　2012.03.06
    For i = 0 To UBound(K_Item_Tbl)
        K_Item_Tbl(i).JGYOBU = ""
        K_Item_Tbl(i).NAIGAI = ""
    Next i

    For i = 0 To UBound(G_Item_Tbl)
        G_Item_Tbl(i).JGYOBU = ""
        G_Item_Tbl(i).NAIGAI = ""
    Next i

    For i = 0 To UBound(D_Item_Tbl)
        D_Item_Tbl(i).JGYOBU = ""
        D_Item_Tbl(i).NAIGAI = ""
        D_Item_Tbl(i).HIN_GAI = ""
    
    Next i
'--------------------------------------------------- 大阪  部材対応　2012.03.06
    
    
    
    k = -1
    g = -1
    d = -1
    
    K_Index = ptxK_HIN_GAI01
    G_Index = ptxG_HIN_GAI01
    DT_Index = ptxD_HIN_GAI01
    DC_Index = pcmbD_SYUBETSU01
    
    Do
        
        
        sts = BTRV(BtOpGetNext, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
            
                            
                If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                    StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                    StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                    Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHIN_GAI).text) Then
                
                    Exit Do
            
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetNext, "構成マスタ")
                Exit Function
        
        
        End Select
        
        Select Case StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode)
        
            Case P_KOSOU    '個装資材
            
                k = k + 1
                K_Item_Tbl(k).JGYOBU = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
                K_Item_Tbl(k).NAIGAI = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
                            '品番
                Text1(K_Index).text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                
                
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                Select Case K_Item_Tbl(k).JGYOBU
                    Case SHIZAI
                        Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
                    Case SETSUBI
                        Call UniCode_Conv(K0_ITEM.JGYOBU, YUKO_JGYOBU)
                    Case Else
                        Call UniCode_Conv(K0_ITEM.JGYOBU, K_Item_Tbl(k).JGYOBU)
                End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                Call UniCode_Conv(K0_ITEM.NAIGAI, K_Item_Tbl(k).NAIGAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        '品名
                        Text1(K_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                        '標準棚番
                        If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                            Text1(K_Index + 4) = ""
                        Else
                            Text1(K_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_DAN, vbUnicode)
                        End If
                    
                    Case BtErrKeyNotFound
                        
                        
'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21
                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))

                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                                '品名
                                Text1(K_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                '標準棚番
                                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                                    Text1(K_Index + 4) = ""
                                Else
                                    Text1(K_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_DAN, vbUnicode)
                                End If
                            
                            Case BtErrKeyNotFound


                                Text1(K_Index + 1) = "未登録品番"
                                Text1(K_Index + 4) = ""
                            Case Else
                                Call Input_UnLock             '2008.01.15
                                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                Exit Function

                        End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21
                        
                        
                        
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "")
                        Exit Function
                
                End Select
            
            
                Text1(K_Index + 2).text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                
                If IsNumeric(Text1(ptxSHIJI_QTY).text) Then
                    Text1(K_Index + 3).text = Format(CDbl(CLng(Text1(ptxSHIJI_QTY).text) * CDbl(Text1(K_Index + 2).text)), "#0.00")
                Else
                    Text1(K_Index + 3).text = ""
                End If
                            
                K_Index = K_Index + 5
            
            
            
            Case P_GAISOU   '外装資材
                g = g + 1
                G_Item_Tbl(g).JGYOBU = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
                G_Item_Tbl(g).NAIGAI = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
                            '品番
                Text1(G_Index).text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                Select Case G_Item_Tbl(g).JGYOBU
                    Case SHIZAI
                        Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
                    Case SETSUBI
                        Call UniCode_Conv(K0_ITEM.JGYOBU, YUKO_JGYOBU)
                    Case Else
                        Call UniCode_Conv(K0_ITEM.JGYOBU, G_Item_Tbl(g).JGYOBU)
                End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                Call UniCode_Conv(K0_ITEM.NAIGAI, G_Item_Tbl(g).NAIGAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(G_Index).text)
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        '品名
                        Text1(G_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                        '標準棚番
                        If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                            Text1(G_Index + 4) = ""
                        Else
                            Text1(G_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_DAN, vbUnicode)
                        End If
                    
                    Case BtErrKeyNotFound
                        
                        
'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21
                            Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
    
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            Select Case sts
                                Case BtNoErr
                                    '品名
                                    Text1(G_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                    '標準棚番
                                    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                                        Text1(G_Index + 4) = ""
                                    Else
                                        Text1(G_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                                    End If
                                
                                Case BtErrKeyNotFound
    
                                    Text1(G_Index + 1) = "未登録品番"
                                    Text1(G_Index + 4) = ""
                                Case Else
                                    Call Input_UnLock             '2008.01.15
                                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                    Exit Function
    
                            End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21
                        
                        
                        
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "")
                        Exit Function
                
                End Select
            
            
                Text1(G_Index + 2).text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                If IsNumeric(Text1(ptxSHIJI_QTY).text) Then
                    Text1(G_Index + 3).text = Format(CDbl(CLng(Text1(ptxSHIJI_QTY).text) * CDbl(Text1(G_Index + 2).text)), "#0.00")
                Else
                    Text1(G_Index + 3).text = ""
                End If
                            
                G_Index = G_Index + 5
            
            
            Case P_DOUKON   '同梱／構成
            
                d = d + 1
                D_Item_Tbl(d).SYUBETSU = StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode)
' 2013.01.07 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                D_Item_Tbl(d).JGYOBU = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
'                D_Item_Tbl(d).JGYOBU = BUZAI            '同梱/構成の事業部を「部材」固定に変更
' 2013.01.07 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                D_Item_Tbl(d).NAIGAI = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
                D_Item_Tbl(d).HIN_GAI = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                D_Item_Tbl(d).QTY = CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))
                D_Item_Tbl(d).BIKOU = StrConv(P_COMPO_K_REC.KO_BIKOU, vbUnicode)
                            
                If d > 5 Then
                Else
                            '種別
                    For i = 0 To Combo1(DC_Index).ListCount - 1
                    
                        If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = Right(Combo1(DC_Index).List(i), 2) Then
                            Combo1(DC_Index).ListIndex = i
                            Exit For
                        End If
                    
                    Next i
                                
                                
                                
                                '事業部 2016.01.27
                    Combo2(DC_Index).ListIndex = -1
                    For i = 0 To Combo2(DC_Index).ListCount - 1

                    
                        If D_Item_Tbl(d).JGYOBU = Right(Combo2(DC_Index).List(i), 1) Then
                            Combo2(DC_Index).ListIndex = i
                            Exit For
                        End If
                    
                    Next i
                                
                                
                                
                                
                    DC_Index = DC_Index + 1
                                
                                '品番
                    Text1(DT_Index).text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2016.01.27
'                    Select Case D_Item_Tbl(d).JGYOBU
'                        Case SHIZAI
'                            Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
'                            wkJgyobu = BUZAI
'                        Case SETSUBI
'                            Call UniCode_Conv(K0_ITEM.JGYOBU, YUKO_JGYOBU)
'                            wkJgyobu = YUKO_JGYOBU
'                        Case Else
'                            Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(d).JGYOBU)
'                            wkJgyobu = D_Item_Tbl(d).JGYOBU
'                    End Select

                     wkJgyobu = D_Item_Tbl(d).JGYOBU

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2016.01.27
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                    Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(d).JGYOBU)
                    Call UniCode_Conv(K0_ITEM.NAIGAI, D_Item_Tbl(d).NAIGAI)
                    
                    
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(DT_Index).text)
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                            '品名
                            Text1(DT_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                            '標準棚番
                            If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                                Text1(DT_Index + 4) = ""
                            Else
                                Text1(DT_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                            End If
                        
                        
'--------------------------------------------------- 大阪  部材対応　2012.03.18
                            
                            '>>>>>>>>>>>>>>>>   2013.01.07  DEL
                            'If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI Then
                            '    wkJgyobu = BUZAI
                            'Else
                            '    wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)  '2012.04.04
                            '    'wkJgyobu = YUKO_JGYOBU                          '2012.04.04
                            'End If
                            '>>>>>>>>>>>>>>>>   2013.01.07  DEL
                            
                            '標準棚番を設定 2013.01.11
                            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, wkJgyobu, _
                                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                                                    (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)), _
                                                                    , , Jyogai_Soko_umu) Then
                                Exit Function
                            
                            End If
'--------------------------------------------------- 大阪  部材対応　2012.03.18
                        
                            Text1(DT_Index + 5) = Format(Sumi_Qty + Mi_Qty, "#0")
                        
                        
                        Case BtErrKeyNotFound
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2016.01.27 読み直し廃止
                            
'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21
'                            Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
'                            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
'                            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
'
'                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                            Select Case sts
'                                Case BtNoErr
'                                    '品名
'                                    Text1(DT_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
'                                    '標準棚番
'                                    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
'                                        Text1(DT_Index + 4) = ""
'                                    Else
'                                        Text1(DT_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
'                                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
'                                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
'                                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
'                                    End If
'
'
'                                    '標準棚番分を設定   2013.01.11
'                                    If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
'                                                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
'                                                                            StrConv(ITEMREC.HIN_GAI, vbUnicode), _
'                                                                            (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode))) Then
'                                        Exit Function
'
'                                    End If
'
'                                    Text1(DT_Index + 5) = Format(Sumi_Qty + Mi_Qty, "#0")
'
'                                Case BtErrKeyNotFound
'
'
'                                    Text1(DT_Index + 1) = "未登録品番"
'                                    Text1(DT_Index + 4) = ""
'                                    Text1(DT_Index + 5) = ""
'                                Case Else
'                                    Call Input_UnLock             '2008.01.15
'                                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
'                                    Exit Function
'
'                            End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2016.01.27
                                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "")
                            Exit Function
                    
                    End Select
                
                
                    Text1(DT_Index + 2).text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                    If IsNumeric(Text1(ptxSHIJI_QTY).text) Then
                        Text1(DT_Index + 3).text = Format(CDbl(CLng(Text1(ptxSHIJI_QTY).text) * CDbl(Text1(DT_Index + 2).text)), "#0.00")
                    Else
                        Text1(DT_Index + 3).text = ""
                    End If
                    Text1(DT_Index + 6).text = Trim(StrConv(P_COMPO_K_REC.KO_BIKOU, vbUnicode))
                                
                    DT_Index = DT_Index + 7
                End If
        
        End Select
        
        
        
        
        com = BtOpGetNext
    
    Loop

        
    
    
    P_COMPO_Disp_Proc = False

End Function

Private Function Tbl_To_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   ﾃｰﾌﾞﾙより同梱／構成の表示
'----------------------------------------------------------------------------
Dim sts         As Integer


Dim i           As Integer
Dim j           As Integer

Dim DC_Index    As Integer
Dim DT_Index    As Integer

Dim Sumi_Qty    As Long
Dim Mi_Qty      As Long

Dim wkJgyobu    As String * 1

    Tbl_To_Disp_Proc = True

    
    DT_Index = ptxD_HIN_GAI01
    DC_Index = pcmbD_SYUBETSU01
                
                
    For i = 0 To 5          '最初の６行を表示
                    
                    '種別
        Combo1(DC_Index).ListIndex = -1
        For j = 0 To Combo1(DC_Index).ListCount - 1
        
            If D_Item_Tbl(i).SYUBETSU = Right(Combo1(DC_Index).List(j), 2) Then
                Combo1(DC_Index).ListIndex = j
                Exit For
            End If
        
        Next j
    
    
                    '事業部
'        Combo1(DC_Index).ListIndex = -1            '2017.10.17
        Combo2(DC_Index).ListIndex = -1             '2017.10.17
        For j = 0 To Combo2(DC_Index).ListCount - 1
        
            If D_Item_Tbl(i).JGYOBU = Right(Combo2(DC_Index).List(j), 1) Then
                Combo2(DC_Index).ListIndex = j
                Exit For
            End If
        
        Next j
    
    
        DC_Index = DC_Index + 1
    
    
        If Trim(D_Item_Tbl(i).HIN_GAI) = "" Then
            Text1(DT_Index).text = ""
            Text1(DT_Index + 1).text = ""
            Text1(DT_Index + 2).text = ""
            Text1(DT_Index + 3).text = ""
            Text1(DT_Index + 4).text = ""
            Text1(DT_Index + 5).text = ""
            Text1(DT_Index + 6).text = ""
        Else
        
            Text1(DT_Index).text = D_Item_Tbl(i).HIN_GAI    '品番
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2016.01.27
'            Select Case D_Item_Tbl(i).JGYOBU
'                Case SHIZAI
'                    Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)
'                Case SETSUBI
'                    Call UniCode_Conv(K0_ITEM.JGYOBU, YUKO_JGYOBU)
'                Case Else
'                    Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(i).JGYOBU)
'            End Select

            Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(i).JGYOBU)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2016.01.27

'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
            Call UniCode_Conv(K0_ITEM.NAIGAI, D_Item_Tbl(i).NAIGAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(DT_Index).text)
                    
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    '品名
                    Text1(DT_Index + 1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                    '標準棚番
                    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                        Text1(DT_Index + 4) = ""
                    Else
                        Text1(DT_Index + 4) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                    End If
                        
                
'--------------------------------------------------- 大阪  部材対応　2012.03.18
                    
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                    'If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI Then
                    '    wkJgyobu = BUZAI
                    'Else
                    '    'wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)  '2012.04.04
                    '    wkJgyobu = YUKO_JGYOBU                          '2012.04.04
                    'End If

'>>>>>>>>>  2016.01.27
'                    Select Case StrConv(ITEMREC.JGYOBU, vbUnicode)
'                        Case SHIZAI
'                            wkJgyobu = SHIZAI
'                        Case SETSUBI
'                            wkJgyobu = YUKO_JGYOBU
'                        Case Else
'                            wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
'                    End Select



'                   Select Case StrConv(ITEMREC.JGYOBU, vbUnicode)
'                       Case SHIZAI
'                            wkJgyobu = BUZAI
'                        Case SETSUBI
'                            wkJgyobu = YUKO_JGYOBU
'                        Case Else
'                            wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
'                    End Select

                    wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
'>>>>>>>>>  2016.01.27

'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07 引数に標準棚番を追加
                    If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, wkJgyobu, _
                                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                            StrConv(ITEMREC.HIN_GAI, vbUnicode), StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                                                StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                                                StrConv(ITEMREC.ST_DAN, vbUnicode), , , Jyogai_Soko_umu) Then
                        Exit Function
                    
                    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.01.07 引数に標準棚番を追加
'--------------------------------------------------- 大阪  部材対応　2012.03.18
                    '在庫数
                    Text1(DT_Index + 5) = Format(Sumi_Qty + Mi_Qty, "#0")
                        
                Case BtErrKeyNotFound
                    
                    Text1(DT_Index + 1) = "未登録品番"
                    Text1(DT_Index + 4) = ""
                    Text1(DT_Index + 5) = ""
                                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            
            End Select
        
            '員数
            Text1(DT_Index + 2).text = Format(D_Item_Tbl(i).QTY, "#0.00")
            '数量
            Text1(DT_Index + 3).text = Format(D_Item_Tbl(i).SHIJI_QTY, "#0.00")
            '備考
            Text1(DT_Index + 6).text = D_Item_Tbl(i).BIKOU
        
        End If
    
        DT_Index = DT_Index + 7
    
    Next i
                
    Tbl_To_Disp_Proc = False


End Function

Private Function Y_SYUKA_Make_Proc(i As Integer) As Integer
'----------------------------------------------------------------------------
'                   出荷指示の作成
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer

Dim ID_NO   As String * 12
Dim DEN_NO  As String * 6


    Y_SYUKA_Make_Proc = True

    '品目マスタ読み込み（在庫有無の判定）
    Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(i).JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, D_Item_Tbl(i).NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, D_Item_Tbl(i).HIN_GAI)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            
            If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI Then
            
                If StrConv(ITEMREC.ZAIKO_F, vbUnicode) <> P_ZAIKO_F_ON Then
                    '在庫対象外品目
                    Y_SYUKA_Make_Proc = False
                    Exit Function
                End If
            End If
        Case BtErrKeyNotFound
            D_Item_Tbl(i).ID_NO = ""
            Y_SYUKA_Make_Proc = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    
    End Select

    '-------------------------------------------------- 出荷予定編集
    Call UniCode_Conv(Y_SYUREC.WEL_ID, "")                                  '使用子機ＩＤ
    Call UniCode_Conv(Y_SYUREC.PRG_ID, "")                                  '使用中プログラム
    Call UniCode_Conv(Y_SYUREC.KAN_KBN, "0")                                '完了区分
    Call UniCode_Conv(Y_SYUREC.DT_SYU, "R")                                 'データ種別
    Call UniCode_Conv(Y_SYUREC.JGYOBU, D_Item_Tbl(i).JGYOBU)                '事業部区分
    Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN)                        '注文区分
    Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN)

    If Den_No_Set_Proc(21, D_Item_Tbl(i).JGYOBU, ID_NO) Then                'IDNO
        Exit Function
    End If
    
    Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO)
    Call UniCode_Conv(Y_SYUREC.ID_NO, ID_NO)

    Call UniCode_Conv(Y_SYUREC.NAIGAI, D_Item_Tbl(i).NAIGAI)                '国内外
                                                                    
    Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, D_Item_Tbl(i).HIN_GAI)           '品目番号
    Call UniCode_Conv(Y_SYUREC.HIN_NO, D_Item_Tbl(i).HIN_GAI)               '品目番号

                                                                            '得意先コード
    Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, MTS_CODE)
    Call UniCode_Conv(Y_SYUREC.MUKE_CODE, MTS_CODE)
                                                                            '直送先コード
    Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, SS_CODE)
    Call UniCode_Conv(Y_SYUREC.SS_CODE, SS_CODE)
                                                                            '出荷日
    Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode))
    Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode))

    Call UniCode_Conv(Y_SYUREC.JGYOBA, "")                                  '事業場
    Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")                                'データ区分
    Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")                                '取引区分
                                                                            '伝票№
    If Den_No_Set_Proc(20, D_Item_Tbl(i).JGYOBU, DEN_NO) Then
        Exit Function
    End If
    Call UniCode_Conv(Y_SYUREC.DEN_NO, DEN_NO)
                                                                            '出庫数量
    Call UniCode_Conv(Y_SYUREC.SURYO, Format(Int(D_Item_Tbl(i).SHIJI_QTY + 0.9), "0000000"))
        
    Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, "")                             '出庫収支
    Call UniCode_Conv(Y_SYUREC.ODER_NO, "")                                 'オーダー番号
    Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")                                 'アイテム番号
    Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")                               'オーダー番号略号
                                                                            '得意先名称
    Call UniCode_Conv(Y_SYUREC.MUKE_NAME, StrConv(MTSREC.MUKE_NAME, vbUnicode))
                                                                            '注文区分名称
    Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_N)
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
    Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                            StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                            StrConv(ITEMREC.ST_DAN, vbUnicode))
    
    Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")                               '印刷日付
    Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")                                 '完了日付
    Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")                              '検品日付
    Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")                                 '特売り区分
    
    Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "00000000")                      '実績数量
                                                                            '更新日時
    Call UniCode_Conv(Y_SYUREC.INS_NOW, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
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
                                        '自動発番データ重複は再発行
                sts = Den_No_Set_Proc(21, D_Item_Tbl(i).JGYOBU, ID_NO)
                If sts Then
                    Exit Function
                End If

                Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO)
                Call UniCode_Conv(Y_SYUREC.ID_NO, ID_NO)
                
            Case Else
                Call File_Error(sts, BtOpInsert, "出荷予定データ")
                Exit Function
        End Select
    Loop
    
    
        
    
    If SYUKA_LOG_ON Then
        Call SYUKA_LOG_OUT_PROC("INS", "AFT")
    End If
    
    D_Item_Tbl(i).ID_NO = ID_NO
    
    Y_SYUKA_Make_Proc = False
    
End Function
Public Function wP_SSHIJI_O_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              商品化指図(親)ワーク  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    wP_SSHIJI_O_Open = True
                                            '商品化指図(親)ﾃﾞｰﾀフルパス取込み
    sts = GetIni("FILE", P_SSHIJI_O_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SSHIJI_O]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, wP_SSHIJI_O_POS, wP_SSHIJI_O_REC, Len(wP_SSHIJI_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "商品化指図(親)ﾜｰｸ")
                Exit Function
        End Select
    Loop
    
    wP_SSHIJI_O_Open = False

End Function


Private Function ORDER_Check_Proc(Order_QTY As Long, SHIJI_QTY As Long) As Integer

'----------------------------------------------------------------------------
'                   発注検討　親品番注文Ｆとのすり合わせ
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim com     As Integer
    
    
    ORDER_Check_Proc = True
    
    
    Order_QTY = 0
    SHIJI_QTY = Val(Text1(ptxSHIJI_QTY).text)
    
    
    
    
    
    
    '------------------------------------   発注検討　親品番注文Ｆ　チェック
    Call UniCode_Conv(K6_ODR_ORDER.ORDER_NO, Text1(ptxORDER_NO).text)
    
    
        
    sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K6_ODR_ORDER, Len(K6_ODR_ORDER), 6)
    Select Case sts
        Case BtNoErr
                    
            Order_QTY = Val(Text1(ptxSHIJI_QTY).text)
        
        Case BtErrKeyNotFound
        
        
        Case Else
            Call File_Error(sts, com, "発注検討：親品番注文Ｆ")
            Exit Function
    End Select
    
    
    '------------------------------------   商品化指図データ（親）　チェック
    Call UniCode_Conv(K5_P_SSHIJI_O.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(K5_P_SSHIJI_O.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K5_P_SSHIJI_O.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K5_P_SSHIJI_O.HIN_GAI, Text1(ptxHIN_GAI).text)
    Call UniCode_Conv(K5_P_SSHIJI_O.KAN_F, "")
    com = BtOpGetGreaterEqual
    Do
        
       DoEvents
        
        sts = BTRV(com, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K5_P_SSHIJI_O, Len(K5_P_SSHIJI_O), 5)
        Select Case sts
            Case BtNoErr
                If StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                    StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                    StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                    StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode) <> Text1(ptxHIN_GAI).text Then
                    
                    Exit Do
                
                End If
            
                If Trim(StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode)) <> "" Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "商品化指図データ（親）")
                Exit Function
        End Select
    
        If Trim(StrConv(ODR_ORDER_REC.BUN_NO, vbUnicode)) <> "" Then
        Else
            Order_QTY = Order_QTY + Val(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode))
        End If
        
        com = BtOpGetNext
    
    Loop
        
'    Order_QTY = Val(Text1(ptxSHIJI_QTY).text)
    
    
    
    
    
    ORDER_Check_Proc = False
End Function

Private Function Zaiko_Check_Proc(ZAIKO_F As Boolean) As Integer
'----------------------------------------------------------------------------
'                   現在庫とのすり合わせ
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim i               As Integer
Dim j               As Integer
 
Dim Sumi_Qty        As Long
Dim Mi_Qty          As Long
 
Dim wkJgyobu        As String * 1
 
 
    Zaiko_Check_Proc = True
    
       
    Erase ZAIKO_FUSOKU
    
    
    
    ZAIKO_F = False
    
    j = -1
    
'---------------------------------------------------------------------- 個装資材の使用集計
    For i = ptxK_HIN_GAI01 To ptxK_HIN_GAI05 Step 5
    
    
        If Trim(Text1(i).text) = "" Then
        Else
            
            
            If Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1) = SHIZAI Then
                wkJgyobu = BUZAI
            Else
                'wkJgyobu = Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1)   '2012.04.04
                wkJgyobu = YUKO_JGYOBU                                      '2012.04.04
            End If
                        
            Call UniCode_Conv(K0_ITEM.JGYOBU, wkJgyobu)
           
            
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(i).text)
            
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(ITEMREC.JGYOBU, wkJgyobu)
                    Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(ITEMREC.HIN_GAI, Text1(i).text)
            
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                    Exit Function
            
            End Select
    
    
            If j = -1 Then
                j = 0
                ReDim Preserve ZAIKO_FUSOKU(0 To j)
                ZAIKO_FUSOKU(j).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                ZAIKO_FUSOKU(j).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
                ZAIKO_FUSOKU(j).HIN_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                ZAIKO_FUSOKU(j).USE_QTY = Val(Text1(i + 3).text)
                ZAIKO_FUSOKU(j).ZAIKO_QTY = 0
            Else
                For j = 0 To UBound(ZAIKO_FUSOKU)
                
                    If ZAIKO_FUSOKU(j).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode) Or _
                        ZAIKO_FUSOKU(j).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode) Or _
                        RTrim(ZAIKO_FUSOKU(j).HIN_GAI) = RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                
                        Exit For
                    End If
                
                Next j
                
                If j > UBound(ZAIKO_FUSOKU) Then
                    ReDim Preserve ZAIKO_FUSOKU(0 To j)
                    ZAIKO_FUSOKU(j).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                    ZAIKO_FUSOKU(j).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
                    ZAIKO_FUSOKU(j).HIN_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                    ZAIKO_FUSOKU(j).USE_QTY = Val(Text1(i + 3).text)
                    ZAIKO_FUSOKU(j).ZAIKO_QTY = 0
                Else
                    ZAIKO_FUSOKU(j).USE_QTY = ZAIKO_FUSOKU(j).USE_QTY + Val(Text1(i + 3).text)
                End If
            End If
        End If
    Next i
'---------------------------------------------------------------------- 個装資材の使用集計
    
    
'---------------------------------------------------------------------- 外装資材の使用集計
    For i = ptxG_HIN_GAI01 To ptxK_HIN_GAI03 Step 5
    
    
        If Trim(Text1(i).text) = "" Then
        Else
    
            If Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1) = SHIZAI Then
                wkJgyobu = BUZAI
            Else
                'wkJgyobu = Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1)   '2012.04.04
                wkJgyobu = YUKO_JGYOBU                                      '2012.04.04
            End If
                        
            Call UniCode_Conv(K0_ITEM.JGYOBU, wkJgyobu)
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(i).text)
            
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(ITEMREC.JGYOBU, wkJgyobu)
                    Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(ITEMREC.HIN_GAI, Text1(i).text)
            
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                    Exit Function
            
            End Select
    
    
            If j = -1 Then
                j = 0
                ReDim Preserve ZAIKO_FUSOKU(0 To j)
                ZAIKO_FUSOKU(j).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                ZAIKO_FUSOKU(j).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
                ZAIKO_FUSOKU(j).HIN_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                ZAIKO_FUSOKU(j).USE_QTY = Val(Text1(i + 3).text)
                ZAIKO_FUSOKU(j).ZAIKO_QTY = 0
            
            
            Else
                For j = 0 To UBound(ZAIKO_FUSOKU)
                
                    If ZAIKO_FUSOKU(j).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode) Or _
                        ZAIKO_FUSOKU(j).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode) Or _
                        RTrim(ZAIKO_FUSOKU(j).HIN_GAI) = RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                
                        Exit For
                
                    End If
                Next j
                
                If j > UBound(ZAIKO_FUSOKU) Then
                    ReDim Preserve ZAIKO_FUSOKU(0 To j)
                    ZAIKO_FUSOKU(j).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                    ZAIKO_FUSOKU(j).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
                    ZAIKO_FUSOKU(j).HIN_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                    ZAIKO_FUSOKU(j).USE_QTY = Val(Text1(i + 3).text)
                    ZAIKO_FUSOKU(j).ZAIKO_QTY = 0
                Else
                    ZAIKO_FUSOKU(j).USE_QTY = ZAIKO_FUSOKU(j).USE_QTY + Val(Text1(i + 3).text)
                End If
            End If
        End If
    Next i
'---------------------------------------------------------------------- 外装資材の使用集計
    
    
    
'---------------------------------------------------------------------- 同梱分の使用集計
    For i = 0 To UBound(D_Item_Tbl)
        If Trim(D_Item_Tbl(i).HIN_GAI) = "" Then
        Else
            '>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
            'If D_Item_Tbl(i).JGYOBU = SHIZAI Then
            '    wkJgyobu = BUZAI
            'Else
            '    'wkJgyobu = D_Item_Tbl(i).JGYOBU        '2012.04.04
            '    wkJgyobu = YUKO_JGYOBU                  '2012.04.04
            'End If
            
            
            Select Case D_Item_Tbl(i).JGYOBU
'2016.05.31                Case SHIZAI
'2016.05.31                    wkJgyobu = BUZAI
'2016.05.31                Case SETSUBI
'2016.05.31                    wkJgyobu = YUKO_JGYOBU
                Case Else
                    wkJgyobu = D_Item_Tbl(i).JGYOBU
            End Select
            '>>>>>>>>>>>>>>>>>>>>>> 2013.01.07
                        
            Call UniCode_Conv(K0_ITEM.JGYOBU, wkJgyobu)
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, D_Item_Tbl(i).HIN_GAI)
            
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(ITEMREC.JGYOBU, wkJgyobu)
                    Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(ITEMREC.HIN_GAI, D_Item_Tbl(i).HIN_GAI)
            
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                    Exit Function
            
            End Select
    
    
            If j = -1 Then
                j = 0
                ReDim Preserve ZAIKO_FUSOKU(0 To j)
                ZAIKO_FUSOKU(j).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                ZAIKO_FUSOKU(j).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
                ZAIKO_FUSOKU(j).HIN_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                ZAIKO_FUSOKU(j).USE_QTY = D_Item_Tbl(i).SHIJI_QTY
                ZAIKO_FUSOKU(j).ZAIKO_QTY = 0
                
            
            Else
                For j = 0 To UBound(ZAIKO_FUSOKU)
                
                    If ZAIKO_FUSOKU(j).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode) And _
                        ZAIKO_FUSOKU(j).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode) And _
                        RTrim(ZAIKO_FUSOKU(j).HIN_GAI) = RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                
                        Exit For
                
                    End If
                
                Next j
                
                
                If j > UBound(ZAIKO_FUSOKU) Then
                    ReDim Preserve ZAIKO_FUSOKU(0 To j)
                    ZAIKO_FUSOKU(j).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                    ZAIKO_FUSOKU(j).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
                    ZAIKO_FUSOKU(j).HIN_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                    ZAIKO_FUSOKU(j).USE_QTY = D_Item_Tbl(i).SHIJI_QTY
                Else
                    ZAIKO_FUSOKU(j).USE_QTY = ZAIKO_FUSOKU(j).USE_QTY + D_Item_Tbl(i).SHIJI_QTY
                End If
            End If
        End If
    
    
    
    Next i
'---------------------------------------------------------------------- 同梱分の使用集計
    
    If j = -1 Then
        Zaiko_Check_Proc = False
        Exit Function
    End If
    
'---------------------------------------------------------------------- 現在庫の集計
    For i = 0 To UBound(ZAIKO_FUSOKU)
        
'        If ZAIKO_FUSOKU(i).JGYOBU = SHIZAI Then
'            wkJgyobu = BUZAI
'        Else
'            'wkJgyobu = ZAIKO_FUSOKU(i).JGYOBU  '2012.04.04
'            wkJgyobu = YUKO_JGYOBU              '2012.04.04
'        End If
        
        
If ZAIKO_FUSOKU(i).JGYOBU = "B" Then
Debug.Print
End If
        
        
        
        If Zaiko_Syukei_Proc(Sumi_Qty, _
                                Mi_Qty, _
                                ZAIKO_FUSOKU(i).JGYOBU, _
                                ZAIKO_FUSOKU(i).NAIGAI, _
                                ZAIKO_FUSOKU(i).HIN_GAI, , , , Jyogai_Soko_umu) Then
            Exit Function
        End If
        ZAIKO_FUSOKU(i).ZAIKO_QTY = Sumi_Qty + Mi_Qty
    Next i
'---------------------------------------------------------------------- 現在庫の集計
    
'---------------------------------------------------------------------- 使用予約分の集計
    For i = 0 To UBound(ZAIKO_FUSOKU)
        
'        If ZAIKO_FUSOKU(i).JGYOBU = SHIZAI Then
'            wkJgyobu = BUZAI
'        Else
'            'wkJgyobu = ZAIKO_FUSOKU(i).JGYOBU  '2012.04.04
'            wkJgyobu = YUKO_JGYOBU              '2012.04.04
'        End If
        
        
        Call UniCode_Conv(K2_P_SSHIJI_K.KO_JGYOBU, ZAIKO_FUSOKU(i).JGYOBU)
        Call UniCode_Conv(K2_P_SSHIJI_K.KO_NAIGAI, ZAIKO_FUSOKU(i).NAIGAI)
        Call UniCode_Conv(K2_P_SSHIJI_K.KO_HIN_GAI, ZAIKO_FUSOKU(i).HIN_GAI)
        Call UniCode_Conv(K2_P_SSHIJI_K.IDO_SUMI, "")
    
        com = BtOpGetGreaterEqual
    
        Do
            DoEvents
            
            sts = BTRV(com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K2_P_SSHIJI_K, Len(K2_P_SSHIJI_K), 2)
            Select Case sts
                Case BtNoErr
                    
                    If StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) <> ZAIKO_FUSOKU(i).JGYOBU Or _
                        StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode) <> ZAIKO_FUSOKU(i).NAIGAI Or _
                        StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode) <> ZAIKO_FUSOKU(i).HIN_GAI Then
                        Exit Do
                    End If
                            


                Case BtErrEOF
                    Exit Do


                Case Else
                    Call File_Error(sts, com, "商品化指図データ（子）")
                    Exit Function


            
            End Select
        
        
            If StrConv(P_SSHIJI_K_REC.CALCEL_F, vbUnicode) <> P_CANCEL_ON Then
            
                ZAIKO_FUSOKU(i).ZAIKO_QTY = ZAIKO_FUSOKU(i).ZAIKO_QTY - Val(StrConv(P_SSHIJI_K_REC.HIKIATE_QTY, vbUnicode))
            
            End If
        
        
            com = BtOpGetNext
        
        Loop
    Next i
'---------------------------------------------------------------------- 使用予約分の集計
    
'---------------------------------------------------------------------- 差異数の集計
    For i = 0 To UBound(ZAIKO_FUSOKU)
        ZAIKO_FUSOKU(i).SAI_QTY = ZAIKO_FUSOKU(i).ZAIKO_QTY - ZAIKO_FUSOKU(i).USE_QTY
        If ZAIKO_FUSOKU(i).SAI_QTY < 0 Then
            ZAIKO_F = True
        End If
    Next i
'---------------------------------------------------------------------- 差異数の集計
    
    Zaiko_Check_Proc = False

End Function

Private Sub Text1_LostFocus(Index As Integer)



'--------------------------------------------------- 大阪  部材対応　2012.03.09
    Select Case Index
        
        
        
        Case ptxORDER_NO                                        '2016.06.22
        
            If Error_Check_Proc(Index, 0, 0) Then   'エラーチェック
                Exit Sub
            End If
        
        
        
        Case ptxHIN_GAI
            Text1(Index).text = StrConv(Text1(Index).text, vbUpperCase)
    
            '>>>>>>>>>>>>>> 2013.12.28
'            If svHin_Gai <> Text1(Index).text Then             '2016.01.27
            If Trim(svHin_Gai) <> Text1(Index).text Then        '2016.01.27
            
                        
                If Error_Check_Proc(Index, 0, 0) Then   'エラーチェック
                    Exit Sub
                End If
            
            End If
            '>>>>>>>>>>>>>> 2013.12.28
    
    End Select
'--------------------------------------------------- 大阪  部材対応　2012.03.09




End Sub

Private Sub Disp_Lock_Proc(Mode As Integer)
'----------------------------------------------------------------------------
'                   画面Lock/UnLock
'           2012.03.18
'----------------------------------------------------------------------------
Dim i   As Integer


    Text1(ptxHIN_GAI).Locked = Mode
    Text1(ptxSHIJI_QTY).Locked = Mode
    
    For i = ptxK_HIN_GAI01 To ptxD_BIKOU06
'        Text1(ptxSHIJI_QTY).Locked = Mode      '2016.05.18
        Text1(i).Locked = Mode                  '2016.05.18
    Next i


    For i = pcmbD_SYUBETSU01 To pcmbD_SYUBETSU06
        Combo1(i).Locked = Mode
    Next i

        
    For i = pcmbD_JGYOBU01 To pcmbD_JGYOBU06    '2016.05.18
        Combo2(i).Locked = Mode                 '2016.05.18
    Next i                                      '2016.05.18
        
        
        
    For i = 0 To 24
        PI000152.Combo1(i).Locked = Mode
    Next i
    
    
    For i = 0 To 24                             '2016.05.18
        PI000152.Combo2(i).Locked = Mode        '2016.05.18
    Next i                                      '2016.05.18
    
    
    For i = 0 To 174
        PI000152.Text1(i).Locked = Mode
    Next i




End Sub


Private Sub DEF_JGYOBU_PROC()
'-------------------------------------------------------------------------
'
'   読み替えデフォルト事業部　セット
'
'       2014.03.24
'
'
'-------------------------------------------------------------------------
Dim i   As Integer
Dim j   As Integer
    
Dim c   As String * 128
    
    
    i = 0
    j = 0
    Do
        If GetIni("JIGYOBU", "code" & RTrim(Format$(i + 1, "#0")), "SYS", c) Then
            Exit Sub
        End If
        If RTrim(c) = "0" Then
            Exit Do
        End If

        ReDim Preserve YOMI_JGYOBU(j)

        YOMI_JGYOBU(j) = RTrim(c)
        j = j + 1
        i = i + 1
    Loop

End Sub
