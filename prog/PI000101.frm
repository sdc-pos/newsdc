VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form PI000101 
   Caption         =   "商品化指図票発行 "
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ClipControls    =   0   'False
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
   ScaleHeight     =   10155
   ScaleWidth      =   15240
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Command1 
      Caption         =   "キャンセル"
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
      Index           =   12
      Left            =   8040
      TabIndex        =   193
      TabStop         =   0   'False
      Top             =   9360
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "紙"
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   27
      Top             =   3720
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "プラ"
      Height          =   255
      Index           =   6
      Left            =   2880
      TabIndex        =   28
      Top             =   3720
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "適用機種ラベル"
      Height          =   255
      Index           =   7
      Left            =   3720
      TabIndex        =   29
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox lstGensankoku 
      Height          =   780
      Left            =   12735
      Sorted          =   -1  'True
      TabIndex        =   181
      Top             =   2220
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.TextBox txGensankoku 
      Height          =   375
      Left            =   12180
      TabIndex        =   177
      Top             =   240
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   98
      Left            =   13230
      MaxLength       =   8
      TabIndex        =   119
      Top             =   9000
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   11
      Left            =   14160
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   10
      Left            =   13200
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1080
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   1800
      TabIndex        =   173
      Top             =   2160
      Width           =   4455
      Begin VB.OptionButton Option1 
         Caption         =   "欠品解除"
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "事前"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   175
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "スポット"
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   97
      Left            =   10200
      TabIndex        =   118
      Top             =   8520
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   96
      Left            =   9600
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   117
      TabStop         =   0   'False
      Top             =   8520
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   95
      Left            =   8160
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   116
      TabStop         =   0   'False
      Top             =   8520
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   94
      Left            =   7080
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   115
      TabStop         =   0   'False
      Top             =   8520
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   93
      Left            =   6240
      MaxLength       =   6
      TabIndex        =   114
      Top             =   8520
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   92
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   113
      TabStop         =   0   'False
      Top             =   8520
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   91
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   112
      Top             =   8520
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   90
      Left            =   10200
      TabIndex        =   110
      Top             =   8160
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   89
      Left            =   9600
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   109
      TabStop         =   0   'False
      Top             =   8160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   88
      Left            =   8160
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   87
      Left            =   7080
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   86
      Left            =   6240
      MaxLength       =   6
      TabIndex        =   106
      Top             =   8160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   85
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   8160
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   84
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   104
      Top             =   8160
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   83
      Left            =   10200
      TabIndex        =   102
      Top             =   7800
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   82
      Left            =   9600
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   7800
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   81
      Left            =   8160
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   80
      Left            =   7080
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   79
      Left            =   6240
      MaxLength       =   6
      TabIndex        =   98
      Top             =   7800
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   78
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   7800
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   77
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   96
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   76
      Left            =   10200
      TabIndex        =   94
      Top             =   7440
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   75
      Left            =   9600
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   74
      Left            =   8160
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   73
      Left            =   7080
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   72
      Left            =   6240
      MaxLength       =   6
      TabIndex        =   90
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   71
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   7440
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   70
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   88
      Top             =   7440
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   69
      Left            =   10200
      TabIndex        =   86
      Top             =   7080
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   68
      Left            =   9600
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   7080
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   67
      Left            =   8160
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   66
      Left            =   7080
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   65
      Left            =   6240
      MaxLength       =   6
      TabIndex        =   82
      Top             =   7080
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   64
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   7080
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   63
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   80
      Top             =   7080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   55
      Left            =   13320
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   54
      Left            =   12240
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   53
      Left            =   11400
      MaxLength       =   6
      TabIndex        =   68
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   52
      Left            =   9240
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   5280
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   51
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   66
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   50
      Left            =   13320
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   49
      Left            =   12240
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   48
      Left            =   11400
      MaxLength       =   6
      TabIndex        =   63
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   47
      Left            =   9240
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   4920
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   46
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   61
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   45
      Left            =   13320
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   44
      Left            =   12240
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   43
      Left            =   11400
      MaxLength       =   6
      TabIndex        =   58
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   42
      Left            =   9240
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   41
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   56
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   40
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   39
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   38
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   53
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   37
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   6000
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   36
      Left            =   600
      MaxLength       =   20
      TabIndex        =   51
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   35
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   34
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   33
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   48
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   32
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   31
      Left            =   600
      MaxLength       =   20
      TabIndex        =   46
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   30
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   29
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   28
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   43
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   27
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5280
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   26
      Left            =   600
      MaxLength       =   20
      TabIndex        =   41
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   25
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   24
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   23
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   38
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   15
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   18
      Top             =   1800
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1095
      Index           =   0
      Left            =   8040
      TabIndex        =   30
      Top             =   2640
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1931
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"PI000101.frx":0000
   End
   Begin VB.ComboBox Combo1 
      Height          =   336
      Index           =   8
      Left            =   240
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   111
      Top             =   8520
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   336
      Index           =   7
      Left            =   240
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   103
      Top             =   8160
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   336
      Index           =   6
      Left            =   240
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   95
      Top             =   7800
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   336
      Index           =   5
      Left            =   240
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   87
      Top             =   7440
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   336
      Index           =   4
      Left            =   240
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   79
      Top             =   7080
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   336
      Index           =   3
      Left            =   240
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   71
      Top             =   6720
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "出力対象"
      Height          =   732
      Left            =   240
      TabIndex        =   155
      Top             =   2880
      Width           =   6630
      Begin VB.ComboBox Combo2 
         Height          =   336
         Index           =   0
         Left            =   1440
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   184
         Top             =   240
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "機種ラベル"
         Height          =   375
         Index           =   4
         Left            =   6240
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.CheckBox Check1 
         Caption         =   "外装ラベル"
         Height          =   375
         Index           =   3
         Left            =   4320
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "パーツラベル"
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "指図票"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   22
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4920
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   21
      Left            =   600
      MaxLength       =   20
      TabIndex        =   36
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   20
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   19
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   18
      Left            =   4080
      MaxLength       =   6
      TabIndex        =   33
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   17
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   16
      Left            =   600
      MaxLength       =   20
      TabIndex        =   31
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   5
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   12
      Left            =   240
      MaxLength       =   5
      TabIndex        =   14
      Top             =   1800
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "見本作成"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   2
      Left            =   9840
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   19
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   14
      Left            =   5640
      MaxLength       =   10
      TabIndex        =   17
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   13
      Left            =   4200
      MaxLength       =   10
      TabIndex        =   16
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   336
      Index           =   1
      Left            =   960
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   15
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   9
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   9960
      MaxLength       =   8
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   7
      Left            =   5640
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   7
      Top             =   1080
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   240
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   6
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   6240
      MaxLength       =   5
      TabIndex        =   4
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   3
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   3340
      MaxLength       =   5
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   1
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
      Width           =   1050
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
      TabIndex        =   131
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
      TabIndex        =   130
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ｺﾋﾟｰして登録"
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
      Index           =   9
      Left            =   8040
      TabIndex        =   129
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1575
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
      Left            =   7200
      TabIndex        =   128
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
      Left            =   6240
      TabIndex        =   127
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
      Left            =   5400
      TabIndex        =   126
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
      Left            =   4560
      TabIndex        =   125
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
      Left            =   3720
      TabIndex        =   124
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
      TabIndex        =   123
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
      TabIndex        =   122
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
      TabIndex        =   121
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
      TabIndex        =   120
      Top             =   9000
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   62
      Left            =   10200
      TabIndex        =   78
      Top             =   6720
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   61
      Left            =   9600
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   6720
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   60
      Left            =   8160
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   59
      Left            =   7080
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   58
      Left            =   6240
      MaxLength       =   6
      TabIndex        =   74
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   57
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   6720
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   56
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   72
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label lblL_URIKIN3 
      Height          =   135
      Left            =   9840
      TabIndex        =   192
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblL_URIKIN2 
      Height          =   375
      Left            =   9240
      TabIndex        =   191
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblGAI_BUHIN 
      Height          =   135
      Left            =   8760
      TabIndex        =   190
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblL_Hin_Name_E 
      Height          =   255
      Left            =   8760
      TabIndex        =   189
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblL_JGYOBU_N 
      Height          =   315
      Left            =   11280
      TabIndex        =   188
      Top             =   9600
      Width           =   690
   End
   Begin VB.Label lblL_KAISHA_N 
      Height          =   315
      Left            =   9960
      TabIndex        =   187
      Top             =   9600
      Width           =   690
   End
   Begin VB.Label lblKISHU2 
      Height          =   255
      Left            =   11280
      TabIndex        =   186
      Top             =   5880
      Width           =   3375
   End
   Begin VB.Label lblKISHU1 
      Height          =   255
      Left            =   7560
      TabIndex        =   185
      Top             =   5880
      Width           =   3375
   End
   Begin VB.Label lblL_JGYOBU 
      Height          =   315
      Left            =   14490
      TabIndex        =   183
      Top             =   3300
      Width           =   690
   End
   Begin VB.Label lblL_KAISHA 
      Height          =   315
      Left            =   14490
      TabIndex        =   182
      Top             =   2880
      Width           =   690
   End
   Begin VB.Label lblGensankoku 
      BorderStyle     =   1  '実線
      Height          =   375
      Index           =   1
      Left            =   12735
      TabIndex        =   180
      Top             =   1800
      Width           =   2445
   End
   Begin VB.Label lblGensankoku 
      Height          =   255
      Index           =   0
      Left            =   13590
      TabIndex        =   179
      Top             =   1560
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "原産国"
      Height          =   255
      Index           =   25
      Left            =   12780
      TabIndex        =   178
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "ラベル発行枚数"
      Height          =   255
      Index           =   24
      Left            =   11445
      TabIndex        =   176
      Top             =   9120
      Width           =   1710
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "商品化済"
      Height          =   255
      Index           =   23
      Left            =   14040
      TabIndex        =   13
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "未商品"
      Height          =   255
      Index           =   17
      Left            =   13200
      TabIndex        =   174
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "棚番"
      Height          =   252
      Index           =   16
      Left            =   8520
      TabIndex        =   172
      Top             =   6480
      Width           =   492
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "数量"
      Height          =   252
      Index           =   22
      Left            =   7560
      TabIndex        =   171
      Top             =   6480
      Width           =   492
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "品名"
      Height          =   252
      Index           =   21
      Left            =   3720
      TabIndex        =   170
      Top             =   6480
      Width           =   492
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "内職ｸﾗｽ"
      Height          =   255
      Index           =   20
      Left            =   7320
      TabIndex        =   169
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "備考"
      Height          =   252
      Index           =   19
      Left            =   10920
      TabIndex        =   168
      Top             =   6480
      Width           =   492
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "在庫"
      Height          =   252
      Index           =   18
      Left            =   9720
      TabIndex        =   167
      Top             =   6480
      Width           =   492
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "員数"
      Height          =   252
      Index           =   15
      Left            =   6480
      TabIndex        =   166
      Top             =   6480
      Width           =   492
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "品番"
      Height          =   252
      Index           =   14
      Left            =   1440
      TabIndex        =   165
      Top             =   6480
      Width           =   492
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "種別"
      Height          =   252
      Index           =   12
      Left            =   240
      TabIndex        =   164
      Top             =   6480
      Width           =   492
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "外装資材№"
      Height          =   372
      Index           =   17
      Left            =   7560
      TabIndex        =   163
      Top             =   4200
      Width           =   1692
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "①"
      Height          =   372
      Index           =   16
      Left            =   7560
      TabIndex        =   162
      Top             =   4560
      Width           =   372
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "②"
      Height          =   372
      Index           =   15
      Left            =   7560
      TabIndex        =   161
      Top             =   4920
      Width           =   372
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "③"
      Height          =   372
      Index           =   14
      Left            =   7560
      TabIndex        =   160
      Top             =   5280
      Width           =   372
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "品名"
      Height          =   372
      Index           =   13
      Left            =   9240
      TabIndex        =   159
      Top             =   4200
      Width           =   2172
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "入数"
      Height          =   372
      Index           =   12
      Left            =   11400
      TabIndex        =   158
      Top             =   4200
      Width           =   852
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "数量"
      Height          =   372
      Index           =   11
      Left            =   12240
      TabIndex        =   157
      Top             =   4200
      Width           =   1092
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "棚番"
      Height          =   372
      Index           =   10
      Left            =   13320
      TabIndex        =   156
      Top             =   4200
      Width           =   1452
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "棚番"
      Height          =   372
      Index           =   9
      Left            =   6000
      TabIndex        =   154
      Top             =   4200
      Width           =   1452
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "数量"
      Height          =   372
      Index           =   8
      Left            =   4920
      TabIndex        =   153
      Top             =   4200
      Width           =   1092
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "員数"
      Height          =   372
      Index           =   7
      Left            =   4080
      TabIndex        =   152
      Top             =   4200
      Width           =   852
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "品名"
      Height          =   372
      Index           =   6
      Left            =   1920
      TabIndex        =   151
      Top             =   4200
      Width           =   2172
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "⑤"
      Height          =   372
      Index           =   5
      Left            =   240
      TabIndex        =   150
      Top             =   6000
      Width           =   372
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "④"
      Height          =   372
      Index           =   4
      Left            =   240
      TabIndex        =   149
      Top             =   5640
      Width           =   372
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "③"
      Height          =   372
      Index           =   3
      Left            =   240
      TabIndex        =   148
      Top             =   5280
      Width           =   372
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "②"
      Height          =   372
      Index           =   2
      Left            =   240
      TabIndex        =   147
      Top             =   4920
      Width           =   372
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "①"
      Height          =   372
      Index           =   1
      Left            =   240
      TabIndex        =   146
      Top             =   4560
      Width           =   372
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "個装資材№"
      Height          =   372
      Index           =   0
      Left            =   240
      TabIndex        =   145
      Top             =   4200
      Width           =   1692
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "備考"
      Height          =   255
      Index           =   13
      Left            =   8040
      TabIndex        =   144
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "収単/担当者"
      Height          =   255
      Index           =   11
      Left            =   9960
      TabIndex        =   143
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "付加ｸﾗｽ"
      Height          =   255
      Index           =   10
      Left            =   5880
      TabIndex        =   142
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
      TabIndex        =   141
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
      TabIndex        =   140
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "標準棚番"
      Height          =   255
      Index           =   7
      Left            =   11520
      TabIndex        =   139
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
      TabIndex        =   138
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
      TabIndex        =   137
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
      TabIndex        =   136
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "承認"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   6240
      TabIndex        =   135
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "担当者"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   134
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "発行日"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   133
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "指図票№"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   132
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "PI000101"
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


Private TEHAI       As String


'テキスト用添字
Private Const ptxSHIJI_NO% = 0              '指図票№
Private Const ptxHAKKO_DT% = 1              '発行日
Private Const ptxTANTO_CODE% = 2            '担当者ｺｰﾄﾞ
Private Const ptxTANTO_NAME% = 3            '担当者名称
Private Const ptxSHONIN_CODE% = 4           '承認者ｺｰﾄﾞ
Private Const ptxSHONIN_NAME% = 5           '承認者名称
Private Const ptxHIN_GAI% = 6               '品番
Private Const ptxHIN_NAME% = 7              '品名
Private Const ptxSHIJI_QTY% = 8             '数量
Private Const ptxST_LOCATION% = 9           '標準棚番
Private Const ptxMI_QTY% = 10               '未商品
Private Const ptxSUMI_QTY% = 11             '商品化済
Private Const ptxUKEHARAI_CODE% = 12        '手配先ｺｰﾄﾞ
Private Const ptxS_CLASS_CODE% = 13         '商品化ｸﾗｽ
Private Const ptxF_CLASS_CODE% = 14         '付加ｸﾗｽ
Private Const ptxN_CLASS_CODE% = 15         '内職ｸﾗｽ


Private Const ptxK_HIN_GAI01% = 16          '①　個装資材№
Private Const ptxK_HIN_NAME01% = 17         '①　個装資材名称
Private Const ptxK_QTY01% = 18              '①　員数
Private Const ptxK_SHIJI_QTY01% = 19        '①　数量
Private Const ptxK_ST_LOCATION01% = 20      '①　棚番

Private Const ptxK_HIN_GAI02% = 21          '②　個装資材№
Private Const ptxK_HIN_NAME02% = 22         '②　個装資材名称
Private Const ptxK_QTY02% = 23              '②　員数
Private Const ptxK_SHIJI_QTY02% = 24        '②　数量
Private Const ptxK_ST_LOCATION02% = 25      '②　棚番

Private Const ptxK_HIN_GAI03% = 26          '③　個装資材№
Private Const ptxK_HIN_NAME03% = 27         '③　個装資材名称
Private Const ptxK_QTY03% = 28              '③　員数
Private Const ptxK_SHIJI_QTY03% = 29        '③　数量
Private Const ptxK_ST_LOCATION03% = 30      '③
Private Const ptxK_HIN_GAI04% = 31          '④　個装資材№
Private Const ptxK_HIN_NAME04% = 32         '④　個装資材名称
Private Const ptxK_QTY04% = 33              '④　員数
Private Const ptxK_SHIJI_QTY04% = 34        '④　数量
Private Const ptxK_ST_LOCATION04% = 35      '④　棚番

Private Const ptxK_HIN_GAI05% = 36          '⑤　個装資材№
Private Const ptxK_HIN_NAME05% = 37         '⑤　個装資材名称
Private Const ptxK_QTY05% = 38              '⑤　員数
Private Const ptxK_SHIJI_QTY05% = 39        '⑤　数量
Private Const ptxK_ST_LOCATION05% = 40      '⑤　棚番


Private Const ptxG_HIN_GAI01% = 41          '①　外装資材№
Private Const ptxG_HIN_NAME01% = 42         '①　外装資材名称
Private Const ptxG_QTY01% = 43              '①　員数
Private Const ptxG_SHIJI_QTY01% = 44        '①　数量
Private Const ptxG_ST_LOCATION01% = 45      '①　棚番

Private Const ptxG_HIN_GAI02% = 46          '②　外装資材№
Private Const ptxG_HIN_NAME02% = 47         '②　外装資材名称
Private Const ptxG_QTY02% = 48              '②　員数
Private Const ptxG_SHIJI_QTY02% = 49        '②　数量
Private Const ptxG_ST_LOCATION02% = 50      '②　棚番

Private Const ptxG_HIN_GAI03% = 51          '③　外装資材№
Private Const ptxG_HIN_NAME03% = 52         '③　外装資材名称
Private Const ptxG_QTY03% = 53              '③　員数
Private Const ptxG_SHIJI_QTY03% = 54        '③　数量
Private Const ptxG_ST_LOCATION03% = 55      '③　棚番

Private Const ptxD_HIN_GAI01% = 56          '①　同梱／構成品番
Private Const ptxD_HIN_NAME01% = 57         '①　同梱／構成品目
Private Const ptxD_QTY01% = 58              '①　員数
Private Const ptxD_SHIJI_QTY01% = 59        '①　数量
Private Const ptxD_ST_LOCATION01% = 60      '①　棚番
Private Const ptxD_ZAIKO_QTY01% = 61        '①　在庫数
Private Const ptxD_BIKOU01% = 62            '①　備考

Private Const ptxD_HIN_GAI02% = 63          '②　同梱／構成品番
Private Const ptxD_HIN_NAME02% = 64         '②　同梱／構成品目
Private Const ptxD_QTY02% = 65              '②　員数
Private Const ptxD_SHIJI_QTY02% = 66        '②　数量
Private Const ptxD_ST_LOCATION02% = 67      '②　棚番
Private Const ptxD_ZAIKO_QTY02% = 68        '②　在庫数
Private Const ptxD_BIKOU02% = 69            '②　備考

Private Const ptxD_HIN_GAI03% = 70          '③　同梱／構成品番
Private Const ptxD_HIN_NAME03% = 71         '③　同梱／構成品目
Private Const ptxD_QTY03% = 72              '③　員数
Private Const ptxD_SHIJI_QTY03% = 73        '③　数量
Private Const ptxD_ST_LOCATION03% = 74      '③　棚番
Private Const ptxD_ZAIKO_QTY03% = 75        '③　在庫数
Private Const ptxD_BIKOU03% = 76            '③　備考

Private Const ptxD_HIN_GAI04% = 77          '④　同梱／構成品番
Private Const ptxD_HIN_NAME04% = 78         '④　同梱／構成品目
Private Const ptxD_QTY04% = 79              '④　員数
Private Const ptxD_SHIJI_QTY04% = 80        '④　数量
Private Const ptxD_ST_LOCATION04% = 81      '④　棚番
Private Const ptxD_ZAIKO_QTY04% = 82        '④　在庫数
Private Const ptxD_BIKOU04% = 83            '④　備考

Private Const ptxD_HIN_GAI05% = 84          '⑤　同梱／構成品番
Private Const ptxD_HIN_NAME05% = 85         '⑤　同梱／構成品目
Private Const ptxD_QTY05% = 86              '⑤　員数
Private Const ptxD_SHIJI_QTY05% = 87        '⑤　数量
Private Const ptxD_ST_LOCATION05% = 88      '⑤　棚番
Private Const ptxD_ZAIKO_QTY05% = 89        '⑤　在庫数
Private Const ptxD_BIKOU05% = 90            '⑤　備考

Private Const ptxD_HIN_GAI06% = 91          '⑥　同梱／構成品番
Private Const ptxD_HIN_NAME06% = 92         '⑥　同梱／構成品目
Private Const ptxD_QTY06% = 93              '⑥　員数
Private Const ptxD_SHIJI_QTY06% = 94        '⑥　数量
Private Const ptxD_ST_LOCATION06% = 95      '⑥　棚番
Private Const ptxD_ZAIKO_QTY06% = 96        '⑥　在庫数
Private Const ptxD_BIKOU06% = 97            '⑥　備考


Private Const ptxLabel_QTY% = 98            'ラベル発行枚数 2007.12.11




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

'チェック用添字
Private Const pchkSAMPLE_F% = 0             '見本作成
Private Const pchkPRI_SHIJI% = 1            '出力対象　指図票
Private Const pchkPRI_PARTS% = 2            '出力対象　ﾊﾟｰﾂﾗﾍﾞﾙ
Private Const pchkPRI_GAISOU% = 3           '出力対象　外装ﾗﾍﾞﾙ
Private Const pchkPRI_KISHU% = 4            '出力対象　機種ﾗﾍﾞﾙ

Private Const pchkL_PAPER% = 5              '紙             2010.11.12
Private Const pchkL_PLASTIC% = 6            'ﾌﾟﾗｽﾁｯｸ        2010.11.12
Private Const pchkL_LABEL% = 7              '適用機種ﾗﾍﾞﾙ   2010.11.12

'ｵﾌﾟｼｮﾝﾎﾞﾀﾝ用添字
Private Const poptSHIJI_NORMAL% = 0         '通常
Private Const poptSHIJI_SPOT% = 1           'スポット
Private Const poptSHIJI_KEPPIN% = 2         '欠品解除


'リッチテキスト用添字
Private Const prchBIKOU% = 0                '備考



'コマンドボタン固有操作
Private Const cmdMUPDATE% = 3               'ﾏｽﾀ更新

Private Const cmdNext% = 5                  '構成部品画面へ
Private Const cmdCen% = 10                  '取り消し

Private GENSANKOKU_FLG  As String * 1       '原産国　印字有無   2008.06.13


Private wkSURYO         As Long             '208
Private chenge_F        As Boolean          '2008.07.30
Private svJGYOBU        As String * 1       '2008.07.30
Private svNAIGAI        As String * 1       '2008.07.30

Private svSHIMUKE       As String * 4       '2019.06.11 追加


Private GENSANKOKU_CHECK_TBL _
                        As Variant          '原産国ﾁｪｯｸ有無（事業部） 2009.03.28

Private L_GENSANKOKU    As String           '2009.03.28

Private KAISYA_CHK_F    As Boolean          '会社／事業部のエラーﾁｪｯｸ有無 2010.07.20

Private KISHU_CHECK     As Boolean          '代表機種のﾁｪｯｸ 2012.09.03

Private GAI_BUHIN_CHECK As Boolean          '海外供給区分ﾁｪｯｸ有無   2016.02.01

Private TANKA_SPACE_F   As String           '2016.02.01

Private KAISYA_RESTRICT_F   As String


Private SHIMUKE_CHK_TBL As Variant          '半製品　仕向け先   2013.08.29
Private svSHIMUKE_CODE  As String * 2       '2013.08.29

Private LABEL_PRINT_F       As Integer      'ラベル印刷デフォルト表示   2019.03.07
Private GA_LABEL_PRINT_F    As Integer      '外装ラベル印刷デフォルト表示   2019.03.07

Dim L_print_Flg     As Boolean

'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.04.18 10:00)"
'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.04.18 11:45)"
'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.05.27 18:05)" '高沢
'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.05.28 13:50)" '高沢
'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.05.28 16:15)"  '高沢
'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.06.02 17:15)"  '高沢
'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.06.04 17:15)"  '高沢
'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.06.04 11:00)"  '高沢
'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.06.05 11:30)"  '高沢
'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.06.10 18:30)"  '高沢
'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.06.11 11:30)"  '高沢
'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.06.11 16:20)"  '高沢
'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.06.12 17:35)"  '高沢
'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.06.18 10:55)"  '高沢
'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.06.30 20:55)"  '高沢
'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.08.27 11:50)テスト版"  '高沢   品目入力後の処理で一部変更
'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.08.28 10:25)"  '高沢   品目入力後の処理で一部変更
'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.09.24 13:35)"  '高沢   Init_Proc2に一部追加
'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.11.07 15:30) 出庫予定バーコード対応"
'Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.12.18 12:00) 出力対象前回表示が残る件を修正(適用機種ラベル)"
Private Const Last_Update_day$ = "商品化指図票発行 (PI00010 2019.12.18 16:30) 半商品化パーツラベルなし切替対応"

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PI000101.MousePointer = vbHourglass

    Call Ctrl_Lock(PI000101)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PI000101)


    PI000101.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer, chk As Integer, flg, Optional opt As Integer = 0) As Integer
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

Dim com         As Integer

Dim wkTanto     As String

Dim L           As Integer  '2011.02.10

Dim m           As Integer  '2013.01.17


Dim Shimuke_flg    As Integer  '2013.09.04

Dim wkGENSANKOKU    As String * 20  '2015.10.09


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

                    If Text1(ptxSHIJI_NO).text = StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode) Then
                    Else

Start_Proc1:        '2015.03.13


                        chenge_F = False
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
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                If sts > 3000 Or sts = 3 Then
            
                
                                    Call File_Error(sts, BtOpGetEqual, "指図票(親)", 0)
                                    '>>>>>>>>>>>>>  2015.04.24
                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                    'If sts Then
                                    '    Call File_Error(sts, BtOpReset, "")
                                    'End If
                                
                                
                                    'Call File_Open_Proc
                                    Do
                                        If Not File_Open_Proc() Then
                                            Exit Do
                                        End If
                                    Loop
                                    '>>>>>>>>>>>>>  2015.04.24
                    
                                
                                    GoTo Start_Proc1
                                End If
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                Exit Function
                        End Select
                        Text1(Mode).SetFocus        '2008.01.15

                    End If
                End If


                '=========================================== 2007/03/19 =====
''                Text1(ptxSHIJI_NO).BackColor = G_INPUT_NG
''                Text1(ptxSHIJI_NO).Locked = True
''                Text1(ptxSHIJI_NO).TabStop = False
                '============================================================



            End If

        Case ptxHAKKO_DT    '発行日

            If chk = 1 Then
            Else
                If Not IsDate(Text1(ptxHAKKO_DT).text) Then
                    MsgBox "入力した項目はエラーです。(発行日)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxHAKKO_DT).text = Format(CDate(Text1(ptxHAKKO_DT).text), "YYYY/MM/DD")
                End If
            End If

        Case ptxTANTO_CODE      '担当者

           If chk = 1 Then
            Else
                
Start_Proc2:    '2015.03.13
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
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "担当者マスタ", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc2
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        
                        
                        
                        
                        
                        Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                        Exit Function

                End Select
            End If

        Case ptxSHONIN_CODE     '承認者

            If chk = 1 Then
            Else
                
                
Start_Proc3:    '2015.03.13
                
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
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "担当者マスタ", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc3
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        
                        Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                        Exit Function



                End Select
            End If
        Case ptxHIN_GAI         '品番


            '========================================================= 2007/03/19 =====

Start_Proc4:    '2015.03.13

            chenge_F = False
            
            
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


                    lblL_Hin_Name_E.Caption = ""        '2016.02.10
                    lblGAI_BUHIN.Caption = ""           '2016.02.10
                    lblL_URIKIN2.Caption = ""           '2016.02.10
                    lblL_URIKIN3.Caption = ""           '2016.02.10

'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    Check1(pchkL_PAPER).Value = vbUnchecked         '紙
                    Check1(pchkL_PLASTIC).Value = vbUnchecked       'プラ
                    Check1(pchkL_LABEL).Value = vbUnchecked         '適用機種ラベル
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> MT_2009.06.01
                    'MsgBox "入力した項目はエラーです。(品番)"
                    'Text1(Mode).SetFocus
                    'Exit Function

                    lblGensankoku(0).Caption = ""
                    lblGensankoku(1).Caption = ""


                    If Trim(Text1(ptxHIN_GAI).text) = "" Then
                        MsgBox "入力した項目はエラーです。(品番)"
                        Text1(Mode).SetFocus
                        Exit Function
                    End If

                    wkTanto = Text1(ptxTANTO_CODE)
                    If Trim(wkTanto) = "" Then
                        wkTanto = "PSHIJ"
                    End If

                    Last_JGYOBU = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1)
                    If PN_CHK(Text1(Mode), "G", wkTanto, 1) Then          '外部品番
                        Text1(Mode).SetFocus
                        Call Text1_GotFocus(Mode)
                        Exit Function
                    End If

                    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


                    lblL_KAISHA.Caption = ""
                    lblL_JGYOBU.Caption = ""
                    

                    lblKISHU1.Caption = ""          '2016.02.01
                    lblKISHU2.Caption = ""          '2016.02.01
                Case Else
                    
                    
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24

                        
                            GoTo Start_Proc4
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                    
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function

            End Select

'''            yn = False
'''
'''            For i = 0 To Combo1(pcmbSHIMUKE).ListCount - 1
'''                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).List(i), 4), 3, 1))
'''                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).List(i), 4), 4, 1))
'''                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)
'''
'''
'''                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'''                Select Case sts
'''                    Case BtNoErr
'''                        Combo1(pcmbSHIMUKE).ListIndex = i
'''                        yn = True
'''                        Exit For
'''                    Case BtErrKeyNotFound
'''
'''                    Case Else
'''                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
'''                        Exit Function
'''
'''                End Select
'''            Next i
'''
'''            If yn = False Then
'''                Text1(ptxHIN_NAME).text = ""
'''                Text1(ptxST_LOCATION).text = ""
'''                Text1(ptxMI_QTY).text = ""
'''                Text1(ptxSUMI_QTY).text = ""
'''
'''                MsgBox "入力した項目はエラーです。(品番)"
'''                Text1(Mode).SetFocus
'''                Exit Function
'''            End If

            '==========================================================================

            Text1(ptxHIN_NAME).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)


            lblL_Hin_Name_E.Caption = StrConv(ITEMREC.L_HIN_NAME_E, vbUnicode)  '2016.02.10

            lblGAI_BUHIN.Caption = StrConv(ITEMREC.GAI_BUHIN, vbUnicode)        '2016.02.10
            lblL_URIKIN2.Caption = StrConv(ITEMREC.L_URIKIN2, vbUnicode)        '2016.02.10
            lblL_URIKIN3.Caption = StrConv(ITEMREC.L_URIKIN3, vbUnicode)        '2016.02.10

'2013.09.04 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            Shimuke_flg = False
'            If svSHIMUKE_CODE <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Then '2019/12/18 コメントアウト
                For i = 0 To UBound(SHIMUKE_CHK_TBL)
                
                    If SHIMUKE_CHK_TBL(i) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Then
                        Shimuke_flg = True
                        Exit For
                    End If
                
                Next i
'            End If
'2013.09.04 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            'If opt <> 9 Then                   '2013.09.04
            If opt <> 9 And Not Shimuke_flg Then      '2013.09.04
                If StrConv(ITEMREC.L_PAPER, vbUnicode) = L_PAPER_ON Then        '紙
                    Check1(pchkL_PAPER).Value = vbChecked
                Else
                    Check1(pchkL_PAPER).Value = vbUnchecked
                End If
    
                If StrConv(ITEMREC.L_PLASTIC, vbUnicode) = L_PLASTIC_ON Then    'プラ
                    Check1(pchkL_PLASTIC).Value = vbChecked
                Else
                    Check1(pchkL_PLASTIC).Value = vbUnchecked
                End If
    
                If StrConv(ITEMREC.L_LABEL, vbUnicode) = L_LABEL_ON Then        '適用機種ラベル
                    Check1(pchkL_LABEL).Value = vbChecked
                Else
                    Check1(pchkL_LABEL).Value = vbUnchecked
                End If
            
'                If LABEL_PRINT_F = 1 Then          '2019/12/18 <> 1 → = 1 へ変更
                    '2011.02.10
'                    If Trim(StrConv(ITEMREC.L_LABEL, vbUnicode)) = "" Then
'                        Combo2(0).ListIndex = 1
                    '2019.08.27 岸見さんからの指示で下記とした。というか、テスト版
'                    If Trim(StrConv(ITEMREC.L_LABEL, vbUnicode)) = "" Or Trim(StrConv(ITEMREC.L_LABEL, vbUnicode)) = "0" Then
'                        Combo2(0).ListIndex = 1
'                    Else
                        For L = 1 To Combo2(0).ListCount
                            If StrConv(ITEMREC.L_LABEL, vbUnicode) = Right(Combo2(0).List(L), 1) Then
                                Combo2(0).ListIndex = L
                                Exit For
                            
                            End If
                        Next L
'                    End If
'               End If                              '2019.03.07
                
                
                If Trim(StrConv(ITEMREC.L_LABEL, vbUnicode)) = "" Or Trim(StrConv(ITEMREC.L_LABEL, vbUnicode)) = "2" Then
                    Check1(pchkL_PAPER).Enabled = False
                    Check1(pchkL_PLASTIC).Enabled = False
                Else
                    Check1(pchkL_PAPER).Enabled = True
                    Check1(pchkL_PLASTIC).Enabled = True
                End If
                
                '2011.02.10
            
            
            
            
            End If
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<



            If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                Text1(ptxST_LOCATION).text = ""
            Else
                Text1(ptxST_LOCATION).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                StrConv(ITEMREC.ST_DAN, vbUnicode)
            End If

            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                Exit Function

            End If

            Text1(ptxMI_QTY).text = Format(Mi_Qty, "#0")
            Text1(ptxSUMI_QTY).text = Format(Sumi_Qty, "#0")



            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    TORI_GENSANKOKUの有無チェック＆書き込み   2012.01.31

'            chk_TORI_GENSANKOKU = StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)           '原産国有無ﾁｪｯｸ用   2013.01.08

            
            '>>>>>>>>>>>>>>>>2015.10.09
            wkGENSANKOKU = StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)
            For m = 1 To 20
                If Mid(wkGENSANKOKU, m, 1) < " " Then
                    Mid(wkGENSANKOKU, m, 1) = " "
                End If
            Next m
            '>>>>>>>>>>>>>>>>2015.10.09
            
'            If Trim(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) = "" Then     '2015.10.09
            If Trim(wkGENSANKOKU) = "" Then                                     '2015.10.09
            Else
                
Start_Proc5:        '2015.03.13
                
                Call UniCode_Conv(K0_GENSAN.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_GENSAN.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_GENSAN.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(K0_GENSAN.GENSANKOKU, StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode))
                
                sts = BTRV(BtOpGetEqual, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    
                        Call UniCode_Conv(GENSANREC.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                        Call UniCode_Conv(GENSANREC.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                        Call UniCode_Conv(GENSANREC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                        Call UniCode_Conv(GENSANREC.GENSANKOKU, StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode))
                        Call UniCode_Conv(GENSANREC.FILLER, "")
                
                        Call UniCode_Conv(GENSANREC.INS_TANTO, "PI010")
                        Call UniCode_Conv(GENSANREC.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))
                
                        Call UniCode_Conv(GENSANREC.UPD_TANTO, "")
                        Call UniCode_Conv(GENSANREC.UPD_DATETIME, "")
                    
                    
                        sts = BTRV(BtOpInsert, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                        Select Case sts
                        
                            Case BtNoErr
                            Case BtErrDuplicates
                            Case Else
                                
                                
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                If sts > 3000 Or sts = 3 Then
            
                
                                    Call File_Error(sts, BtOpGetEqual, "原産国マスタ", 0)
                                    '>>>>>>>>>>>>>  2015.04.24
                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                    'If sts Then
                                    '    Call File_Error(sts, BtOpReset, "")
                                    'End If
                                
                                
                                    'Call File_Open_Proc
                                    Do
                                        If Not File_Open_Proc() Then
                                            Exit Do
                                        End If
                                    Loop
                                    '>>>>>>>>>>>>>  2015.04.24
                    
                                
                                    GoTo Start_Proc5
                                End If
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                
                                
                                
                                Call File_Error(sts, com, "原産国マスタ")
                                Exit Function
                        End Select
                    
                    
                    
                    
                    Case Else
                        Exit Function
                End Select
            End If
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    TORI_GENSANKOKUの有無チェック＆書き込み   2012.01.31







            txGensankoku.text = Trim(StrConv(ITEMREC.GENSANKOKU, vbUnicode))            '2009.03.28
            chk_TORI_GENSANKOKU = Trim(StrConv(ITEMREC.GENSANKOKU, vbUnicode))          '原産国有無ﾁｪｯｸ用   2013.01.08


            For m = 1 To Len(chk_TORI_GENSANKOKU)
                If Mid(chk_TORI_GENSANKOKU, m, 1) < " " Then
                    chk_TORI_GENSANKOKU = ""
                End If
            Next m
            

            '2010.07.20 ▽
            
Start_Proc6:        '2013.03.13
            
            lstGensankoku.Clear

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
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "指図票(親)", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc6
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        
                        
                        
                        Exit Function
                End Select
                
                'lstGensankoku.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)    2013.02.19

'                If Trim(StrConv(GENSANREC.UPD_DATETIME, vbUnicode)) = "" Then
'                    lstGensankoku.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2013.02.19    2014.02.18
'                Else                                                                                                                '2013.02.19    2014.02.18
'                    lstGensankoku.AddItem StrConv(GENSANREC.UPD_DATETIME, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2013.02.19    2014.02.18
'                End If


                If StrConv(GENSANREC.UPD_DATETIME, vbUnicode) > StrConv(GENSANREC.Ins_DateTime, vbUnicode) Then                     '2014.02.18
                    lstGensankoku.AddItem StrConv(GENSANREC.UPD_DATETIME, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2014.02.18
                Else                                                                                                                '2014.02.18
                    lstGensankoku.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2014.02.18
                End If



                com = BtOpGetNext
            Loop

            lblGensankoku(0).Caption = ""
            lblGensankoku(1).Caption = ""
            If lstGensankoku.ListCount < 1 Then
                lblGensankoku(1).Caption = Trim(StrConv(ITEMREC.GENSANKOKU, vbUnicode))
                If Trim(lblGensankoku(1).Caption) <> "" Then
                    lblGensankoku(0).Caption = ""
                End If
            Else

                lblGensankoku(1).Caption = Right(lstGensankoku.List(lstGensankoku.ListCount - 1), 20)
                lblGensankoku(0).Caption = "ｘ" & StrConv(Format(lstGensankoku.ListCount, "#0"), vbWide)
            End If
            txGensankoku.text = Trim(lblGensankoku(1).Caption)


            lblL_KAISHA.Caption = Trim(StrConv(ITEMREC.L_KAISHA_CODE, vbUnicode))
            lblL_JGYOBU.Caption = Trim(StrConv(ITEMREC.L_JGYOBU_CODE, vbUnicode))



            '名称セット2016.02.01
Start_Proc6_1:        '2016.02.01
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN07_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, lblL_KAISHA.Caption)
            
            


            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                    lblL_KAISHA_N.Caption = StrConv(P_CODEREC.C_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    lblL_KAISHA_N.Caption = lblL_KAISHA.Caption
                Case Else
                    If sts > 3000 Or sts = 3 Then

    
                        Call File_Error(sts, BtOpGetEqual, "ｺｰﾄﾞﾏｽﾀ", 0)
                        Do
                            If Not File_Open_Proc() Then
                                Exit Do
                            End If
                        Loop
                        GoTo Start_Proc6_1
                    End If
                    Exit Function
            End Select
            
            
Start_Proc6_2:        '2016.02.01
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN07_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, lblL_JGYOBU.Caption)
            
            
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                     lblL_JGYOBU_N.Caption = StrConv(P_CODEREC.C_NAME, vbUnicode)
                Case BtErrKeyNotFound
                     lblL_JGYOBU_N.Caption = lblL_JGYOBU.Caption
                Case Else
                    If sts > 3000 Or sts = 3 Then

    
                        Call File_Error(sts, BtOpGetEqual, "ｺｰﾄﾞﾏｽﾀ", 0)
                        Do
                            If Not File_Open_Proc() Then
                                Exit Do
                            End If
                        Loop
                        GoTo Start_Proc6_2
                    End If
                    Exit Function
            End Select
            
            
            
            
            '名称セット2016.02.01



            '2010.07.20 △




            lblKISHU1.Caption = Trim(StrConv(ITEMREC.L_KISHU1, vbUnicode))          '2016.02.01
            lblKISHU2.Caption = Trim(StrConv(ITEMREC.L_KISHU2, vbUnicode))          '2016.02.01








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
                        Text1(Mode).SetFocus         '2008.01.15
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
                        Text1(Mode).SetFocus         '2008.01.15


                    End If
                End If
            End If
            
            '2019.06.10
            Text1(ptxSHIJI_QTY).SetFocus
            
        Case ptxSHIJI_QTY       '数量

            If chk = 1 Then
            Else
                If Not IsNumeric(Text1(ptxSHIJI_QTY).text) Then
                    MsgBox "入力した項目はエラーです。(数量)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxSHIJI_QTY).text = Format(CLng(Text1(ptxSHIJI_QTY).text), "#0")


                    If Trim(Text1(ptxLabel_QTY).text) = "" Then '2008.02.06
'                        Text1(ptxLabel_QTY).text = Format(CLng(Text1(ptxSHIJI_QTY).text) + 1, "#0")                '2015.04.02
                        Text1(ptxLabel_QTY).text = Format(CLng(Text1(ptxSHIJI_QTY).text) + LABEL_PLUS, "#0")        '2015.04.02
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

Start_Proc9:        '2015.03.13
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
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "商品化ｸﾗｽ", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc9
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        
                        
                        
                        Call File_Error(sts, BtOpGetEqual, "商品化ｸﾗｽ")
                        Exit Function

                End Select
            End If
        Case ptxF_CLASS_CODE    '付加ｸﾗｽ

            If Trim(Text1(ptxF_CLASS_CODE).text) = "" Then
            Else
                
                
Start_Proc10:       '2015.03.13
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
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "商品化ｸﾗｽ", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc10
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        
                        
                        Call File_Error(sts, BtOpGetEqual, "商品化ｸﾗｽ")
                        Exit Function

                End Select
            End If

        Case ptxN_CLASS_CODE    '内職ｸﾗｽ

            If Trim(Text1(ptxN_CLASS_CODE).text) = "" Then
            Else
                
Start_Proc11:       '2015.03.13
                
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
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "商品化ｸﾗｽ", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc11
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        
                        
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
                
Start_Proc12:       '2015.03.13
                
                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)

                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        '資材品で読み替え
Start_Proc12_2:       '2015.03.13

                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)

                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound

                                If HIN_INV Then
                                    Call Rclr_ITEMREC                               '2019.06.02 １行追加（高沢）
                                    '未登録品番　可　資材としておく
                                    Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                                    Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                                    Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番")
                                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                                Else
                                    MsgBox "入力した項目はエラーです。(個装資材　品番)"
                                    Text1(Mode).SetFocus
                                    Exit Function
                                End If
                            Case Else
                                
                                
                                
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                If sts > 3000 Or sts = 3 Then
            
                
                                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ", 0)
                                    '>>>>>>>>>>>>>  2015.04.24
                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                    'If sts Then
                                    '    Call File_Error(sts, BtOpReset, "")
                                    'End If
                                
                                
                                    'Call File_Open_Proc
                                    Do
                                        If Not File_Open_Proc() Then
                                            Exit Do
                                        End If
                                    Loop
                                    '>>>>>>>>>>>>>  2015.04.24
                    
                                
                                    GoTo Start_Proc12_2
                                End If
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                
                                
                                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                Exit Function

                        End Select

                    Case Else
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc12
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        
                        
                        
                        Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                        Exit Function

                End Select

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
                If Trim(Text1(Mode - 2).text) <> "" Then
                    MsgBox "入力した項目はエラーです。(個装資材　員数)"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            Else
                If Trim(Text1(Mode - 2).text) = "" Then
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
                
Start_Proc13:   '2015.03.13
                
                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)

                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        '資材品で読み替え
Start_Proc14:   '2015.03.13
                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)

                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound

                                If HIN_INV Then
                                    Call Rclr_ITEMREC                               '2019.06.02 １行追加（高沢）
                                    '未登録品番　可　資材としておく
                                    Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                                    Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                                    Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番")
                                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                                Else

                                    MsgBox "入力した項目はエラーです。(外装資材　品番)"
                                    Text1(Mode).SetFocus
                                    Exit Function
                                End If
                            Case Else
                                
                                
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                If sts > 3000 Or sts = 3 Then
            
                
                                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ", 0)
                                    '>>>>>>>>>>>>>  2015.04.24
                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                    'If sts Then
                                    '    Call File_Error(sts, BtOpReset, "")
                                    'End If
                                    'Call File_Open_Proc
                                    Do
                                        If Not File_Open_Proc() Then
                                            Exit Do
                                        End If
                                    Loop
                                    '>>>>>>>>>>>>>  2015.04.24
                    
                                
                                    GoTo Start_Proc14
                                End If
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                
                                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                Exit Function

                        End Select

                    Case Else
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc13
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        
                        
                        
                        Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                        Exit Function

                End Select

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
                
Start_Proc15:       '2015.03.13
                
                
                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)

                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
Start_Proc16:       '2015.03.13

                        '品番（内）で読み替え
                        Call UniCode_Conv(K2_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                        Call UniCode_Conv(K2_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).text, 4), 4, 1))
                        Call UniCode_Conv(K2_ITEM.HIN_NAI, Text1(Mode).text)

                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound


Start_Proc17:       '2015.03.13




                                '資材品で読み替え

                                Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)

                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                Select Case sts
                                    Case BtNoErr
                                    Case BtErrKeyNotFound

                                        If HIN_INV Then
                                            Call Rclr_ITEMREC                               '2019.06.02 １行追加（高沢）
                                            '未登録品番　可　資材としておく
                                            Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                                            Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                                            Call UniCode_Conv(ITEMREC.HIN_GAI, Text1(Mode).text)
                                            Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番")
                                            Call UniCode_Conv(ITEMREC.ST_SOKO, "")

                                        Else

                                            MsgBox "入力した項目はエラーです。(同梱／構成　品番)"
                                            Text1(Mode).SetFocus
                                            Exit Function
                                        End If
                                    Case Else
                                        
                                        
                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                        If sts > 3000 Or sts = 3 Then
                    
                        
                                            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ", 0)
                                            '>>>>>>>>>>>>>  2015.04.24
                                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                            'If sts Then
                                            '    Call File_Error(sts, BtOpReset, "")
                                            'End If
                                        
                                        
                                            'Call File_Open_Proc
                                            Do
                                                If Not File_Open_Proc() Then
                                                    Exit Do
                                                End If
                                            Loop
                                            '>>>>>>>>>>>>>  2015.04.24
                            
                                        
                                            GoTo Start_Proc17
                                        End If
                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                        
                                        
                                        Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                        Exit Function

                                End Select

                            Case Else
                                
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                If sts > 3000 Or sts = 3 Then
            
                
                                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ", 0)
                                    '>>>>>>>>>>>>>  2015.04.24
                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                    'If sts Then
                                    '    Call File_Error(sts, BtOpReset, "")
                                    'End If
                                
                                
                                    'Call File_Open_Proc
                                    Do
                                        If Not File_Open_Proc() Then
                                            Exit Do
                                        End If
                                    Loop
                                    '>>>>>>>>>>>>>  2015.04.24
                    
                                
                                    GoTo Start_Proc16
                                End If
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                
                                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                Exit Function
                       End Select

                    Case Else
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc15
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                        Exit Function

                End Select

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

                '在庫数
                If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                        StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                        StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
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

                D_Item_Tbl(i).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                D_Item_Tbl(i).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
                D_Item_Tbl(i).HIN_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)


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
                End If
            End If
                
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

        Case ptxLabel_QTY       'ラベル発行枚数 2007.12.11

            If chk = 1 Then
            Else
                If Not IsNumeric(Text1(ptxSHIJI_QTY).text) Then
                    MsgBox "入力した項目はエラーです。(ラベル発行枚数)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxLabel_QTY).text = Format(CLng(Text1(ptxLabel_QTY).text), "#0")
                    If CLng(Text1(ptxLabel_QTY).text) <= 0 Then
                        MsgBox "入力した項目はエラーです。(ラベル発行枚数)"
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
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


Dim L           As Integer  '2011.02.10

Dim Wk_LOC      As String   '2013.01.07


Dim Ret_sts     As Integer  '2015.03.13

Dim Zaiko_sts   As Integer  '2015.03.13


    Item_Disp_Proc = True

    Call Input_Lock         '2008.01.15

    For i = ptxK_HIN_GAI01 To ptxD_BIKOU06
        Text1(i).text = ""
    Next i

                                '2008.07.30
    For i = pcmbD_SYUBETSU01 To pcmbD_SYUBETSU06

            Combo1(i).ListIndex = -1

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

    Text1(ptxLabel_QTY).text = ""       '2008.02.27


    '--------------------------------   「親」情報


    Text1(ptxSHIJI_NO).text = StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode)           '指図票№
                                                                                    '発行日
    Text1(ptxHAKKO_DT).text = Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 7, 2)

    Text1(ptxTANTO_CODE).text = StrConv(P_SSHIJI_O_REC.TANTO_CODE, vbUnicode)       '担当者ｺｰﾄﾞ／名称
    
    
Start_Proc1:    '2015.03.13
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTANTO_CODE).text)

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxTANTO_NAME).text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text1(ptxTANTO_NAME).text = ""
        Case Else
            
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
            If sts > 3000 Or sts = 3 Then


                Call File_Error(sts, BtOpGetEqual, "担当者マスタ", 0)
                '>>>>>>>>>>>>>  2015.04.24
                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                'If sts Then
                '    Call File_Error(sts, BtOpReset, "")
                'End If
            
            
                'Call File_Open_Proc
                Do
                    If Not File_Open_Proc() Then
                        Exit Do
                    End If
                Loop
                '>>>>>>>>>>>>>  2015.04.24

            
                GoTo Start_Proc1
            End If
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
            
            
            Call Input_UnLock         '2008.01.15
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Exit Function

    End Select

    Text1(ptxSHONIN_CODE).text = StrConv(P_SSHIJI_O_REC.SHONIN_CODE, vbUnicode)     '承認者ｺｰﾄﾞ／名称
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxSHONIN_CODE).text)

    
Start_Proc2:        '2015.03.13
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxSHONIN_NAME).text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text1(ptxSHONIN_NAME).text = ""
        Case Else
            
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
            If sts > 3000 Or sts = 3 Then


                Call File_Error(sts, BtOpGetEqual, "担当者マスタ", 0)
                '>>>>>>>>>>>>>  2015.04.24
                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                'If sts Then
                '    Call File_Error(sts, BtOpReset, "")
                'End If
            
            
                'Call File_Open_Proc
                Do
                    If Not File_Open_Proc() Then
                        Exit Do
                    End If
                Loop
                '>>>>>>>>>>>>>  2015.04.24

            
                GoTo Start_Proc2
            End If
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
            
            
            Call Input_UnLock         '2008.01.15
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

Start_Proc3:    '2015.03.13

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





            lblKISHU1.Caption = Trim(StrConv(ITEMREC.L_KISHU1, vbUnicode))          '2012.10.26
            lblKISHU2.Caption = Trim(StrConv(ITEMREC.L_KISHU2, vbUnicode))          '2012.10.26


'>>>>>>>>>>>>>>>>>> 2015.03.13
'            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
'                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
'                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), , , , , , Ret_sts) Then
'                Exit Function
'
'            End If


Start_Proc4:        '2015.03.13
            Zaiko_sts = Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), , , , , , Ret_sts, 0)
             If Zaiko_sts Then
                If Ret_sts > 3000 Or Ret_sts = 3 Then


                    Call File_Error(Ret_sts, BtOpGetEqual, "在庫ﾃﾞｰﾀ", 0)
                    '>>>>>>>>>>>>>  2015.04.24
                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                    'If sts Then
                    '    Call File_Error(sts, BtOpReset, "")
                    'End If
                
                
                    'Call File_Open_Proc
                    Do
                        If Not File_Open_Proc() Then
                            Exit Do
                        End If
                    Loop
                    '>>>>>>>>>>>>>  2015.04.24
                
                    GoTo Start_Proc4
                End If
                Call File_Error(Ret_sts, BtOpGetEqual, "在庫ﾃﾞｰﾀ")
                Exit Function
            End If
'>>>>>>>>>>>>>>>>>> 2015.03.13

            Text1(ptxMI_QTY).text = Format(Mi_Qty, "#0")
            Text1(ptxSUMI_QTY).text = Format(Sumi_Qty, "#0")


'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            If StrConv(ITEMREC.L_PAPER, vbUnicode) = L_PAPER_ON Then        '紙
                Check1(pchkL_PAPER).Value = vbChecked
            Else
                Check1(pchkL_PAPER).Value = vbUnchecked
            End If

            If StrConv(ITEMREC.L_PLASTIC, vbUnicode) = L_PLASTIC_ON Then    'プラ
                Check1(pchkL_PLASTIC).Value = vbChecked
            Else
                Check1(pchkL_PLASTIC).Value = vbUnchecked
            End If

            If StrConv(ITEMREC.L_LABEL, vbUnicode) = L_LABEL_ON Then        '適用機種ラベル
                Check1(pchkL_LABEL).Value = vbChecked
            Else
                Check1(pchkL_LABEL).Value = vbUnchecked
            End If
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

            '2011.02.10
'            If Trim(StrConv(ITEMREC.L_LABEL, vbUnicode)) = "" Then
'                Combo2(0).ListIndex = 1
            '2019.08.28 ↑を下記に変更
            If Trim(StrConv(ITEMREC.L_LABEL, vbUnicode)) = "" Or Trim(StrConv(ITEMREC.L_LABEL, vbUnicode)) = "0" Then
'''                Combo2(0).ListIndex = 1
            
            Else
                For L = 1 To Combo2(0).ListCount
                
                
                    If StrConv(ITEMREC.L_LABEL, vbUnicode) = Right(Combo2(0).List(L), 1) Then
                
                        Combo2(0).ListIndex = L
                        Exit For
                    
                    End If
                Next L
            
            
            
            
            
            End If
            
            
            
            
            
            
            '2011.02.10




        Case BtErrKeyNotFound
            Text1(ptxHIN_NAME).text = ""
            Text1(ptxST_LOCATION).text = ""
            Text1(ptxMI_QTY).text = ""
            Text1(ptxSUMI_QTY).text = ""

'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            Check1(pchkL_PAPER).Value = vbUnchecked         '紙
            Check1(pchkL_PLASTIC).Value = vbUnchecked       'プラ
            Check1(pchkL_LABEL).Value = vbUnchecked         '適用機種ラベル
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

            lblKISHU1.Caption = ""                      '2012.10.26
            lblKISHU2.Caption = ""                      '2012.10.26


        Case Else
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
          
            If sts > 3000 Or sts = 3 Then


                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                '>>>>>>>>>>>>>  2015.04.24
                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                'If sts Then
                '    Call File_Error(sts, BtOpReset, "")
                'End If
            
            
                'Call File_Open_Proc
                Do
                    If Not File_Open_Proc() Then
                        Exit Do
                    End If
                Loop
                '>>>>>>>>>>>>>  2015.04.24

            
                GoTo Start_Proc3
            End If
            
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
            
            
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
                                                                                    '出力対象　ﾊﾟｰﾂﾗﾍﾞﾙ 2010.07.20
    If StrConv(P_SSHIJI_O_REC.PRI_PARTS, vbUnicode) = P_PRI_PARTS_OFF Or Trim(StrConv(P_SSHIJI_O_REC.PRI_PARTS, vbUnicode)) = "" Then
        Check1(pchkPRI_PARTS).Value = vbUnchecked
    Else
        Check1(pchkPRI_PARTS).Value = vbChecked
    End If


    RichTextBox1(prchBIKOU).text = StrConv(P_SSHIJI_O_REC.BIKOU, vbUnicode)         '備考




    txGensankoku.text = Trim(StrConv(ITEMREC.GENSANKOKU, vbUnicode))                '2009.03.28



    '2010.07.20 ▽
    
Start_Proc4_2:           '2015.03.13
    
    lstGensankoku.Clear

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
                
                
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                If sts > 3000 Or sts = 3 Then


                    Call File_Error(sts, BtOpGetEqual, "原産国ﾏｽﾀ", 0)
                    '>>>>>>>>>>>>>  2015.04.24
                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                    'If sts Then
                    '    Call File_Error(sts, BtOpReset, "")
                    'End If
                
                
                    'Call File_Open_Proc
                    Do
                        If Not File_Open_Proc() Then
                            Exit Do
                        End If
                    Loop
                    '>>>>>>>>>>>>>  2015.04.24
    
                
                    GoTo Start_Proc4_2
                End If
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                
                
                
                Exit Function
        End Select


        'lstGensankoku.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)        2013.01.28
'        If Trim(StrConv(GENSANREC.UPD_DATETIME, vbUnicode)) = "" Then                                                       '2013.01.28     2014.02.18
'            lstGensankoku.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2013.01.28     2014.02.18
'        Else                                                                                                                '2013.01.28     2014.02.18
'            lstGensankoku.AddItem StrConv(GENSANREC.UPD_DATETIME, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2013.01.28     2014.02.18
'        End If

        If StrConv(GENSANREC.UPD_DATETIME, vbUnicode) > StrConv(GENSANREC.Ins_DateTime, vbUnicode) Then                     '2014.02.18
            lstGensankoku.AddItem StrConv(GENSANREC.UPD_DATETIME, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2014.02.18
        Else                                                                                                                '2014.02.18
            lstGensankoku.AddItem StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode)     '2014.02.18
        End If


        com = BtOpGetNext
    Loop

    lblGensankoku(0).Caption = ""
    lblGensankoku(1).Caption = ""
    If lstGensankoku.ListCount < 1 Then
        lblGensankoku(1).Caption = Trim(StrConv(ITEMREC.GENSANKOKU, vbUnicode))
        If Trim(lblGensankoku(1).Caption) <> "" Then
            lblGensankoku(0).Caption = ""
        End If
    Else

        lblGensankoku(1).Caption = Right(lstGensankoku.List(lstGensankoku.ListCount - 1), 20)
        lblGensankoku(0).Caption = "ｘ" & StrConv(Format(lstGensankoku.ListCount, "#0"), vbWide)
    End If
    txGensankoku.text = Trim(lblGensankoku(1).Caption)




    '2010.07.20 ▽



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


Start_Proc5:        '2013.03.13

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
                
                

                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                If sts > 3000 Or sts = 3 Then


                    Call File_Error(sts, BtOpGetEqual, "指図票(子)", 0)
                    '>>>>>>>>>>>>>  2015.04.24
                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                    'If sts Then
                    '    Call File_Error(sts, BtOpReset, "")
                    'End If
                
                
                    'Call File_Open_Proc
                    Do
                        If Not File_Open_Proc() Then
                            Exit Do
                        End If
                    Loop
                    '>>>>>>>>>>>>>  2015.04.24
    
                
                    GoTo Start_Proc5
                End If
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                
                
                
                
                Call Input_UnLock         '2008.01.15
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

Start_Proc6:        '2015.03.13

                Call UniCode_Conv(K0_ITEM.JGYOBU, K_Item_Tbl(k).JGYOBU)
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
                        
Start_Proc7:        '2015.03.13
                        
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
                                
                                
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                If sts > 3000 Or sts = 3 Then
            
                
                                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ", 0)
                                    '>>>>>>>>>>>>>  2015.04.24
                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                    'If sts Then
                                    '    Call File_Error(sts, BtOpReset, "")
                                    'End If
                                
                                
                                    'Call File_Open_Proc
                                    Do
                                        If Not File_Open_Proc() Then
                                            Exit Do
                                        End If
                                    Loop
                                    '>>>>>>>>>>>>>  2015.04.24
                    
                                
                                    GoTo Start_Proc7
                                End If
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                
                                
                                
                                
                                
                                Call Input_UnLock             '2008.01.15
                                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                Exit Function
    
                        End Select
'                        Text1(K_Index + 1) = "未登録品番"
'                        Text1(K_Index + 4) = ""
'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21
                    Case Else
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc6
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        
                        
                        
                        Call Input_UnLock         '2008.01.15
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


Start_Proc8:        '2015.03.13

                Call UniCode_Conv(K0_ITEM.JGYOBU, G_Item_Tbl(g).JGYOBU)
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
Start_Proc9:        '2015.03.13
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
                                
                                
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                If sts > 3000 Or sts = 3 Then
            
                
                                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ", 0)
                                    '>>>>>>>>>>>>>  2015.04.24
                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                    'If sts Then
                                    '    Call File_Error(sts, BtOpReset, "")
                                    'End If
                                
                                
                                    'Call File_Open_Proc
                                    Do
                                        If Not File_Open_Proc() Then
                                            Exit Do
                                        End If
                                    Loop
                                    '>>>>>>>>>>>>>  2015.04.24
                    
                                
                                    GoTo Start_Proc9
                                End If
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                
                                
                                Call Input_UnLock             '2008.01.15
                                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                Exit Function
    
                        End Select
'                        Text1(G_Index + 1) = "未登録品番"
'                        Text1(G_Index + 4) = ""
'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21
                        
                    Case Else
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc8
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                        Call Input_UnLock         '2008.01.15
                        Call File_Error(sts, BtOpGetEqual, "")
                        Exit Function

                End Select


                Text1(G_Index + 2).text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                Text1(G_Index + 3).text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")

                G_Index = G_Index + 5


            Case P_DOUKON   '同梱／構成

                d = d + 1
                D_Item_Tbl(d).JGYOBU = StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode)
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

                    DC_Index = DC_Index + 1

                                '品番
                    Text1(DT_Index).text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)


Start_Proc10:   '2015.03.13

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


' 2013.01.07 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'             同梱部品は標準棚番の在庫数のみ表示する様に変更 大阪ＰＣ以外は現状のまま

Start_Proc11:   '2015.03.13
                            
                            
'>>>>>>>>>>>>>>>    2015.03.13
'                            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
'                                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
'                                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), , , , , , Ret_sts, 0) Then
'                                Exit Function
'                            End If


                            Zaiko_sts = Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), , , , , , Ret_sts, 0)
                                
                            If Zaiko_sts Then
                                If Ret_sts > 3000 Or Ret_sts = 3 Then
                                    Call File_Error(sts, BtOpGetEqual, "在庫ﾃﾞｰﾀ", 0)
                                    '>>>>>>>>>>>>>  2015.04.24
                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                    'If sts Then
                                    '    Call File_Error(sts, BtOpReset, "")
                                    'End If
                                
                                
                                    'Call File_Open_Proc
                                    Do
                                        If Not File_Open_Proc() Then
                                            Exit Do
                                        End If
                                    Loop
                                    '>>>>>>>>>>>>>  2015.04.24
                    
                                
                                    GoTo Start_Proc11
                                End If
                                Call File_Error(sts, BtOpGetEqual, "在庫ﾃﾞｰﾀ")
                                Exit Function
                            End If
'>>>>>>>>>>>>>>>    2015.03.13

'                            Wk_LOC = StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & _
'                                     StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)
'
'                            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
'                                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
'                                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), _
'                                                                    Wk_LOC) Then
'                                Exit Function
'
'                            End If
' 2013.01.07 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

                            Text1(DT_Index + 5) = Format(Sumi_Qty + Mi_Qty, "#0")


                        Case BtErrKeyNotFound




'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21
                            
                            
Start_Proc12:   '2015.03.13
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
                                    
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                If sts > 3000 Or sts = 3 Then
            
                
                                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ", 0)
                                    '>>>>>>>>>>>>>  2015.04.24
                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                    'If sts Then
                                    '    Call File_Error(sts, BtOpReset, "")
                                    'End If
                                
                                
                                    'Call File_Open_Proc
                                    Do
                                        If Not File_Open_Proc() Then
                                            Exit Do
                                        End If
                                    Loop
                                    '>>>>>>>>>>>>>  2015.04.24
                    
                                
                                    GoTo Start_Proc12
                                End If
                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                                    
                                    
                                    Call Input_UnLock             '2008.01.15
                                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                    Exit Function
        
                            End Select
'                            Text1(DT_Index + 1) = "未登録品番"
'                            Text1(DT_Index + 4) = ""
'                            Text1(DT_Index + 5) = ""
'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21
                    
                        Case Else
                            
                            
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ", 0)
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                            
                            
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc10
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
                            
                            
                            Call Input_UnLock         '2008.01.15
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
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

    Call Input_UnLock         '2008.01.15



    If Right(Combo2(0).text, 1) = " " Or Right(Combo2(0).text, 1) = "2" Then
        Check1(pchkL_PAPER).Enabled = False
        Check1(pchkL_PLASTIC).Enabled = False
    Else
        Check1(pchkL_PAPER).Enabled = True
        Check1(pchkL_PLASTIC).Enabled = True
    End If



    Item_Disp_Proc = False

End Function

Private Function Update_Proc(Mode As Integer, MSG As Integer) As Integer
'----------------------------------------------------------------------------
'                   構成マスタ＆商品化指示ﾃﾞｰﾀ出力
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim ans         As Integer
Dim com         As Integer

Dim SEQNO       As Integer

Dim i           As Integer
Dim j           As Integer

Dim SHIJINO     As Long


Dim NEW_F       As Boolean          '2008.05.19

    Update_Proc = True

    Call Input_Lock

                                        
Start_Proc0:        '2015.03.26
                                        
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If

    NEW_F = False       '2008.05.19

    If Text1(ptxSHIJI_NO).text = "" Then
        NEW_F = True    '2008.05.19


        Do                              '<----------------------    2013.10.04

            DoEvents                                                '2013.10.04


            '2008.02.02
            'If IsNumeric(Text1(ptxSHIJI_QTY).text) And CDbl(Text1(ptxSHIJI_QTY).text) = 0 Then
            '    SHIJINO = "00000"
            'Else
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
                            
                            
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "管理マスタ", 0)
    
            
                                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpAbortTransaction, "")
                                End If
    
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                            
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc0
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            
                            
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
                            
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "管理マスタ", 0)
    
            
                                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpAbortTransaction, "")
                                End If
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                            
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc0
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            
                            
                            Call File_Error(sts, BtOpUpdate, "管理マスタ")
                            GoTo Abort_Tran
                    End Select
                Loop
    
                SHIJINO = CLng(StrConv(P_KANRIREC.SASHIZU_NO, vbUnicode))
    
    
                Text1(ptxSHIJI_NO).text = Format(SHIJINO, "00000000")
            'End If
        
                                        '                           2013.10.04  指図データの存在チェックを追加
                Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, Format(SHIJINO, "00000000"))
                sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        Exit Do
                    Case Else
                        
                        
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "指図票ﾃﾞｰﾀ(親)", 0)
    
            
                                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpAbortTransaction, "")
                                End If
    
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                            
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc0
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
                        Call File_Error(sts, BtOpUpdate, "指図票ﾃﾞｰﾀ(親)")
                        GoTo Abort_Tran
                End Select
                                        '                           2013.10.04  指図データの存在チェックを追加
        Loop                            '<----------------------    2013.10.04
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
                    
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
    
            
                                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpAbortTransaction, "")
                                End If
    
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                            
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc0
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                    
                    
                    
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                    GoTo Abort_Tran

            End Select


        Loop


'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        
    If Trim(Right(Combo2(0).text, 1)) = "0" Or Trim(Right(Combo2(0).text, 1)) = "1" Then
        If Check1(pchkL_PAPER).Value = vbChecked Then                           '紙
            Call UniCode_Conv(ITEMREC.L_PAPER, L_PAPER_ON)
        Else
            Call UniCode_Conv(ITEMREC.L_PAPER, L_PAPER_OFF)
        End If

        If Check1(pchkL_PLASTIC).Value = vbChecked Then                         'プラスチック
            Call UniCode_Conv(ITEMREC.L_PLASTIC, L_PLASTIC_ON)
        Else
            Call UniCode_Conv(ITEMREC.L_PLASTIC, L_PLASTIC_OFF)
        End If

        If Check1(pchkL_LABEL).Value = vbChecked Then                           '適用機種ラベル
            Call UniCode_Conv(ITEMREC.L_LABEL, L_LABEL_ON)
        Else
            Call UniCode_Conv(ITEMREC.L_LABEL, L_LABEL_OFF)
        End If
            

    End If

'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


'2011.02.10
    If Trim(Right(Combo2(0).text, 1)) <> "" Then
        Call UniCode_Conv(ITEMREC.L_LABEL, Right(Combo2(0).text, 1))
    End If
'2011.02.10

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


        '2011.02.16
        Call UniCode_Conv(ITEMREC.UPD_TANTO, "PI010")
        Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        


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
                    
                    
                    
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
    
            
                                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpAbortTransaction, "")
                                End If
    
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                            
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc0
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                    
                    
                    Call File_Error(sts, BtOpUpdate, "品目マスタ")
                    GoTo Abort_Tran
            End Select
        Loop

'    End If
    '---------------------------------------------------    '構成マスタ更新

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
                    
                    
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "構成マスタ", 0)
    
            
                                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpAbortTransaction, "")
                                End If
    
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                            
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc0
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                    
                    
                    Call File_Error(sts, com + BtSNoWait, "構成マスタ")
                    GoTo Abort_Tran
            End Select

        Loop

        If sts = BtErrEOF Then
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
                    
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                    If sts > 3000 Or sts = 3 Then

    
                        Call File_Error(sts, BtOpGetEqual, "構成マスタ", 0)

    
                        sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                        If sts <> BtNoErr Then
                            Call File_Error(sts, BtOpAbortTransaction, "")
                        End If

                        '>>>>>>>>>>>>>  2015.04.24
                        'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                        'If sts Then
                        '    Call File_Error(sts, BtOpReset, "")
                        'End If
                    
                        'Call File_Open_Proc
                        Do
                            If Not File_Open_Proc() Then
                                Exit Do
                            End If
                        Loop
                        '>>>>>>>>>>>>>  2015.04.24
        
                    
                        GoTo Start_Proc0
                    End If
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                    
                    
                    
                    Call File_Error(sts, BtOpDelete, "構成マスタ")
                    GoTo Abort_Tran
            End Select
        Loop

        com = BtOpGetNext

    Loop

    '構成マスタ(ﾍｯﾀﾞｰ)出力
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
        
    Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_TANTO, "")                   '品番ﾁｪｯｸ担当者ｺｰﾄﾞ     2013.08.21
    Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_DATETIME, "")                '品番ﾁｪｯｸ日時           2013.08.21
    Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_LABEL_CNT, "")               '品番ﾁｪｯｸﾗﾍﾞﾙ件数       2013.08.21
    Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_GENPIN_CNT, "")              '品番ﾁｪｯｸ現品票件数     2013.08.21


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
                
                
                
                
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                If sts > 3000 Or sts = 3 Then


                    Call File_Error(sts, BtOpGetEqual, "構成マスタ", 0)


                    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                    If sts <> BtNoErr Then
                        Call File_Error(sts, BtOpAbortTransaction, "")
                    End If

                    '>>>>>>>>>>>>>  2015.04.24
                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                    'If sts Then
                    '    Call File_Error(sts, BtOpReset, "")
                    'End If
                
                    'Call File_Open_Proc
                    Do
                        If Not File_Open_Proc() Then
                            Exit Do
                        End If
                    Loop
                    '>>>>>>>>>>>>>  2015.04.24
    
                
                    GoTo Start_Proc0
                End If
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                
                
                
                
                Call File_Error(sts, BtOpInsert, "構成マスタ")
                GoTo Abort_Tran
        End Select

    Loop

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
                        
                        
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "構成マスタ", 0)
    
            
                                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpAbortTransaction, "")
                                End If
    
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                                
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc0
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
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
                        
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "構成マスタ", 0)
    
            
                                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpAbortTransaction, "")
                                End If
                                
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                            
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc0
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
                        Call File_Error(sts, BtOpInsert, "構成マスタ")
                        GoTo Abort_Tran
                End Select

            Loop



        End If

        j = j + 1


    Next i


    '同梱／構成分
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
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "構成マスタ", 0)

        
                            sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                            If sts <> BtNoErr Then
                                Call File_Error(sts, BtOpAbortTransaction, "")
                            End If

                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc0
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
                        Call File_Error(sts, BtOpInsert, "構成マスタ")
                        GoTo Abort_Tran
                End Select

            Loop


        End If

    Next i



    If Mode = 1 Then
        GoTo End_Tran
    End If

    If CDbl(Text1(ptxSHIJI_QTY).text) = 0 Then      '2008.02.02
        GoTo End_Tran
    End If
    '---------------------------------------------------    '指図票データ更新

    '指図票データ(ﾍｯﾀﾞｰ)処理


    Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, Format(SHIJINO, "00000000"))  '2008.02.13

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
                
                
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "商品化指図票ﾃﾞｰﾀ(親)", 0)
    
            
                                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpAbortTransaction, "")
                                End If
                                
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                            
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc0
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "商品化指図票ﾃﾞｰﾀ(親)")
                GoTo Abort_Tran
        End Select

    Loop


    If com = BtOpInsert Then
        '新規作成
        Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_No, Format(SHIJINO, "00000000")) '指図票№   2008.02.13
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

'        Call UniCode_Conv(P_SSHIJI_O_REC.FILLER, "")                           '2016.02.01
        Call UniCode_Conv(P_SSHIJI_O_REC.HIN_CHECK_GAISOU_CNT, "")              '2016.02.01 品番ﾁｪｯｸ外装品番件数

    End If
                                                                                '発行日
    Call UniCode_Conv(P_SSHIJI_O_REC.HAKKO_DT, Format(Text1(ptxHAKKO_DT).text, "YYYYMMDD"))
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
            
            
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
            If sts > 3000 Or sts = 3 Then


                Call File_Error(sts, BtOpGetEqual, "受払先マスタ", 0)


                sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                If sts <> BtNoErr Then
                    Call File_Error(sts, BtOpAbortTransaction, "")
                End If

                '>>>>>>>>>>>>>  2015.04.24
                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                'If sts Then
                '    Call File_Error(sts, BtOpReset, "")
                'End If
            
                'Call File_Open_Proc
                Do
                    If Not File_Open_Proc() Then
                        Exit Do
                    End If
                Loop
                '>>>>>>>>>>>>>  2015.04.24

            
                GoTo Start_Proc0
            End If
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
            
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
            End If
        End If
    End If


    If Check1(pchkPRI_SHIJI).Value = vbChecked Then                             '出力対象　指図票
        Call UniCode_Conv(P_SSHIJI_O_REC.PRI_SHIJI, P_PRI_SHIJI_ON)
        Call UniCode_Conv(P_SSHIJI_O_REC.Print_datetime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
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
                
                
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                If sts > 3000 Or sts = 3 Then


                    Call File_Error(sts, BtOpGetEqual, "商品化指図ﾃﾞｰﾀ(親)", 0)


                    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                    If sts <> BtNoErr Then
                        Call File_Error(sts, BtOpAbortTransaction, "")
                    End If

                    '>>>>>>>>>>>>>  2015.04.24
                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                    'If sts Then
                    '    Call File_Error(sts, BtOpReset, "")
                    'End If
                
                    'Call File_Open_Proc
                    Do
                        If Not File_Open_Proc() Then
                            Exit Do
                        End If
                    Loop
                    '>>>>>>>>>>>>>  2015.04.24
    
                
                    GoTo Start_Proc0
                End If
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                
                
                Call File_Error(sts, com, "商品化指図ﾃﾞｰﾀ(親)")
                GoTo Abort_Tran
        End Select

    Loop

    If com = BtOpUpdate Then
        '対象の子を削除する
        Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, Format(SHIJINO, "00000000"))  '2008.02.13
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
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "商品化指図ﾃﾞｰﾀ(子)", 0)

        
                            sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                            If sts <> BtNoErr Then
                                Call File_Error(sts, BtOpAbortTransaction, "")
                            End If

                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc0
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        Call File_Error(sts, com + BtSNoWait, "商品化指図ﾃﾞｰﾀ(子)")
                        GoTo Abort_Tran
                End Select

            Loop

            If sts = BtErrEOF Then
                Exit Do
            End If


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
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "商品化指図ﾃﾞｰﾀ(子)", 0)

        
                            sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                            If sts <> BtNoErr Then
                                Call File_Error(sts, BtOpAbortTransaction, "")
                            End If

                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc0
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
                        
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
            Call UniCode_Conv(P_SSHIJI_K_REC.SHIJI_No, Format(SHIJINO, "00000000"))     '指図票№ 2008.02.13



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


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI_QTY, "")                          '移動済み数量 2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_TANTO, "")                           '構成ﾁｪｯｸ   担当者          2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_Sumi_Cnt, "")                        '           ﾁｪｯｸ済み数      2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_ALL_Cnt, "")                         '           構成数          2012.04.20
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.04.20
            
            Call UniCode_Conv(P_SSHIJI_K_REC.FILLER, "")
                                                                                        '更新日時
            Call UniCode_Conv(P_SSHIJI_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))

                                                                                        '出荷予定ＩＤ
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_ID_NO, "")

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
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "商品化指図ﾃﾞｰﾀ(子)", 0)

        
                            sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                            If sts <> BtNoErr Then
                                Call File_Error(sts, BtOpAbortTransaction, "")
                            End If

                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc0
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
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

            Call UniCode_Conv(P_SSHIJI_K_REC.SHIJI_No, Format(SHIJINO, "00000000"))     '指図票№   2008.02.13
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

            Call UniCode_Conv(P_SSHIJI_K_REC.FILLER, "")
                                                                                        
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI_QTY, "")                          '移動済み数量 2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_TANTO, "")                           '構成ﾁｪｯｸ   担当者          2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_Sumi_Cnt, "")                        '           ﾁｪｯｸ済み数      2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_ALL_Cnt, "")                         '           構成数          2012.04.20
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.04.20
                                                                                        
                                                                                        
                                                                                        
                                                                                        '更新日時
            Call UniCode_Conv(P_SSHIJI_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))


                                                                                        '出荷予定ＩＤ
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_ID_NO, "")

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
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "商品化指図ﾃﾞｰﾀ(子)", 0)

        
                            sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                            If sts <> BtNoErr Then
                                Call File_Error(sts, BtOpAbortTransaction, "")
                            End If
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
           
                        
                            GoTo Start_Proc0
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
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
            Call UniCode_Conv(P_SSHIJI_K_REC.SHIJI_No, Format(SHIJINO, "00000000"))     '指図票№   2008.02.13

            SEQNO = SEQNO + 10

            Call UniCode_Conv(P_SSHIJI_K_REC.DATA_KBN, P_DOUKON)                        'データ区分
            Call UniCode_Conv(P_SSHIJI_K_REC.SEQNO, Format(SEQNO, "000"))               '追番

            Call UniCode_Conv(P_SSHIJI_K_REC.KO_SYUBETSU, D_Item_Tbl(i).SYUBETSU)       '種別
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_JGYOBU, D_Item_Tbl(i).JGYOBU)           '事業部
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_NAIGAI, D_Item_Tbl(i).NAIGAI)           '国内外
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_HIN_GAI, D_Item_Tbl(i).HIN_GAI)         '品番
                                                                                        '員数
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_QTY, Format(D_Item_Tbl(i).QTY, "000.00"))
                                                                                        '数
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_SHIJI_QTY, Format(D_Item_Tbl(i).SHIJI_QTY, "00000000.00"))
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_BIKOU, D_Item_Tbl(i).BIKOU)             '備考

            Call UniCode_Conv(P_SSHIJI_K_REC.CALCEL_F, P_CANCEL_OFF)                    'ｷｬﾝｾﾙﾌﾗｸﾞ
            Call UniCode_Conv(P_SSHIJI_K_REC.CANCEL_DATETIME, "")                       'ｷｬﾝｾﾙ日時

            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.IDO_SUMI_QTY, "")                          '移動済み数量 2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_TANTO, "")                           '構成ﾁｪｯｸ   担当者          2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_Sumi_Cnt, "")                        '           ﾁｪｯｸ済み数      2012.04.20
            Call UniCode_Conv(P_SSHIJI_K_REC.COMPO_ALL_Cnt, "")                         '           構成数          2012.04.20
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.04.20
            
            
            
            Call UniCode_Conv(P_SSHIJI_K_REC.FILLER, "")
                                                                                        '更新日時
            Call UniCode_Conv(P_SSHIJI_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))


'''Y_SYUKA の出力やめる　2010.09.17
            If POS_UMU Then
                '出荷指示の作成
                If Y_SYUKA_Make_Proc(i) Then
                    GoTo Abort_Tran
                End If
            End If

                                                                                        '出荷予定ＩＤ
            Call UniCode_Conv(P_SSHIJI_K_REC.KO_ID_NO, D_Item_Tbl(i).ID_NO)


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
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "商品化指図ﾃﾞｰﾀ(子)", 0)

        
                            sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                            If sts <> BtNoErr Then
                                Call File_Error(sts, BtOpAbortTransaction, "")
                            End If
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc0
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
                        Call File_Error(sts, BtOpInsert, "商品化指図ﾃﾞｰﾀ(子)")
                        GoTo Abort_Tran
                End Select

            Loop

        End If

    Next i





End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If

'    If Mode = 0 Then       '2007.11.21
    If MSG = 0 Then         '2007.11.21

'2008.05.19        If Text1(ptxSHIJI_NO).text = "" Then
        If NEW_F Then       '2008.05.19
            MsgBox "指図票№：" & Format(SHIJINO, "00000000") & "を作成しました。"  '2008.02.13
        End If
    End If

    Call Input_UnLock
                                        '印刷に対象指図票№を通知
    Taget_Key = Format(SHIJINO, "00000000") '2008.02.13

    Update_Proc = False

    Exit Function

Abort_Tran:

    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If

    Call Input_UnLock

End Function

Private Sub Check1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'
'2010.11.12  ﾊﾟｰﾂﾗﾍﾞﾙ発行(PM0040)の動きに合わせる為、ｲﾍﾞﾝﾄﾌﾟﾛｼｰｼﾞｬを追加

    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    Call Tab_Ctrl(Shift)        '移動

End Sub


Private Sub Combo1_Click(Index As Integer)

Dim sts         As Integer

Dim Sumi_Qty    As Long
Dim Mi_Qty      As Long

Dim TABLCTRL_SW As Integer      '2019.06.11 追加
    
    TABLCTRL_SW = 0
    
    Select Case Index
        Case pcmbSHIMUKE        '仕向け先
'            svJGYOBU = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1)
'            svNAIGAI = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1)
            '2019.06.11 ↑ここでセットすると、下記のIf文が無意味！
            '           コメントにした。
'            If Trim(Text1(ptxHIN_GAI)) <> "" Then
'                If svJGYOBU <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
'                    svNAIGAI <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
'
'                    chenge_F = True
'                End If
'            End If
            
            '2019.06.11 仕向け先の４桁で判定
            If Trim(Text1(ptxHIN_GAI)) <> "" Then
                If svSHIMUKE <> Right(Combo1(pcmbSHIMUKE), 4) Then
                    chenge_F = True
                End If
            End If
            
            svSHIMUKE_CODE = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2)   '2013.08.29

            svSHIMUKE = Right(Combo1(pcmbSHIMUKE), 4)





            If chenge_F Then
                TABLCTRL_SW = 1
                
                If Error_Check_Proc(ptxHIN_GAI, 0, 0) Then   'エラーチェック
                    chenge_F = True
                    Text1(ptxHIN_GAI).SelLength = Len(Text1(ptxHIN_GAI))
                    Text1(ptxHIN_GAI).SelStart = 0
                    Text1(ptxHIN_GAI).SetFocus
                    Exit Sub
                End If

Start_Proc1:    '2015.03.26
                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)


                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                sts = BTRV(BtOpGetGreaterEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Text1(ptxHIN_GAI) = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
                    Case BtErrKeyNotFound, BtErrEOF

                        Text1(ptxHIN_NAME).text = ""
                        Text1(ptxST_LOCATION).text = ""
                        Text1(ptxMI_QTY).text = ""
                        Text1(ptxSUMI_QTY).text = ""

                        
'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                        Check1(pchkL_PAPER).Value = vbUnchecked         '紙
                        Check1(pchkL_PLASTIC).Value = vbUnchecked       'プラ
                        Check1(pchkL_LABEL).Value = vbUnchecked         '適用機種ラベル
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                        
                        MsgBox "入力した項目はエラーです。(品番)"
                        Text1(ptxHIN_GAI).SetFocus
                        Exit Sub                        '2019.06.10 高沢
'                        Exit Sub    '2010.11.17
                    Case Else
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                        
                        
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                           GoTo Start_Proc1
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
                        
                        
                        
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Unload Me

                End Select



                Text1(ptxHIN_NAME).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)

'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                If StrConv(ITEMREC.L_PAPER, vbUnicode) = L_PAPER_ON Then        '紙
                    Check1(pchkL_PAPER).Value = vbChecked
                Else
                    Check1(pchkL_PAPER).Value = vbUnchecked
                End If
    
                If StrConv(ITEMREC.L_PLASTIC, vbUnicode) = L_PLASTIC_ON Then    'プラ
                    Check1(pchkL_PLASTIC).Value = vbChecked
                Else
                    Check1(pchkL_PLASTIC).Value = vbUnchecked
                End If
    
                If StrConv(ITEMREC.L_LABEL, vbUnicode) = L_LABEL_ON Then        '適用機種ラベル
                    Check1(pchkL_LABEL).Value = vbChecked
                Else
                    Check1(pchkL_LABEL).Value = vbUnchecked
                End If
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Text1(ptxST_LOCATION).text = ""
                Else
                    Text1(ptxST_LOCATION).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If

                If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                        StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                        StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then

                    Unload Me
                End If

                Text1(ptxMI_QTY).text = Format(Mi_Qty, "#0")
                Text1(ptxSUMI_QTY).text = Format(Sumi_Qty, "#0")



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
                            Case Else
                                
                                
                                
                                
                                
                                Call File_Error(sts, BtOpGetEqual, "構成マスタ")
                                Unload Me
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
                            Case Else
                                
                                
                                
                                
                                
                                
                                Call File_Error(sts, BtOpGetEqual, "構成マスタ")
                                Unload Me
                        End Select

                    End If
                End If

                chenge_F = False
                
                DoEvents
                
'                Combo1(pcmbSHIMUKE).SetFocus
                
                Text1(ptxHIN_GAI).SelLength = Len(Text1(ptxHIN_GAI))
                Text1(ptxHIN_GAI).SelStart = 0
                Text1(ptxHIN_GAI).SetFocus
                Exit Sub
            End If
        
        svJGYOBU = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1)
        svNAIGAI = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1)
        
        svSHIMUKE = Right(Combo1(pcmbSHIMUKE), 4)
        
        '2019.06.10 高沢
        If TABLCTRL_SW = 1 Then
            Text1(ptxHIN_GAI).SetFocus
        Else
            Call Tab_Ctrl(0)        '移動
        End If
        
        Exit Sub
    Case Else
    
    End Select

End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'
'
'Dim sts         As Integer
'
'Dim Sumi_Qty    As Long
'Dim Mi_Qty      As Long
'
'
'
'
'    If KeyCode <> vbKeyReturn Then
'        Exit Sub
'    End If
'
'    Select Case Index
'        Case pcmbSHIMUKE        '仕向け先
'
'            If svJGYOBU <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
'                svNAIGAI <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
'
'                chenge_F = True
'
'
'            End If
'
'
'
'            If chenge_F Then
'
'Start_Proc1:    '2015.03.13
'                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
'                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)
'
'
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                Select Case sts
'                    Case BtNoErr
'                    Case BtErrKeyNotFound
'
'                        Text1(ptxHIN_NAME).text = ""
'                        Text1(ptxST_LOCATION).text = ""
'                        Text1(ptxMI_QTY).text = ""
'                        Text1(ptxSUMI_QTY).text = ""
'
''2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                        Check1(pchkL_PAPER).Value = vbUnchecked         '紙
'                        Check1(pchkL_PLASTIC).Value = vbUnchecked       'プラ
'                        Check1(pchkL_LABEL).Value = vbUnchecked         '適用機種ラベル
''2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'
'                        MsgBox "入力した項目はエラーです。(品番)"
'                        Text1(ptxHIN_GAI).SetFocus
'
'                        Exit Sub    '2010.11.17
'                    Case Else
'
'                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
'                        If sts > 3000 Or sts = 3 Then
'
'
'                            Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
'                            '>>>>>>>>>>>>>  2015.04.24
'                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
'                            'If sts Then
'                            '    Call File_Error(sts, BtOpReset, "")
'                            'End If
'
'                            'Call File_Open_Proc
'                            Do
'                                If Not File_Open_Proc() Then
'                                    Exit Do
'                                End If
'                            Loop
'                            '>>>>>>>>>>>>>  2015.04.24
'
'
'                            GoTo Start_Proc1
'                        End If
'                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.13
'
'
'
'                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
'                        Unload Me
'
'                End Select
'
'
'
'                Text1(ptxHIN_NAME).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
'
''2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                If StrConv(ITEMREC.L_PAPER, vbUnicode) = L_PAPER_ON Then        '紙
'                    Check1(pchkL_PAPER).Value = vbChecked
'                Else
'                    Check1(pchkL_PAPER).Value = vbUnchecked
'                End If
'
'                If StrConv(ITEMREC.L_PLASTIC, vbUnicode) = L_PLASTIC_ON Then    'プラ
'                    Check1(pchkL_PLASTIC).Value = vbChecked
'                Else
'                    Check1(pchkL_PLASTIC).Value = vbUnchecked
'                End If
'
'                If StrConv(ITEMREC.L_LABEL, vbUnicode) = L_LABEL_ON Then        '適用機種ラベル
'                    Check1(pchkL_LABEL).Value = vbChecked
'                Else
'                    Check1(pchkL_LABEL).Value = vbUnchecked
'                End If
''2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'
'
'                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
'                    Text1(ptxST_LOCATION).text = ""
'                Else
'                    Text1(ptxST_LOCATION).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
'                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
'                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
'                                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
'                End If
'
'                If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
'                                                        StrConv(ITEMREC.NAIGAI, vbUnicode), _
'                                                        StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
'
'                    Unload Me
'                End If
'
'                Text1(ptxMI_QTY).text = Format(Mi_Qty, "#0")
'                Text1(ptxSUMI_QTY).text = Format(Sumi_Qty, "#0")
'
'
'
'                If Trim(Text1(ptxSHIJI_NO).text) = "" Then
'                    If StrConv(P_COMPO_O_REC.SHIMUKE_CODE, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) And _
'                        StrConv(P_COMPO_O_REC.JGYOBU, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) And _
'                        StrConv(P_COMPO_O_REC.NAIGAI, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) And _
'                        Trim(StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode)) = Trim(Text1(ptxHIN_GAI).text) Then
'                    Else
'
'
'
'                        sts = P_COMPO_Disp_Proc()
'                        Select Case sts
'                            Case BtNoErr
'                            Case BtErrKeyNotFound
'                            Case Else
'
'
'
'                                Call File_Error(sts, BtOpGetEqual, "構成マスタ")
'                                Unload Me
'                        End Select
'                    End If
'                Else
'                    If StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) And _
'                        StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) And _
'                        StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) And _
'                        Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)) = Trim(Text1(ptxHIN_GAI).text) Then
'                    Else
'
'
'                        sts = P_COMPO_Disp_Proc()
'                        Select Case sts
'                            Case BtNoErr
'                            Case BtErrKeyNotFound
'                            Case Else
'
'
'
'                                Call File_Error(sts, BtOpGetEqual, "構成マスタ")
'                                Unload Me
'                        End Select
'
'                    End If
'                End If
'                Text1(ptxSHIJI_QTY).SetFocus
'
'                chenge_F = False
'                Exit Sub
'            End If
'
'            '2019.06.10 下記を追加
'            Text1(ptxHIN_GAI).SetFocus
'
'            Exit Sub
'
'        Case pcmbUKEHARAI       '手配先
'            Text1(ptxUKEHARAI_CODE).text = Trim(Right(Combo1(pcmbUKEHARAI).text, 5))
'        Case pcmbS_TANTO        '収単／担当者
'
'                                '同梱／構成　種別
'        Case pcmbD_SYUBETSU01, pcmbD_SYUBETSU02, pcmbD_SYUBETSU03, pcmbD_SYUBETSU04, pcmbD_SYUBETSU05, pcmbD_SYUBETSU06
'
'            D_Item_Tbl(Index - pcmbD_SYUBETSU01).SYUBETSU = Right(Combo1(Index).text, 2)
'    End Select
'
    Call Tab_Ctrl(Shift)        '移動

End Sub


Private Sub Combo1_LostFocus(Index As Integer)
                                               '2019.06.10　全て、コメントにしてみた。
'                                               '2019.06.11　全て、復帰してみた。
Dim i   As Integer  '2013.08.29

    Select Case Index
        Case pcmbSHIMUKE        '仕向け先

'            If svJGYOBU <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
'                svNAIGAI <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
'
'                chenge_F = True
'
'            End If
            '2019.06.11 ４桁で判定に変更
            If svSHIMUKE <> Right(Combo1(pcmbSHIMUKE), 4) Then
                chenge_F = True
            End If


'2013.08.29 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'2013.11.21            If svSHIMUKE_CODE <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Then
                For i = 0 To UBound(SHIMUKE_CHK_TBL)

                    If SHIMUKE_CHK_TBL(i) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Then
                        Combo2(0).ListIndex = 0
                        Check1(pchkPRI_GAISOU).Value = vbUnchecked
                        Exit For
                    End If

                Next i
'2013.11.21            End If
'2013.08.29 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>



        Case pcmbUKEHARAI       '手配先
            Text1(ptxUKEHARAI_CODE).text = Trim(Right(Combo1(pcmbUKEHARAI).text, 5))
        Case pcmbS_TANTO        '収単／担当者

                                '同梱／構成　種別
        Case pcmbD_SYUBETSU01, pcmbD_SYUBETSU02, pcmbD_SYUBETSU03, pcmbD_SYUBETSU04, pcmbD_SYUBETSU05, pcmbD_SYUBETSU06

            D_Item_Tbl(Index - pcmbD_SYUBETSU01).SYUBETSU = Right(Combo1(Index).text, 2)
    End Select

End Sub


Private Sub Combo2_Click(Index As Integer)
    
    If Right(Combo2(0).text, 1) = " " Or Right(Combo2(0).text, 1) = "2" Then
        Check1(pchkL_PAPER).Enabled = False
        Check1(pchkL_PLASTIC).Enabled = False
    Else
        Check1(pchkL_PAPER).Enabled = True
        Check1(pchkL_PLASTIC).Enabled = True
    End If


End Sub

Private Sub Combo2_GotFocus(Index As Integer)

    '2011.11.19
    If Right(Combo2(0).text, 1) = "2" Then
        Check1(pchkL_PAPER).Enabled = False
        Check1(pchkL_PLASTIC).Enabled = False
    Else
        Check1(pchkL_PAPER).Enabled = True
        Check1(pchkL_PLASTIC).Enabled = True
    End If
    '2011.11.19


End Sub

Private Sub Combo2_LostFocus(Index As Integer)

    If Right(Combo2(0).text, 1) = " " Or Right(Combo2(0).text, 1) = "2" Then
        Check1(pchkL_PAPER).Enabled = False
        Check1(pchkL_PLASTIC).Enabled = False
    Else
        Check1(pchkL_PAPER).Enabled = True
        Check1(pchkL_PLASTIC).Enabled = True
    End If

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans             As Integer
Dim i               As Integer

Dim rpt             As New PI00010F1
Dim f               As New PI000103

Dim rpt2            As New PI00010F2


Dim com             As Integer
Dim sts             As Integer


Dim Parts_F         As Integer
Dim Gaisou_F        As Integer
Dim Kishu_F         As Integer

Dim objAccess       As Access.Application
Dim strAccessPath   As String

Dim GAISOU_QTY          As Long
Dim GAISOU_SHIJI_QYU    As Long

'Dim L_print_Flg     As Boolean

Dim FileNo      As Long         '2008.05.30


'=============================== 2007/03/19 =====
Dim wk_SHIJI_NORMAL As Integer
Dim wk_SHIJI_SPOT   As Integer
Dim wk_SHIJI_KEPPIN As Integer
'================================================


'=============================== 2011.02.16 =====
Dim Parts       As String   '品番
Dim ID          As Long     '指示№

Dim PartsLabel  As Integer  '品番ﾗﾍﾞﾙ 0:なし 以外：枚数
Dim KisyuLabel  As Integer  '機種ﾗﾍﾞﾙ 0:なし
Dim JanLabel    As Integer  'JANﾗﾍﾞﾙ 0:なし
Dim GLabel      As Integer  '外装ﾗﾍﾞﾙ 0:なし
Dim ItemLabel   As Integer  'ｱｲﾃﾑﾗﾍﾞﾙ枚数

Dim OrderNo     As String
Dim ItemNo      As String

Dim Pri_Date    As String

Dim L_QTY       As Long         '2008.10.03

'=============================== 2011.02.16 =====

Dim KISHU1      As String       '2012.09.03
Dim KISHU2      As String       '2012.09.03


Dim LABEL_CHECK_F   As Boolean  '2013.11.05


Dim GYO_SU      As Long         '2016.01.05


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  エラーチェック 2016.01.29
Dim GEN_NG_F        As Integer      '原産国空白
Dim GEN_AT_GAI_F    As Integer      '原産国注意(海外)
Dim GEN_AT_PLU_F    As Integer      '原産国注意(複数)
Dim TANKA_SP_F      As Integer      '単価空白(単価２,単価３)
Dim KISHU_NG_F      As Integer      '機種空白
Dim KAISYA_NG_F     As Integer      '会社／事業部空白

Dim yn              As Integer
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  エラーチェック 2016.01.29


    For i = 0 To 97




        Select Case i
            Case ptxHIN_GAI, ptxK_HIN_GAI01, ptxK_HIN_GAI02, ptxK_HIN_GAI03, ptxK_HIN_GAI04, ptxK_HIN_GAI05, _
                    ptxG_HIN_GAI01, ptxG_HIN_GAI02, ptxG_HIN_GAI03, _
                    ptxD_HIN_GAI01, ptxD_HIN_GAI02, ptxD_HIN_GAI03, ptxD_HIN_GAI04, ptxD_HIN_GAI05, ptxD_HIN_GAI06 _


                Text1(i).text = RTrim(StrConv(Text1(i).text, vbUpperCase))

        End Select

    Next i



        GYO_SU = SendMessage(RichTextBox1(prchBIKOU).hwnd, EM_GETLINECOUNT, 0&, 0&)  '2016.01.05



    Select Case Index
        Case P_CMD_Upd        '更新


            For i = ptxSHIJI_NO To ptxD_BIKOU06

                If chenge_F Then        '2008.07.30

                    If Error_Check_Proc(i, 0, 0, 9) Then 'エラーチェック
                        Exit Sub
                    End If
                Else
                    If Error_Check_Proc(i, 0, 1, 9) Then 'エラーチェック
                        Exit Sub
                    End If
                End If

            Next i

'>>>>>> 2016.02.10
            If LenB(StrConv(RTrim(RichTextBox1(prchBIKOU).text), vbFromUnicode)) > 120 Then
                yn = MsgBox("備考が桁数オーバーしています(最大120文字)、オーバした文字は切り捨てられます。", vbYesNo, "確認入力")
                If yn = vbNo Then
                    RichTextBox1(prchBIKOU).SetFocus
                    Exit Sub
                End If
            End If
'>>>>>> 2016.02.10


            If GYO_SU > 5 Then                                                      '2016.01.05
                MsgBox "備考最大印字行数は５行です。内容を確認して下さい。"         '2016.01.05
                RichTextBox1(prchBIKOU).SetFocus                                    '2016.01.05
                Exit Sub                                                            '2016.01.05
            End If                                                                  '2016.01.05


            Beep
            ans = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Update_Proc(0, 0) Then       '引数追加   2007.11.21
                    Unload Me
                End If

                If Init_Proc() Then
                    Unload Me
                End If
                chenge_F = False                '2019.06.20 クリアしないと品番チェック実行してしまう！
                DoEvents
                Text1(ptxSHIJI_NO).SetFocus

                PI000104_OLD_HIN_GAI = ""       '2019.04.18
                
                
                
            Else
                chenge_F = False                '2019.06.20 クリアしないと品番チェック実行してしまう！
                DoEvents
                Text1(ptxHAKKO_DT).SetFocus
            End If


'        Case P_CMD_DEL                      '削除
        Case cmdMUPDATE                     'ﾏｽﾀ更新

            
            
            
            For i = ptxSHIJI_NO To ptxD_BIKOU06

                If chenge_F Then        '2008.07.30

                    If Error_Check_Proc(i, 1, 0, 9) Then 'エラーチェック
                        Exit Sub
                    End If
                Else
                    If Error_Check_Proc(i, 1, 1, 9) Then 'エラーチェック
                        Exit Sub
                    End If
                End If

            Next i
'>>>>>> 2016.02.10
            If LenB(StrConv(RTrim(RichTextBox1(prchBIKOU).text), vbFromUnicode)) > 120 Then
                yn = MsgBox("備考が桁数オーバーしています(最大120文字)、オーバした文字は切り捨てられます。", vbYesNo, "確認入力")
                If yn = vbNo Then
                    RichTextBox1(prchBIKOU).SetFocus
                    Exit Sub
                End If
            End If
'>>>>>> 2016.02.10

            If GYO_SU > 5 Then                                                      '2016.01.05
                MsgBox "備考最大印字行数は５行です。内容を確認して下さい。"         '2016.01.05
                RichTextBox1(prchBIKOU).SetFocus                                    '2016.01.05
                Exit Sub                                                            '2016.01.05
            End If                                                                  '2016.01.05


            Beep
            ans = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Update_Proc(1, 1) Then       '引数追加   2007.11.21
                    Unload Me
                End If

'                If Init_Proc() Then
'                    Unload Me
'                End If

                Call UniCode_Conv(ITEMREC.JGYOBU, "")
                Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, "")
                Text1(ptxSHIJI_NO) = ""

                If Text1(ptxSHIJI_NO).Locked Then
                    Text1(ptxHAKKO_DT).SetFocus
                Else
                    Text1(ptxSHIJI_NO).SetFocus
                End If


                Text1(ptxHIN_GAI).Locked = False            '2019.03.18
    '            Text1(ptxHIN_GAI).Enabled = True           '2019.03.18
                Text1(ptxHIN_GAI).BackColor = &H80000005     '2019.03.18

                PI000104_OLD_HIN_GAI = ""       '2019.04.18

                chenge_F = False                '2019.06.20 クリアしないと品番チェック実行してしまう！
            Else
                chenge_F = False                '2019.06.20 クリアしないと品番チェック実行してしまう！
                Text1(ptxHAKKO_DT).SetFocus
            End If


        Case P_CMD_DSP                      '検索/表示
        Case cmdNext                        '構成部品画面へ

            Doukon_Start = 1
            PI000102.Show vbModal           '部品詳細フォーム表示
            If G_SCREEN_FLG = SYS_ERR Then
                Unload Me
            End If

            'ﾃｰﾌﾞﾙより構成／同梱を表示
            If Tbl_To_Disp_Proc() Then
                Unload Me
            End If

            chenge_F = False                '2019.06.20 クリアしないと品番チェック実行してしまう！

        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        
        
        
        Case P_CMD_PRT                      '印刷


            For i = ptxSHIJI_NO To ptxD_BIKOU06


                If chenge_F Then        '2008.07.30

                    If Error_Check_Proc(i, 0, 0, 9) Then 'エラーチェック
                        Exit Sub
                    End If
                Else
                    If Error_Check_Proc(i, 0, 1, 9) Then 'エラーチェック
                        Exit Sub
                    End If
                End If

            Next i
'>>>>>> 2016.02.10
            If LenB(StrConv(RTrim(RichTextBox1(prchBIKOU).text), vbFromUnicode)) > 120 Then
                yn = MsgBox("備考が桁数オーバーしています(最大120文字)、オーバした文字は切り捨てられます。", vbYesNo, "確認入力")
                If yn = vbNo Then
                    RichTextBox1(prchBIKOU).SetFocus
                    Exit Sub
                End If
            End If
'>>>>>> 2016.02.10

            If GYO_SU > 5 Then                                                      '2016.01.05
                MsgBox "備考最大印字行数は５行です。内容を確認して下さい。"         '2016.01.05
                RichTextBox1(prchBIKOU).SetFocus                                    '2016.01.05
                Exit Sub                                                            '2016.01.05
            End If                                                                  '2016.01.05



Debug.Print Combo1(0).text


            Beep
            
            
            
            
'>>>>>>>>   2016.02.10 ﾁｪｯｸ位置変更


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
                If GENSANKOKU_FLG <> "1" Then
                    GEN_NG_F = 9
                Else
                    '原産国、空白か？
                    If Trim(txGensankoku.text) = "" Then
                        GEN_NG_F = 1
                    Else
                    End If
                End If
            End If

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>海外供給区分チェック>>>>>>>>>>>>>>>>>>>>>>>>>>
            If GAI_BUHIN_CHECK Then
                If Trim(lblGAI_BUHIN.Caption) = "1" Or _
                   Trim(lblGAI_BUHIN.Caption) = "2" Or _
                    Trim(lblGAI_BUHIN.Caption) = "3" Then
                    GEN_AT_GAI_F = 1
                End If
            End If
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>原産国海外向けチェック＆原産国複数チェック>>>>
            If lstGensankoku.ListCount < 1 Then
                GEN_AT_PLU_F = 0
            Else
                GEN_AT_PLU_F = lstGensankoku.ListCount
            End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>単価チェック>>>>>>>>>>>>>>>>>>>>>>>>>>
            If TANKA_SPACE_F = "1" Then
                If Not IsNumeric(lblL_URIKIN2) Or _
                     Not IsNumeric(lblL_URIKIN3) Then
                    TANKA_SP_F = 1
                End If
            End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>代表機種チェック>>>>>>>>>>>>>>>>>>>>>>>>>>
            
            If KISHU_CHECK Then
                KISHU1 = ""
                KISHU2 = ""
                
                For i = 1 To Len(Trim(lblKISHU1.Caption))
                    If Mid(StrConv(lblKISHU1.Caption, vbUnicode), i, 1) <= " " Then
                    Else
                        KISHU1 = KISHU1 & Mid(lblKISHU1.Caption, i, 1)
                    End If
                Next i
                
                For i = 1 To Len(Trim(lblKISHU2.Caption))
                    If Mid(Len(lblKISHU2.Caption), i, 1) <= " " Then
                    Else
                        KISHU2 = KISHU2 & Mid(lblKISHU1.Caption, i, 1)
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
                    If Trim(lblL_KAISHA.Caption) = "" Or Trim(lblL_JGYOBU.Caption) = "" Then
                        KAISYA_NG_F = 1
                    End If
                End If
            End If

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>メッセージ作成>>>>>>>>>>>>>>>>>>>>>>>>>>
            
            If Right(Combo2(0).text, 1) <> " " Then      '2019.03.07
            
                ans = Mesg_Set_Proc(GEN_NG_F, GEN_AT_GAI_F, GEN_AT_PLU_F, TANKA_SP_F, KISHU_NG_F, KAISYA_NG_F, KISHU1, KISHU2)
                If ans = vbCancel Then
                    GoTo Next_Step
                End If

            End If                                      '2019.03.07
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  エラーチェック 2016.01.29




'>>>>>>>>   2016.02.10 ﾁｪｯｸ位置変更
            
            
            
            
 '2016.02.10           ans = MsgBox("印刷／更新しますか？", vbYesNo + vbQuestion, "確認入力")
            ans = vbYes '2016.02.10
            If ans = vbYes Then
                
                If Update_Proc(0, 1) Then       '引数追加   2007.11.21
                    Unload Me
                End If


                Taget_Key = Text1(ptxSHIJI_NO).text

                If Check1(pchkPRI_SHIJI).Value = vbChecked Then

                    '2008.06.26 ↓
                    On Error Resume Next
                    Set objAccess = GetObject(, "Access.Application")
                    If Err().Number <> 0 Then
                        On Error GoTo 0
                    Else
                        On Error GoTo 0
                    '2008.06.26　↑


                        LABEL_CHECK_F = False '2013.11.05

                        '>>>>>>>>>> 2013.11.12
                        If Trim(Right(Combo2(0).text, 1)) <> "" Or _
                            Check1(pchkPRI_GAISOU).Value = vbChecked Then
                        
                        '>>>>>>>>>> 2013.11.12



                            '↓2008.05.30
                            Do
                                
                                
                                
                                On Error Resume Next
    
    
                                FileNo = FreeFile
    
                                Open LabelPrint_F For Input As FileNo
    
                                Select Case Err.Number
                                    Case 0
    
                                        Close #FileNo
    
                                        ans = MsgBox("ラベル発行中です", vbAbortRetryIgnore + vbDefaultButton3 + vbQuestion, "確認入力")
    
                                        Select Case ans
    
                                            Case vbAbort    '中止
    
                                                Exit Sub
    
                                            Case vbIgnore   '無視
    
                                                LABEL_CHECK_F = True    '2013.11.05
    
                                                Exit Do
    
                                        End Select
    
    
    
    
                                    Case 53
                                        Exit Do
    
    
                                    Case Else
    
    
                                        Exit Sub
    
    
                                End Select
    
                                On Error GoTo 0
    
                            Loop
    
                            'Open LabelPrint_F For Output As FileNo         2013.11.05
                            'Close #FileNo                                  2013.11.05
                            '↑2008.05.30
                        End If                                  '2013.11.12
                    End If              '2008.06.26









                    If CDbl(Text1(ptxSHIJI_QTY).text) <> 0 Then '2008.02.02


'2013.01.08 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'2013.02.19                        If StrConv(ITEMREC.PRT_GENSANKOKU, vbUnicode) = "1" Then
'2013.02.19                            chk_TORI_GENSANKOKU = lblGensankoku(1).Caption
'2013.02.19                        ElseIf StrConv(ITEMREC.PRT_GENSANKOKU, vbUnicode) = "0" Then
'2013.02.19                            chk_TORI_GENSANKOKU = ""
'2013.02.19                        Else
'2013.02.19                                '品目ﾏｽﾀ項目が未設定⇒ini定義値により初期表示
'2013.02.19                            If GENSANKOKU_FLG = "1" Then
'2013.02.19                                chk_TORI_GENSANKOKU = lblGensankoku(1).Caption
'2013.02.19                            Else
'2013.02.19                                chk_TORI_GENSANKOKU = ""
'2013.02.19                            End If
'2013.02.19                        End If

                    
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2016.02.01
'                        '>>>>海外共有区分のチェック　2015.07.23
'                        If GAI_BUHIN_CHK Then
'                            If StrConv(ITEMREC.GAI_BUHIN, vbUnicode) = "1" Or StrConv(ITEMREC.GAI_BUHIN, vbUnicode) = "2" Or StrConv(ITEMREC.GAI_BUHIN, vbUnicode) = "3" Then
'                                MsgBox "原産国注意"
'                            End If
'                        Else
'                        '>>>>海外共有区分のチェック　2015.07.23
'
'                            If GENSANKOKU_MSG_F Then                        '2013.02.19
'                                If Trim(chk_TORI_GENSANKOKU) <> "" Then
'                                        MsgBox "原産国注意"
'                                End If
'                            End If
'                        End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2016.02.01


'2013.01.08 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

                        PRINT_STOP_F = False        '2015.03.26
                        Set rpt = New PI00010F1

                        'レポートを印刷します。（true：印刷ダイアログあり false：なし）
                        
                        If Not PRINT_STOP_F Then    '2015.03.26
                            rpt.PrintReport False
                        End If                      '2015.03.26

                        Set rpt = Nothing


    '                    f.RunReport rpt
    '                    f.Show



''''2011.02.17
''''                        If Check1(pchkPRI_PARTS).Value = vbChecked Or _
''''                            Check1(pchkPRI_GAISOU).Value = vbChecked Then

                
                            If Trim(Right(Combo2(0).text, 1)) <> "" Or _
                                Check1(pchkPRI_GAISOU).Value = vbChecked Then

                            


''''2011.02.17
                            L_print_Flg = True


'>>>>>>>>   2016.02.10 ﾁｪｯｸ位置変更
'
'
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  エラーチェック 2016.01.29
'            GEN_NG_F = 0        '原産国空白
'            GEN_AT_GAI_F = 0    '原産国注意(海外)
'            GEN_AT_PLU_F = 0    '原産国注意(複数)
'            TANKA_SP_F = 0      '単価空白(単価２,単価３)
'            KISHU_NG_F = 0      '機種空白
'            KAISYA_NG_F = 0     '会社／事業部空白
'
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>原産国空白チェック>>>>>>>>>>>>>>>>>>>>>>>>>>
'            For i = 0 To UBound(GENSANKOKU_CHECK_TBL)
'                If Last_JGYOBU = GENSANKOKU_CHECK_TBL(i) Then
'                    Exit For
'                End If
'            Next i
'            If i > UBound(GENSANKOKU_CHECK_TBL) Then
'                GEN_NG_F = 9
'            Else
'                If GENSANKOKU_FLG <> "1" Then
'                    GEN_NG_F = 9
'                Else
'                    '原産国、空白か？
'                    If Trim(txGensankoku.text) = "" Then
'                        GEN_NG_F = 1
'                    Else
'                    End If
'                End If
'            End If
'
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>海外供給区分チェック>>>>>>>>>>>>>>>>>>>>>>>>>>
'            If GAI_BUHIN_CHECK Then
'                If StrConv(ITEMREC.GAI_BUHIN, vbUnicode) = "1" Or _
'                    StrConv(ITEMREC.GAI_BUHIN, vbUnicode) = "2" Or _
'                    StrConv(ITEMREC.GAI_BUHIN, vbUnicode) = "3" Then
'                    GEN_AT_GAI_F = 1
'                End If
'            End If
'
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>原産国海外向けチェック＆原産国複数チェック>>>>
'            If lstGensankoku.ListCount < 1 Then
'                GEN_AT_PLU_F = 0
'            Else
'                GEN_AT_PLU_F = lstGensankoku.ListCount
'            End If
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>単価チェック>>>>>>>>>>>>>>>>>>>>>>>>>>
'            If TANKA_SPACE_F = "1" Then
'                If Not IsNumeric(StrConv(ITEMREC.L_URIKIN2, vbUnicode)) Or _
'                     Not IsNumeric(StrConv(ITEMREC.L_URIKIN2, vbUnicode)) Then
'                    TANKA_SP_F = 1
'                End If
'            End If
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>代表機種チェック>>>>>>>>>>>>>>>>>>>>>>>>>>
'
'            If KISHU_CHECK Then
'                KISHU1 = ""
'                KISHU2 = ""
'
'                For i = 1 To Len(Trim(lblKISHU1.Caption))
'                    If Mid(StrConv(lblKISHU1.Caption, vbUnicode), i, 1) <= " " Then
'                    Else
'                        KISHU1 = KISHU1 & Mid(lblKISHU1.Caption, i, 1)
'                    End If
'                Next i
'
'                For i = 1 To Len(Trim(lblKISHU2.Caption))
'                    If Mid(Len(lblKISHU2.Caption), i, 1) <= " " Then
'                    Else
'                        KISHU2 = KISHU2 & Mid(lblKISHU1.Caption, i, 1)
'                    End If
'                Next i
'
'                If Trim(KISHU1) = "" And Trim(KISHU2) = "" Then
'                    KISHU_NG_F = 1
'                End If
'            End If
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>会社名／事業部名チェック>>>>>>>>>>>>>>>>>>>>>>>>>>
'            If KAISYA_RESTRICT_F Then
'                KAISYA_NG_F = 9
'            Else
'                If KAISYA_CHK_F Then
'                    If Trim(lblL_KAISHA.Caption) = "" Or Trim(lblL_JGYOBU.Caption) = "" Then
'                        KAISYA_NG_F = 1
'                    End If
'                End If
'            End If
'
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>メッセージ作成>>>>>>>>>>>>>>>>>>>>>>>>>>
'            ans = Mesg_Set_Proc(GEN_NG_F, GEN_AT_GAI_F, GEN_AT_PLU_F, TANKA_SP_F, KISHU_NG_F, KAISYA_NG_F, KISHU1, KISHU2)
'            If ans = vbCancel Then
'                GoTo Next_Step
'            End If
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  エラーチェック 2016.01.29




'>>>>>>>>   2016.02.10 ﾁｪｯｸ位置変更












                            If L_URIKIN1 = 0 And L_URIKIN2 = 0 And L_URIKIN3 = 0 Then

                                Beep
'2016.02.01                                ans = MsgBox("単価未設定です。ラベル印刷しますか？", vbYesNo + vbQuestion, "確認入力")
                                ans = vbYes
                                If ans = vbYes Then
                                Else
'                                    L_print_Flg = False            '2013.08.29
                                    
                                    
                                    '>>>>> 2013.11.05
                                    'Text1(ptxHIN_GAI).SetFocus      '2013.08.29
                                    'Exit Sub                        '2013.08.29
                                    GoTo Next_Step
                                    '>>>>> 2013.11.05
                                End If
                            Else
                            End If


                            '会社事業部エラーﾁｪｯｸ有無 2010.07.20
                            If KAISYA_CHK_F Then

                                If Trim(lblL_KAISHA.Caption) = "" Or Trim(lblL_JGYOBU.Caption) = "" Then
'2016.0.01                                    ans = MsgBox("会社名/事業部 が指定されていません。(ＯＫ＝発行、ｷｬﾝｾﾙ=発行しない)", vbOKCancel + vbQuestion + vbDefaultButton2, "確認入力") '2013.08.29
'                                    ans = MsgBox("会社/事業部が指定されていません。ラベル印刷しますか？", vbYesNo + vbQuestion + vbDefaultButton2, "確認入力")                                                      '2013.08.29
                                    ans = vbYes
                                    'If ans = vbCancel Then     '2013.08.29
                                    If ans = vbNo Then          '2013.08.29
' 2013.01.05 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                        L_print_Flg = False
                                        
                                        
                                        '>>>>> 2013.11.05
                                        'Text1(ptxHIN_GAI).SetFocus
                                        'Exit Sub    '2013.08.29
                                        GoTo Next_Step
                                        '>>>>> 2013.11.05
                                    Else
'2013.08.27                                        sts = Shell(App.Path & "\PM00040.exe " & Right(Combo1(pcmbSHIMUKE).text, 2) & Trim(Text1(ptxHIN_GAI).text), vbNormalFocus)
                                    End If

'2013.08.27                                    Exit Sub
' 2013.01.05 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                End If

                            End If
                            '会社事業部エラーﾁｪｯｸ有無 2010.07.20



                            '2009.03.28
                            For i = 0 To UBound(GENSANKOKU_CHECK_TBL)


                                If Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) = GENSANKOKU_CHECK_TBL(i) Then
                                    Exit For
                                End If

                            Next i



                            '2009.03.28
                            If i > UBound(GENSANKOKU_CHECK_TBL) Then
                            Else


                                If Trim(txGensankoku.text) = "" Then

                                    'ans = MsgBox("原産国が空白です。(ＯＫ＝印刷中止、ｷｬﾝｾﾙ=継続)", vbOKCancel + vbQuestion, "確認入力")    '2013.08.29
'2016.02.01                                    ans = MsgBox("原産国が空白です。ラベル印刷しますか？", vbYesNo + vbQuestion + vbDefaultButton2, "確認入力")                    '2013.08.29
                                    ans = vbYes
                                    'If ans = vbCancel Then '2013.08.29
                                    If ans = vbYes Then     '2013.08.29
' 2013.01.05 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'2013.08.27                                        sts = Shell(App.Path & "\PM00040.exe " & Right(Combo1(pcmbSHIMUKE).text, 2) & Trim(Text1(ptxHIN_GAI).text), vbNormalFocus)
                                    Else
'                                        L_print_Flg = False
                                        '>>>>> 2013.11.05
                                        Text1(ptxHIN_GAI).SetFocus
                                        Exit Sub    '2013.08.29
                                        GoTo Next_Step
                                        '>>>>> 2013.11.05
                                    End If

'2013.08.27                                    Exit Sub
' 2013.01.05 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                End If

                            End If



                            '2012.09.03     代表機種ﾁｪｯｸ        2012.10.26 itemrec.L_KISHU1 -- > lblKISHU1,itemrec.L_KISHU2 -- > lblKISHU2
                            If KISHU_CHECK Then
                                KISHU1 = ""
                                KISHU2 = ""
                                
                                For i = 1 To Len(Trim(lblKISHU1.Caption))
                                    If Mid(StrConv(lblKISHU1.Caption, vbUnicode), i, 1) <= " " Then
                                    Else
                                        KISHU1 = KISHU1 & Mid(lblKISHU1.Caption, i, 1)
                                    End If
                                Next i
                                
                                For i = 1 To Len(Trim(lblKISHU2.Caption))
                                    If Mid(Len(lblKISHU2.Caption), i, 1) <= " " Then
                                    Else
                                        KISHU2 = KISHU2 & Mid(lblKISHU1.Caption, i, 1)
                                    End If
                                Next i
                                
                                If Trim(KISHU1) = "" And Trim(KISHU2) = "" Then
                                    'ans = MsgBox("代表機種が空白です。(ＯＫ＝印刷中止、ｷｬﾝｾﾙ=継続)", vbOKCancel + vbQuestion, "確認入力")  '2013.08.29
'2016.02.01                                    ans = MsgBox("代表機種が空白です。ラベル印刷しますか?", vbYesNo + vbQuestion + vbDefaultButton2, "確認入力")                   '2013.08.29
                                    ans = vbYes
                                    If ans = vbYes Then
                                    Else
                                        L_print_Flg = False
                                    End If
                                End If
                            End If
                            
                            '2012.09.03     代表機種ﾁｪｯｸ



                            If L_print_Flg Then


                                On Error Resume Next
                                Set objAccess = GetObject(, "Access.Application")
                                If Err().Number <> 0 Then
                                    MsgBox "この端末では商品ラベル発行は行えません。"
            '                        MsgBox "GetObject(Access.Application)" & Err().Number & " " & Err().Description
                                Else
            '                        MsgBox Err.Number


                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>   他端末の状況再チェック  2013.11.05

                                    
                                    
                                    If Not LABEL_CHECK_F Then   '2013.11.05
                                        Do
                                            On Error Resume Next
                
                                            FileNo = FreeFile
                
                                            Open LabelPrint_F For Input As FileNo
                
                                            Select Case Err.Number
                                                Case 0
                
                                                    Close #FileNo
                
                                                    ans = MsgBox("ラベル発行中です", vbAbortRetryIgnore + vbDefaultButton3 + vbQuestion, "確認入力")
                
                                                    Select Case ans
                
                                                        Case vbAbort    '中止
                
                                                            Exit Sub
                
                                                        Case vbIgnore   '無視
                
                                                            Exit Do
                
                                                    End Select
                
                
                
                
                                                Case 53
                                                    Exit Do
                
                
                                                Case Else
                
                
                                                    Exit Sub
                
                
                                            End Select
                
                                            On Error GoTo 0
                
                                        Loop
                                    End If                  '2013.11.05
                                    
                                    
                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>   他端末の状況再チェック  2013.11.05
                                    
                                    Open LabelPrint_F For Output As FileNo          '2013.11.05
                                    Close #FileNo                                   '2013.11.05
                                    
                                    
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


Start_Proc1:        '2015.03.26

                                        sts = BTRV(com, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
                                        Select Case sts
                                            Case BtNoErr

Start_Proc2:        '2015.03.26
                                                
                                                sts = BTRV(BtOpDelete, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)


                                                Select Case sts

                                                    Case BtNoErr
                                                    Case Else
                                                        
                                                        
                                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                                        If sts > 3000 Or sts = 3 Then
                                    
                                        
                                                            Call File_Error(sts, BtOpGetEqual, "ﾗﾍﾞﾙ用品目ﾏｽﾀ", 0)
                                
                                        
                                                            '>>>>>>>>>>>>>  2015.04.24
                                                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                                            'If sts Then
                                                            '    Call File_Error(sts, BtOpReset, "")
                                                            'End If
                                                        
                                                            'Call File_Open_Proc
                                                            Do
                                                                If Not File_Open_Proc() Then
                                                                    Exit Do
                                                                End If
                                                            Loop
                                                            '>>>>>>>>>>>>>  2015.04.24
                                            
                                                        
                                                            GoTo Start_Proc2
                                                        End If
                                                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                                        
                                                        
                                                        Call File_Error(sts, com, "ﾗﾍﾞﾙ用品目ﾏｽﾀ")
                                                        Exit Sub
                                                End Select

                                            Case BtErrEOF
                                                Exit Do
                                            Case Else
                                                
                                                
                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                                If sts > 3000 Or sts = 3 Then
                            
                                
                                                    Call File_Error(sts, BtOpGetEqual, "ﾗﾍﾞﾙ用品目ﾏｽﾀ", 0)
                        
                                
                                                    '>>>>>>>>>>>>>  2015.04.24
                                                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                                    'If sts Then
                                                    '    Call File_Error(sts, BtOpReset, "")
                                                    'End If
                                                
                                                    'Call File_Open_Proc
                                                    Do
                                                        If Not File_Open_Proc() Then
                                                            Exit Do
                                                        End If
                                                    Loop
                                                    '>>>>>>>>>>>>>  2015.04.24
                                    
                                                
                                                    GoTo Start_Proc1
                                                End If
                                                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                                
                                                
                                                Call File_Error(sts, com, "ﾗﾍﾞﾙ用品目ﾏｽﾀ")
                                                Exit Sub
                                        End Select

                                        com = BtOpGetNext


                                    Loop

Start_Proc3:    '2015.03.26

                                    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                                    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                                    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)


                                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    Select Case sts
                                        Case BtNoErr


                                            Call UniCode_Conv(ITEMREC.L_IRI_QTY, Format(GAISOU_QTY, "00000000"))

''2010.11.15                                            If GENSANKOKU_FLG = "0" Then        '原産国 2008.06.13
''2010.11.15                                                Call UniCode_Conv(ITEMREC.GENSANKOKU, "")
''2010.11.15
''2010.11.15                                            '2010.07.20 ▽
''2010.11.15                                            Else
''2010.11.15                                                Call UniCode_Conv(ITEMREC.GENSANKOKU, lblGensankoku(1).Caption)
''2010.11.15                                            '2010.07.20 △
''2010.11.15                                            End If
                                            
                                            
                                            
                                            
            If StrConv(ITEMREC.PRT_GENSANKOKU, vbUnicode) = "1" Then
                Call UniCode_Conv(ITEMREC.GENSANKOKU, lblGensankoku(1).Caption)
            ElseIf StrConv(ITEMREC.PRT_GENSANKOKU, vbUnicode) = "0" Then
                Call UniCode_Conv(ITEMREC.GENSANKOKU, "")
            
            Else
                    '品目ﾏｽﾀ項目が未設定⇒ini定義値により初期表示
                If GENSANKOKU_FLG = "1" Then
                    Call UniCode_Conv(ITEMREC.GENSANKOKU, lblGensankoku(1).Caption)
                Else
                    Call UniCode_Conv(ITEMREC.GENSANKOKU, "")
                End If
            End If




Debug.Print StrConv(ITEMREC.GENSANKOKU, vbUnicode)

                                            '2008.10.29 棚番(1)に標準棚番をセット
                                            Call UniCode_Conv(ITEMREC.L_TANA1, StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                                                StrConv(ITEMREC.ST_DAN, vbUnicode))

                                            '2008.10.29


Start_Proc4:    '2015.03.26

                                            sts = BTRV(BtOpInsert, L_ITEM_POS, ITEMREC, Len(ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
                                            Select Case sts
                                                Case BtNoErr

                                                    '2007.12.11objAccess.Run "PosPrintLabel", Trim(Text1(ptxHIN_GAI).text), CLng(Text1(ptxSHIJI_QTY).text), Parts_F, Gaisou_F, Kishu_F, GAISOU_QTY, GAISOU_SHIJI_QYU, 0
''2011.02.17 PosPrintLabel-->NewPosPrintLabel
''                                                    objAccess.Run "PosPrintLabel", _
''                                                                    Trim(Text1(ptxHIN_GAI).text), _
''                                                                    CLng(Text1(ptxLabel_QTY).text), _
''                                                                    Parts_F, _
''                                                                    Gaisou_F, _
''                                                                    Kishu_F, _
''                                                                    GAISOU_QTY, _
''                                                                    GAISOU_SHIJI_QYU, _
''                                                                    0




                                                    PartsLabel = 0
                                                    KisyuLabel = 0
                                                    JanLabel = 0
                                                    GLabel = 0
                                                    ItemLabel = 0


                                                    '品目コード
                                                    Parts = Text1(ptxHIN_GAI).text
                                                    'パーツラベル
                                                    If Right(Combo2(0).text, 1) = "0" Then
                                                        PartsLabel = CLng(Text1(ptxLabel_QTY).text)
                                                    End If
                                                    '機種ラベル
                                                    If Right(Combo2(0).text, 1) = "1" Then
                                                        KisyuLabel = CLng(Text1(ptxLabel_QTY).text)
                                                    End If
                                                    'Janラベル
                                                    If Right(Combo2(0).text, 1) = "2" Then
                                                        JanLabel = CLng(Text1(ptxLabel_QTY).text)
                                                    End If

                                                    '外装ﾗﾍﾞﾙ
                                                    If Check1(pchkPRI_GAISOU).Value = vbChecked Then
                                                        GLabel = CLng(Text1(ptxG_SHIJI_QTY01).text)
                                                    End If
                                                    'ID
                                                    ID = 0
                                                    'アイテムラベル
                                                    ItemLabel = 0
                                                    'オーダー№
                                                    OrderNo = ""
                                                    'アイテム№
                                                    ItemNo = ""
                                                    '印刷日付
                                                    Pri_Date = Format(Now, "yyyy/mm/dd")
                                                    '数量
                                                    L_QTY = 1
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



                                                Case Else
                                                    
                                                    
                                                    
                                                    
                                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                                    If sts > 3000 Or sts = 3 Then
                                
                                    
                                                        Call File_Error(sts, BtOpGetEqual, "ﾗﾍﾞﾙ用品目マスタ", 0)
                            
                                    
                            
                                                        '>>>>>>>>>>>>>  2015.04.24
                                                        'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                                        'If sts Then
                                                        '    Call File_Error(sts, BtOpReset, "")
                                                        'End If
                                                    
                                                        'Call File_Open_Proc
                                        
                                                        Do
                                                            If Not File_Open_Proc() Then
                                                                Exit Do
                                                            End If
                                                        Loop
                                                        '>>>>>>>>>>>>>  2015.04.24
                                                    
                                                        GoTo Start_Proc4
                                                    End If
                                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                                    
                                                    
                                                    Call File_Error(sts, BtOpInsert, "ﾗﾍﾞﾙ用品目マスタ")
                                                    Exit Sub


                                            End Select

                                        Case BtErrKeyNotFound

                                        Case Else
                                            
                                            
                                            
                                            
                                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                            If sts > 3000 Or sts = 3 Then
                        
                            
                                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                    
                            
                                                '>>>>>>>>>>>>>  2015.04.24
                                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                                'If sts Then
                                                '    Call File_Error(sts, BtOpReset, "")
                                                'End If
                                            
                                                'Call File_Open_Proc
                                                    
                                                Do
                                                    If Not File_Open_Proc() Then
                                                        Exit Do
                                                    End If
                                                Loop
                                                '>>>>>>>>>>>>>  2015.04.24
                                            
                                                GoTo Start_Proc3
                                            End If
                                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                            
                                            
                                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                            Exit Sub

                                    End Select





                                    Set objAccess = Nothing
                                End If



                            End If
                        End If





                    Else                                                '2008.02.02
                        Taget_SHIMUKE_CODE_KEY = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2)
                        Taget_JGYOBU_key = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1)
                        Taget_NAIGAI_key = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1)
                        Taget_Hin_key = Trim(Text1(ptxHIN_GAI).text)    '2008.02.02
                        Set rpt2 = New PI00010F2                        '2008.02.02
                        'レポートを印刷します。（true：印刷ダイアログあり false：なし）
                        rpt2.PrintReport False                          '2008.02.02
                        Set rpt2 = Nothing                              '2008.02.02
                    End If                                              '2008.02.02

                End If

                'ﾗﾍﾞﾙｼｽﾃﾑ印刷要求




Next_Step:



                '=============================== 2007/03/19 =====
                wk_SHIJI_NORMAL = Option1(poptSHIJI_NORMAL).Value
                wk_SHIJI_SPOT = Option1(poptSHIJI_SPOT).Value
                wk_SHIJI_KEPPIN = Option1(poptSHIJI_KEPPIN).Value
                '================================================

                If Init_Proc() Then
                    Unload Me
                End If
                                '2019.06.18 印刷する時は、chenge_F＝Falseにした。２行追加
                DoEvents
                chenge_F = False

                '=============================== 2007/03/19 =====
                Option1(poptSHIJI_NORMAL).Value = wk_SHIJI_NORMAL
                Option1(poptSHIJI_SPOT).Value = wk_SHIJI_SPOT
                Option1(poptSHIJI_KEPPIN).Value = wk_SHIJI_KEPPIN

'                Text1(ptxSHIJI_NO).SetFocus
                Text1(ptxHIN_GAI).SetFocus
                '================================================


            Else
                                '2019.06.18 印刷する時は、chenge_F＝Falseにした。２行追加
                DoEvents
                chenge_F = False
                
                
                Text1(ptxHAKKO_DT).SetFocus
            End If


        Case 9                              'COPY画面   2019.03.14
            
            PI000104.Show vbModal
        
                    
            If PI000104_CANCEL_F = 1 Then
                Exit Sub
            End If
        
        
        
            If PI000104_Error_F = 1 Then
                Unload Me
            End If
            PI000104_OLD_HIN_GAI = ""
            If Trim(PI000104_HIN_GAI) <> "" Then
                PI000104_OLD_HIN_GAI = Text1(ptxHIN_GAI).text
            End If
                    
            Text1(ptxHIN_GAI).text = PI000104_HIN_GAI
            
                    
            
            
            '2019.05.27 　　　　　　　　　　↓引数の２番目は、１では？
'            If Error_Check_Proc(ptxHIN_GAI, 0, 0) Then
'                Exit Sub
'            End If
            '2019.05.27                     引数を変更してみた。    高沢
            If Error_Check_Proc(ptxHIN_GAI, 1, 0) Then
                Exit Sub
            End If
            
            
            
            Command1(cmdMUPDATE).SetFocus
        
        
            Text1(ptxHIN_GAI).Locked = True             '2019.03.18
            Text1(ptxHIN_GAI).BackColor = &H8000000F    '2019.03.18
'            Text1(ptxHIN_GAI).Enabled = False           '2019.03.18
        
        
        Case cmdCen                         '取り消し
            If Init_Proc() Then
                Unload Me
            End If
            
            
            Text1(ptxHIN_GAI).Locked = False            '2019.03.18
'            Text1(ptxHIN_GAI).Enabled = True           '2019.03.18
            Text1(ptxHIN_GAI).BackColor = &H80000005     '2019.03.18
            
            '2019.06.12 下記２行追加
            DoEvents
            chenge_F = False
            
            
            Text1(ptxSHIJI_NO).SetFocus
        Case P_CMD_End                      '終了
            Unload Me
    
    
        Case 12                                         '2019.03.18
            
            PI000104_OLD_HIN_GAI = ""       '2019.04.18

            
            Text1(ptxHIN_GAI).Locked = False            '2019.03.18
'            Text1(ptxHIN_GAI).Enabled = True           '2019.03.18
            Text1(ptxHIN_GAI).BackColor = &H80000005     '2019.03.18
            
            If Init_Proc() Then
                Unload Me
            End If
            Text1(ptxSHIJI_NO).SetFocus

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


    If App.PrevInstance Then
        MsgBox "同一プログラム実行中です。"
        End
    End If

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.10.09
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "商品化指図票発行　「起動処理中」", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
                                
                                
    Me.Enabled = False
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.10.09
                                
                                
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止します。"
        End
    End If
    LOG_F = RTrim(c)


                                'ラベル印刷用コントロールＦ獲得2008.05.30
    If GetIni("FILE", "labelprint", "SYS", c) Then
        Beep
        MsgBox "ラベル印刷用コントロールＦの獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LabelPrint_F = RTrim(c)

Show    '2015.03.26
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2016.01.29  PI00010.INI --> P_SYS.INI[PLABEL]
'------------------------------------------ P_SYS.INI--> PI00010.INI 2011.08.04

                                '原産国印字有無 2008.06.12
    If GetIni("PLABEL", "GENSANKOKU_DEF_F", "P_SYS", c) Then
        GENSANKOKU_FLG = "0"
    Else
        GENSANKOKU_FLG = RTrim(c)
    End If


                                '原産国空白ﾁｪｯｸ 2009.03.28
    If GetIni("PLABEL", "GENSANKOKU_CHECK", "P_SYS", c) Then
        ReDim GENSANKOKU_CHECK_TBL(0 To 0)
        GENSANKOKU_CHECK_TBL(0) = "*"
    Else
        GENSANKOKU_CHECK_TBL = Split(Trim(c), ",", -1)
    End If


                                '代表機種ﾁｪｯｸ   2012.09.03
    If GetIni("PLABEL", "KISHU_CHECK", "P_SYS", c) Then
        KISHU_CHECK = False
    Else
        If Trim(c) = "1" Then
            KISHU_CHECK = True
        Else
            KISHU_CHECK = False
        End If
    End If


                                '会社事業部エラーﾁｪｯｸ有無 2010.07.20
'    If GetIni(App.EXEName, "KAISYA_CHECK", App.EXEName, c) Then
    If GetIni("PLABEL", "KAISYA_CHECK", "P_SYS", c) Then
        KAISYA_CHK_F = False
    Else

        If Trim(c) = "1" Then
            KAISYA_CHK_F = True
        Else
            KAISYA_CHK_F = False
        End If
    End If
'

                                '原産国海外供給区分ﾁｪｯｸ 2016.02.01
    If GetIni("PLABEL", "GAI_BUHIN_CHK", "P_SYS", c) Then
        GAI_BUHIN_CHK = False
    Else

        If Trim(c) = "1" Then
            GAI_BUHIN_CHK = True
        Else
            GAI_BUHIN_CHK = False
        End If
    End If

                                '原産国海外供給区分ﾁｪｯｸ 2016.02.01
    If GetIni("PLABEL", "TANKA_SPACE_F", "P_SYS", c) Then
        TANKA_SPACE_F = "0"
    Else
        If Trim(c) = "1" Then
            TANKA_SPACE_F = "1"
        Else
            TANKA_SPACE_F = "0"
        End If
    End If



'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2016.01.29  PI00010.INI --> P_SYS.INI[PLABEL]

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 原産国ﾒｯｾｰｼﾞの表示 2013.02.19
    If GetIni(App.EXEName, "GENSANKOKU_MSG_F", App.EXEName, c) Then
        GENSANKOKU_MSG_F = False
    Else

        If Trim(c) = "1" Then
            GENSANKOKU_MSG_F = True
        Else
            GENSANKOKU_MSG_F = False
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 原産国ﾒｯｾｰｼﾞの表示 2013.02.19
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>> 会社名/事業部名非表示設定 1=非表示設定有り 2016.02.01
    If GetIni(App.EXEName, "KAISYA_RESTRICT_F", App.EXEName, c) Then
        KAISYA_RESTRICT_F = False
    Else
        If Trim(c) = "1" Then
            KAISYA_RESTRICT_F = True
        Else
            KAISYA_RESTRICT_F = False
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>> 会社名/事業部名非表示設定 1=非表示設定有り 2016.02.01




                                '出荷ログファイル名取り込み
    If SYUKA_LOGF_GET_PROC() Then
        Beep
        MsgBox "出荷ログファイル名の獲得に失敗しました。処理を中止します。"
        End
    End If

                                '手配先取り込み
    If GetIni(StrConv(App.EXEName, vbUpperCase), "TEHAI", App.EXEName, c) Then
    Else
        TEHAI = RTrim(c)
    End If

                                'POSｼｽﾃﾑ有無の取り込み
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
    If GetIni(StrConv(App.EXEName, vbUpperCase), "DET_BIKOU", App.EXEName, c) Then
        PRI_BIKOU_BCR = False
    Else
        If RTrim(c) = "0" Then
            PRI_BIKOU_BCR = False
        Else
            If Not POS_UMU Then
                MsgBox "ＰＯＳｼｽﾃﾑが未設定です。処理を中止します。"
                End
            End If
            PRI_BIKOU_BCR = True
        End If
    End If

                                '収単／担当者の取り込み
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
    If GetIni(StrConv(App.EXEName, vbUpperCase), "DOUKON", App.EXEName, c) Then
        PRI_DOUKON = False
    Else
        
        '2011.08.04
'        If RTrim(c) = "0" Then
'            PRI_DOUKON = False
'        Else
'            PRI_DOUKON = True
'        End If
        
        Select Case RTrim(c)
            Case "0"
                PRI_DOUKON = 0
            Case "1"
                PRI_DOUKON = 1
            Case "2"
                PRI_DOUKON = 2
            Case Else
                PRI_DOUKON = 0
        End Select
        '2011.08.04
    End If
                                '入庫完了印の取り込み
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
    If GetIni(StrConv(App.EXEName, vbUpperCase), "JISEKI", App.EXEName, c) Then
        JISEKI_TITLE = ""
    Else
        JISEKI_TITLE = Split(Trim(c), ",", -1)
    End If

                                '他責
    If GetIni(StrConv(App.EXEName, vbUpperCase), "TASEKI", App.EXEName, c) Then
        TASEKI_TITLE = ""
    Else
        TASEKI_TITLE = Split(Trim(c), ",", -1)
    End If

                                '未登録品番の可否
    If GetIni(StrConv(App.EXEName, vbUpperCase), "HIN_INV", App.EXEName, c) Then
        HIN_INV = False
    Else
        If Trim(c) = "0" Then
            HIN_INV = False
        Else
            HIN_INV = True
        End If

    End If



    If PRI_BIKOU_BCR Then
                                    '向け先
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

    '秒／分取り込み 2008.08.19
    If GetIni(StrConv(App.EXEName, vbUpperCase), "JISSEKI_DSP", App.EXEName, c) Then
        JISSEKI_DSP = "m"
    Else
        JISSEKI_DSP = Trim(c)
        If JISSEKI_DSP <> "m" And JISSEKI_DSP <> "s" Then
            JISSEKI_DSP = "m"
        End If
    End If




'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 開梱・リード巻線・粘着防止・の表示 2013.01.16
    If GetIni(App.EXEName, "KAIKON_PRI", App.EXEName, c) Then
        KAIKON_PRI = False
    Else

        If Trim(c) = "1" Then
            KAIKON_PRI = True
        Else
            KAIKON_PRI = False
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 開梱・リード巻線・粘着防止・の表示 2013.01.16







'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    半製品仕向け先  2013.08.29
    If GetIni(App.EXEName, "SHIMUKE_CHK", App.EXEName, c) Then
        
        
        ReDim SHIMUKE_CHK_TBL(0 To 0)
        SHIMUKE_CHK_TBL(0) = "**"
    Else
    
        SHIMUKE_CHK_TBL = Split(Trim(c), ",", -1)
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    半製品仕向け先  2013.08.29





'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 入荷時緩衝材の表示 2013.11.05
    If GetIni(App.EXEName, "NYUKA_KANSYOZAI", App.EXEName, c) Then
        NYUKA_KANSYOZAI = False
    Else

        If Trim(c) = "1" Then
            NYUKA_KANSYOZAI = True
        Else
            NYUKA_KANSYOZAI = False
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 入荷時緩衝材の表示 2013.11.05



'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> ラベル発行の表示 2019.03.07
    If GetIni(App.EXEName, "LABEL_PRINT_F", App.EXEName, c) Then
        LABEL_PRINT_F = 0
    Else

        If Trim(c) = "1" Then
            LABEL_PRINT_F = 1
        Else
            LABEL_PRINT_F = 0
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> ラベル発行の表示 2019.03.07


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 外装ラベル発行の表示 2019.03.07
    If GetIni(App.EXEName, "GA_LABEL_PRINT_F", App.EXEName, c) Then
        GA_LABEL_PRINT_F = 0
    Else

        If Trim(c) = "1" Then
            GA_LABEL_PRINT_F = 1
        Else
            GA_LABEL_PRINT_F = 0
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 外装ラベル発行の表示 2019.03.07



'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> ラベル発行枚数の指定 2015.04.02
    If GetIni(App.EXEName, "LABEL_PLUS", App.EXEName, c) Then
        LABEL_PLUS = 1
    Else
        If Not IsNumeric(Trim(c)) Then
            LABEL_PLUS = 1
        Else
            LABEL_PLUS = Val(Trim(c))
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> ラベル発行枚数の指定 2015.04.02


    PI000101.Caption = Last_Update_day      '2016.02.10



'ラベル選択     2011.02.10
    Combo2(0).Clear
    Combo2(0).AddItem "ラベルなし　　" & "          " & " "
    Combo2(0).AddItem "パーツラベル　" & "          " & "0"
    Combo2(0).AddItem "適用機種ラベル" & "          " & "1"
    Combo2(0).AddItem "JANラベル　　 " & "          " & "2"





'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.04.24
'
'                                '発番マスタＯＰＥＮ
'    If HATUBAN_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'                                '品目マスタＯＰＥＮ
'    If ITEM_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'
'                                '商品ﾗﾍﾞﾙ用品目マスタＯＰＥＮ
'    If L_ITEM_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'
'                                'クラスマスタＯＰＥＮ
'    If P_Class_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'                                'コードマスタＯＰＥＮ
'    If P_CODE_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'                                '構成マスタＯＰＥＮ
'    If P_COMPO_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'                                '管理マスタＯＰＥＮ
'    If P_KANRI_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'                                '商品化指図（子）ﾃﾞｰﾀＯＰＥＮ
'    If P_SSHIJI_K_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'                                '商品化指図（親）ﾃﾞｰﾀＯＰＥＮ
'    If P_SSHIJI_O_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'                                '担当者マスタＯＰＥＮ
'    If TANTO_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'                                '出荷予定ﾃﾞｰﾀＯＰＥＮ
'    If Y_SYU_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'                                '受払先マスタＯＰＥＮ
'    If P_UKEHARAI_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'
'
'    '2010.07.20 ▽
'                                '原産国マスタＯＰＥＮ
'    If GENSAN_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'    '2010.07.20 △
'                                '在庫ﾃﾞｰﾀＯＰＥＮ
'    If ZAIKO_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'
'                                '商品化指図（親）ﾜｰｸＯＰＥＮ
'    If wP_SSHIJI_O_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'
'                                '入出庫単価設定マスタＯＰＥＮ   2008.09.20
'    If SE_LOC_TANKA_M_Open(BtOpenNomal) Then
'        Unload Me
'    End If
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
                            
    Do
        If Not File_Open_Proc() Then
            Exit Do
        End If
    Loop
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.04.24


    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc



    Load PI000102
    Load PI000103



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

'2009.03.25
    Combo1(pcmbSHIMUKE).ListIndex = 0
    Last_JGYOBU = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1)      '2016.02.01
    
    PI000104_OLD_HIN_GAI = ""       '2019.03.14
    


    chenge_F = False


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.10.09
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "商品化指図票発行　「準備完了」", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
                                
                                
    Me.Enabled = True
    DoEvents
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.10.09
    Text1(ptxSHIJI_NO).SetFocus
    

End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer



    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  MT_2009.06.01
                                            'PNマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, PN_M_POS, PN_MREC, Len(PN_MREC), K0_PN_M, Len(K0_PN_M), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
'            Call File_Error(sts, BtOpClose, "PNマスタ")             2015.05.14
            Call File_Error(sts, BtOpClose, "PNマスタ", 0)          '2015.05.14
'2015.03.26            Beep
'2015.03.26            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        End If
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<



                                            '発番マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "発番マスタ", 0)
        End If
    End If

                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ", 0)
        End If
    End If
                                            '商品ﾗﾍﾞﾙ用品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品ﾗﾍﾞﾙ用品目マスタ", 0)
        End If
    End If

                                            'クラスマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "クラスマスタ", 0)
        End If
    End If

                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "コードマスタ", 0)
        End If
    End If

                                            '構成マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "構成マスタ", 0)
        End If
    End If
                                            '管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "管理マスタ", 0)
        End If
    End If
                                            '商品化指図ﾃﾞｰﾀ(親)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化指図ﾃﾞｰﾀ(親)", 0)
        End If
    End If
                                            '商品化指図ﾃﾞｰﾀ(子)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化指図ﾃﾞｰﾀ(子)", 0)
        End If
    End If

                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ", 0)
        End If
    End If

                                            '出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定データ", 0)
        End If
    End If
                                            '受払先マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "受払先マスタ", 0)
        End If
    End If
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ", 0)
        End If
    End If

                                            '商品化指図ﾜｰｸ(親)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, wP_SSHIJI_O_POS, wP_SSHIJI_O_REC, Len(wP_SSHIJI_O_REC), K0_wP_SSHIJI_O, Len(K0_wP_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化指図(親)ﾜｰｸ", 0)
        End If
    End If


    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "", 0)
    End If
    Set PI000101 = Nothing
    Set PI000102 = Nothing
    Set PI000103 = Nothing

    End
End Sub

Private Sub RichTextBox1_GotFocus(Index As Integer)
Dim sts         As Integer

Dim Sumi_Qty    As Long
Dim Mi_Qty      As Long


    Select Case Index
'        Case ptxHIN_GAI
        Case Else
            If chenge_F Then

Start_Proc1:        '2015.03.26
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

'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                        Check1(pchkL_PAPER).Value = vbUnchecked         '紙
                        Check1(pchkL_PLASTIC).Value = vbUnchecked       'プラ
                        Check1(pchkL_LABEL).Value = vbUnchecked         '適用機種ラベル
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<



                        MsgBox "入力した項目はエラーです。(品番)"

                        chenge_F = False

                        Text1(ptxHIN_GAI).SetFocus
                        Exit Sub
                    Case Else
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                            Do                                  '2015.04.24
                                If Not File_Open_Proc() Then    '2015.04.24
                                    Exit Do                     '2015.04.24
                                End If                          '2015.04.24
                            Loop                                '2015.04.24
                            '>>>>>>>>>>>>>  2015.04.24
                            
                            GoTo Start_Proc1
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
                        
                        
                        
                        
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Unload Me

                End Select



                Text1(ptxHIN_NAME).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)

'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                If StrConv(ITEMREC.L_PAPER, vbUnicode) = L_PAPER_ON Then        '紙
                    Check1(pchkL_PAPER).Value = vbChecked
                Else
                    Check1(pchkL_PAPER).Value = vbUnchecked
                End If
    
                If StrConv(ITEMREC.L_PLASTIC, vbUnicode) = L_PLASTIC_ON Then    'プラ
                    Check1(pchkL_PLASTIC).Value = vbChecked
                Else
                    Check1(pchkL_PLASTIC).Value = vbUnchecked
                End If
    
                If StrConv(ITEMREC.L_LABEL, vbUnicode) = L_LABEL_ON Then        '適用機種ラベル
                    Check1(pchkL_LABEL).Value = vbChecked
                Else
                    Check1(pchkL_LABEL).Value = vbUnchecked
                End If
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<








                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Text1(ptxST_LOCATION).text = ""
                Else
                    Text1(ptxST_LOCATION).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If

                If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                        StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                        StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then

                    Unload Me
                End If

                Text1(ptxMI_QTY).text = Format(Mi_Qty, "#0")
                Text1(ptxSUMI_QTY).text = Format(Sumi_Qty, "#0")



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
                            Case Else
                                
                                
                                
                                
                                
                                
                                
                                
                                Call File_Error(sts, BtOpGetEqual, "構成マスタ")
                                Unload Me
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
                            Case Else
                                
                                
                                

                                
                                
                                
                                Call File_Error(sts, BtOpGetEqual, "構成マスタ")
                                Unload Me
                        End Select

                    End If
                End If
'                Text1(ptxSHIJI_QTY).SetFocus

                chenge_F = False

                RichTextBox1(Index).SetFocus

            End If
    End Select

End Sub

Private Sub Text1_Change(Index As Integer)
    If Index = ptxSHIJI_QTY Then        '2008.02.27
        Text1(ptxLabel_QTY) = ""
    End If

    If Index = ptxHIN_GAI Then          '2008.07.10
'        If Trim(Text1(ptxHIN_GAI)) <> "" Then
            chenge_F = True
'        End If
    End If

End Sub

Private Sub Text1_GotFocus(Index As Integer)


Dim sts         As Integer

Dim Sumi_Qty    As Long
Dim Mi_Qty      As Long
Dim wkINDEX     As Integer

    If Index = ptxHIN_GAI Then
        If Trim(Text1(ptxHIN_GAI).text) = "" Then
        
        
            If Text1(Index).TabStop = True Then
                Text1(Index) = Trim(Text1(Index).text)
                Text1(Index).SelStart = 0
                Text1(Index).SelLength = Len(Text1(Index).text)
            End If
    
    
            Exit Sub
        End If
    End If
    
'    MsgBox "Index=" & Index
    
    Select Case Index
        
        Case ptxHIN_GAI
            Text1(Index) = Trim(Text1(Index).text)
            Text1(Index).SelStart = 0
            Text1(Index).SelLength = Len(Text1(Index).text)
            Exit Sub
        Case Else
        
'            If chenge_F Then
            '2019.06.20 ↑を＆条件で品番≠空白でのチェックとした。
            If chenge_F = True And Trim(Text1(ptxHIN_GAI)) <> "" Then
'                MsgBox "chenge_F=True"
                
Start_Proc1:        '2015.03.26
                wkINDEX = Index
                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).text)


                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                sts = BTRV(BtOpGetGreaterEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Text1(ptxHIN_GAI) = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
                    Case BtErrKeyNotFound, BtErrEOF

                        Text1(ptxHIN_NAME).text = ""
                        Text1(ptxST_LOCATION).text = ""
                        Text1(ptxMI_QTY).text = ""
                        Text1(ptxSUMI_QTY).text = ""

'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                        Check1(pchkL_PAPER).Value = vbUnchecked         '紙
                        Check1(pchkL_PLASTIC).Value = vbUnchecked       'プラ
                        Check1(pchkL_LABEL).Value = vbUnchecked         '適用機種ラベル
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


                        MsgBox "入力した項目はエラーです。(品番)"


                        Text1(ptxHIN_GAI).SetFocus

                        Text1(ptxHIN_GAI) = Trim(Text1(ptxHIN_GAI).text)
                        Text1(ptxHIN_GAI).SelStart = 0
                        Text1(ptxHIN_GAI).SelLength = Len(Text1(ptxHIN_GAI).text)


                        Exit Sub
                    Case Else
                        
                        
                        
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
                           
            
            
            
                        
                            GoTo Start_Proc1
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        
                        
                        
                        
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Unload Me

                End Select



                Text1(ptxHIN_NAME).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)



'2010.11.12 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                If StrConv(ITEMREC.L_PAPER, vbUnicode) = L_PAPER_ON Then        '紙
                    Check1(pchkL_PAPER).Value = vbChecked
                Else
                    Check1(pchkL_PAPER).Value = vbUnchecked
                End If
    
                If StrConv(ITEMREC.L_PLASTIC, vbUnicode) = L_PLASTIC_ON Then    'プラ
                    Check1(pchkL_PLASTIC).Value = vbChecked
                Else
                    Check1(pchkL_PLASTIC).Value = vbUnchecked
                End If
    
                If StrConv(ITEMREC.L_LABEL, vbUnicode) = L_LABEL_ON Then        '適用機種ラベル
                    Check1(pchkL_LABEL).Value = vbChecked
                Else
                    Check1(pchkL_LABEL).Value = vbUnchecked
                End If
'2010.11.12 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<




                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Text1(ptxST_LOCATION).text = ""
                Else
                    Text1(ptxST_LOCATION).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If

                If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                        StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                        StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then

                    Unload Me
                End If

                Text1(ptxMI_QTY).text = Format(Mi_Qty, "#0")
                Text1(ptxSUMI_QTY).text = Format(Sumi_Qty, "#0")



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
                            Case Else
                                
                                
                                
                                Call File_Error(sts, BtOpGetEqual, "構成マスタ")
                                Unload Me
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
                            Case Else
                                
                                
                                
                                Call File_Error(sts, BtOpGetEqual, "構成マスタ")
                                Unload Me
                        End Select

                    End If
                End If
    '                Text1(ptxSHIJI_QTY).SetFocus


                chenge_F = False
                Text1(wkINDEX).SetFocus
                Exit Sub
            End If

    End Select


    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).text)
    End If

    
    

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim WK_STR  As String


    If KeyCode <> vbKeyReturn Then Exit Sub
    


    Select Case Index
        Case ptxHIN_GAI, ptxK_HIN_GAI01, ptxK_HIN_GAI02, ptxK_HIN_GAI03, ptxK_HIN_GAI04, ptxK_HIN_GAI05, _
                ptxG_HIN_GAI01, ptxG_HIN_GAI02, ptxG_HIN_GAI03, _
                ptxD_HIN_GAI01, ptxD_HIN_GAI02, ptxD_HIN_GAI03, ptxD_HIN_GAI04, ptxD_HIN_GAI05, ptxD_HIN_GAI06 _


            Text1(Index).text = RTrim(StrConv(Text1(Index).text, vbUpperCase))

    End Select
    
    '2019.06.04 追加                要望の解釈で「取消」押下と同一処理として追加！
    If Index = ptxHIN_GAI Then
        WK_STR = Trim(Text1(Index))
'        Call Command1_Click(10)             '取り消しキー押下
        
        '2019.06.05 上記をCallではなく、その内容をコピーした。
        '           指図票№にSetFocusしていた為！
'        If Init_Proc() Then
        '2019.06.10 ↑を変更：「事前、スポット、欠品解除」のクリアしない！
        
        '2019.08.23 ↓一部、修正・・・指図書のコンボ初期処理しない！
        If Init_Proc_2() Then
            Unload Me
        End If
            
        Text1(ptxHIN_GAI).Locked = False            '2019.03.18
'            Text1(ptxHIN_GAI).Enabled = True           '2019.03.18
        Text1(ptxHIN_GAI).BackColor = &H80000005     '2019.03.18

        
        Text1(Index) = WK_STR
        DoEvents
    End If
    '2019.06.04 ここまで追加
    

    If Error_Check_Proc(Index, 0, 0) Then   'エラーチェック
        Exit Sub
    End If
    
    DoEvents
    
    If Index = ptxHIN_GAI Then
        Text1(ptxSHIJI_QTY) = Trim(Text1(ptxSHIJI_QTY).text)
        Text1(ptxSHIJI_QTY).SelStart = 0
        Text1(ptxSHIJI_QTY).SelLength = Len(Text1(ptxSHIJI_QTY).text)
        DoEvents
        Exit Sub
    Else
        Call Tab_Ctrl(Shift)        '移動
    End If
End Sub

Private Function Init_Proc() As Integer
'----------------------------------------------------------------------------
'                   入力画面の初期設定
'----------------------------------------------------------------------------
Dim i           As Integer
Dim sts         As Integer

Dim TANTO_CODE  As String
Dim TANTO_NAME  As String

    Init_Proc = True

    Text1(ptxSHIJI_NO).BackColor = G_INPUT_OK
    Text1(ptxSHIJI_NO).Locked = False
    Text1(ptxSHIJI_NO).TabStop = True


    Combo1(pcmbS_TANTO).Enabled = PRI_S_TANTO

    TANTO_CODE = Text1(ptxTANTO_CODE).text
    TANTO_NAME = Text1(ptxTANTO_NAME).text

    For i = ptxSHIJI_NO To ptxLabel_QTY
        Text1(i).text = ""
    Next i
    Text1(ptxTANTO_CODE).text = TANTO_CODE
    Text1(ptxTANTO_NAME).text = TANTO_NAME



    RichTextBox1(prchBIKOU).text = ""

    For i = pchkSAMPLE_F To pchkPRI_KISHU
        Check1(i).Value = vbChecked
    Next i
    Check1(pchkSAMPLE_F).Value = vbUnchecked    '見本作成
    Check1(pchkPRI_KISHU).Value = vbUnchecked   '出力対象　機種ﾗﾍﾞﾙ

    Check1(pchkL_PAPER).Value = vbUnchecked     '紙           2010.11.12
    Check1(pchkL_PLASTIC).Value = vbUnchecked   'ﾌﾟﾗｽﾁｯｸ      2010.11.12
    Check1(pchkL_LABEL).Value = vbUnchecked     '適用機種ﾗﾍﾞﾙ 2010.11.12


'2009.03.25
'    For i = pcmbSHIMUKE To pcmbD_SYUBETSU06
'
'            Combo1(i).ListIndex = -1
'
'    Next i
'    Combo1(pcmbSHIMUKE).ListIndex = 0


    For i = pcmbUKEHARAI To pcmbD_SYUBETSU06

            Combo1(i).ListIndex = -1

    Next i
'    Combo1(pcmbUKEHARAI).ListIndex = 0
'2009.03.25



    '発行日
    Text1(ptxHAKKO_DT).text = Format(Now, "YYYY/MM/DD")


    '承認者設定
    Text1(ptxSHONIN_CODE).text = StrConv(P_KANRIREC.SHONIN_CODE, vbUnicode)


    PI000104_OLD_HIN_GAI = ""       '2019.03.14

Start_Proc1:        '2015.03.26ok

    Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(P_KANRIREC.SHONIN_CODE, vbUnicode))

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxSHONIN_NAME).text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text1(ptxSHONIN_NAME).text = ""

        Case Else
            
            
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
                            
                            Call File_Error(sts, BtOpGetEqual, "担当者マスタ", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc1
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
            
            
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Exit Function
    End Select
    '手配先
    Text1(ptxUKEHARAI_CODE).text = TEHAI
    txGensankoku.text = ""                  '2009.03.28



    lblGensankoku(0).Caption = ""
    lblGensankoku(1).Caption = ""

    '指示形態
    Option1(poptSHIJI_NORMAL).Value = True
    Option1(poptSHIJI_SPOT).Value = False
    Option1(poptSHIJI_KEPPIN).Value = False


    '2011.02.10
    Combo2(0).ListIndex = 1
    '2011.02.10

    If LABEL_PRINT_F = 1 Then                       '2019.03.07
        Combo2(0).ListIndex = 0                     '2019.03.07
    End If                                          '2019.03.07



    '>>>>>>>>>>>>>>>>>>>>>> 2013.09.12
    For i = 0 To UBound(SHIMUKE_CHK_TBL)
    
        If SHIMUKE_CHK_TBL(i) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Then
            Combo2(0).ListIndex = 0
            Check1(pchkPRI_GAISOU).Value = vbUnchecked
            Exit For
        End If
    
    Next i
    '>>>>>>>>>>>>>>>>>>>>>> 2013.09.12

    
    If GA_LABEL_PRINT_F = 1 Then                    '2019.03.07
        Check1(pchkPRI_GAISOU).Value = vbUnchecked  '2019.03.07
    End If                                          '2019.03.07
    '2019.09.24 下記フラグをクリア（False）
    '                                       これは、Combo2(0)の右端<>"" もしくは外装ラベルの指示でTrueにしている。
    L_print_Flg = False


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

    Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_No, "")




    Text1(ptxHIN_GAI).Locked = False            '2019.03.18
'            Text1(ptxHIN_GAI).Enabled = True           '2019.03.18
    Text1(ptxHIN_GAI).BackColor = &H80000005     '2019.03.18



    Init_Proc = False

End Function


Private Function Init_Proc_2() As Integer
'----------------------------------------------------------------------------
'                   入力画面の初期設定
'----------------------------------------------------------------------------
Dim i           As Integer
Dim sts         As Integer

Dim TANTO_CODE  As String
Dim TANTO_NAME  As String

    Init_Proc_2 = True

    Text1(ptxSHIJI_NO).BackColor = G_INPUT_OK
    Text1(ptxSHIJI_NO).Locked = False
    Text1(ptxSHIJI_NO).TabStop = True


    Combo1(pcmbS_TANTO).Enabled = PRI_S_TANTO

    TANTO_CODE = Text1(ptxTANTO_CODE).text
    TANTO_NAME = Text1(ptxTANTO_NAME).text

    For i = ptxSHIJI_NO To ptxLabel_QTY
        Text1(i).text = ""
    Next i
    Text1(ptxTANTO_CODE).text = TANTO_CODE
    Text1(ptxTANTO_NAME).text = TANTO_NAME



    RichTextBox1(prchBIKOU).text = ""

    For i = pchkSAMPLE_F To pchkPRI_KISHU
        Check1(i).Value = vbChecked
    Next i
    Check1(pchkSAMPLE_F).Value = vbUnchecked    '見本作成
    Check1(pchkPRI_KISHU).Value = vbUnchecked   '出力対象　機種ﾗﾍﾞﾙ

    Check1(pchkL_PAPER).Value = vbUnchecked     '紙           2010.11.12
    Check1(pchkL_PLASTIC).Value = vbUnchecked   'ﾌﾟﾗｽﾁｯｸ      2010.11.12
    Check1(pchkL_LABEL).Value = vbUnchecked     '適用機種ﾗﾍﾞﾙ 2010.11.12


'2009.03.25
'    For i = pcmbSHIMUKE To pcmbD_SYUBETSU06
'
'            Combo1(i).ListIndex = -1
'
'    Next i
'    Combo1(pcmbSHIMUKE).ListIndex = 0


    For i = pcmbUKEHARAI To pcmbD_SYUBETSU06

            Combo1(i).ListIndex = -1

    Next i
'    Combo1(pcmbUKEHARAI).ListIndex = 0
'2009.03.25



    '発行日
    Text1(ptxHAKKO_DT).text = Format(Now, "YYYY/MM/DD")


    '承認者設定
    Text1(ptxSHONIN_CODE).text = StrConv(P_KANRIREC.SHONIN_CODE, vbUnicode)


    PI000104_OLD_HIN_GAI = ""       '2019.03.14

Start_Proc1:        '2015.03.26ok

    Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(P_KANRIREC.SHONIN_CODE, vbUnicode))

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxSHONIN_NAME).text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text1(ptxSHONIN_NAME).text = ""

        Case Else
            
            
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
                            
                            Call File_Error(sts, BtOpGetEqual, "担当者マスタ", 0)
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc1
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
            
            
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Exit Function
    End Select
    '手配先
    Text1(ptxUKEHARAI_CODE).text = TEHAI
    txGensankoku.text = ""                  '2009.03.28



    lblGensankoku(0).Caption = ""
    lblGensankoku(1).Caption = ""
    
    '2019.06.10 「Init_Proc」との違いは、下記の３行が有効か否か！のみ、
'    '指示形態
'    Option1(poptSHIJI_NORMAL).Value = True
'    Option1(poptSHIJI_SPOT).Value = False
'    Option1(poptSHIJI_KEPPIN).Value = False

    '2019.08.23 下記をコメントにした。
    '           ⇒ 仕向先によって「ラベルなし」とセットしてあっても、Clearされてしまう！
    ''2019.08.28 下記を復帰してみた。
    '                   2019.09.24 再度、コメントにした！
    '2011.02.10
'    Combo2(0).ListIndex = 1
    '2011.02.10

    If LABEL_PRINT_F = 1 Then                       '2019.03.07
        Combo2(0).ListIndex = 0                     '2019.03.07
    End If                                          '2019.03.07
    ''2019.08.23 ここまで
    ''2019.08.28 ここまで

    '>>>>>>>>>>>>>>>>>>>>>> 2013.09.12
    For i = 0 To UBound(SHIMUKE_CHK_TBL)
    
        If SHIMUKE_CHK_TBL(i) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Then
            Combo2(0).ListIndex = 0
            Check1(pchkPRI_GAISOU).Value = vbUnchecked
            Exit For
        End If
    
    Next i
    '>>>>>>>>>>>>>>>>>>>>>> 2013.09.12

    '2019.08.23 下記をコメントにした。
    '           ⇒ 仕向先によって「ラベルなし」とセットしてあっても、Ckearされてしまう！
'    If GA_LABEL_PRINT_F = 1 Then                    '2019.03.07
'        Check1(pchkPRI_GAISOU).Value = vbUnchecked  '2019.03.07
'    End If                                          '2019.03.07

    '2019.09.24 上記を復帰
    If GA_LABEL_PRINT_F = 1 Then                    '2019.03.07
        Check1(pchkPRI_GAISOU).Value = vbUnchecked  '2019.03.07
    End If                                          '2019.03.07
    
    '2019.09.24 下記フラグをクリア（False）
    '                                       これは、Combo2(0)の右端<>"" もしくは外装ラベルの指示でTrueにしている。
    L_print_Flg = False
    
    
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

    Call UniCode_Conv(P_SSHIJI_O_REC.SHIJI_No, "")




    Text1(ptxHIN_GAI).Locked = False            '2019.03.18
'            Text1(ptxHIN_GAI).Enabled = True           '2019.03.18
    Text1(ptxHIN_GAI).BackColor = &H80000005     '2019.03.18



    Init_Proc_2 = False

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

Start_Proc0:        '2015.03.26ok

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
                
                
                
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "コードマスタ", 0)
                            
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
                        
                            GoTo Start_Proc0
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                
                
                
                
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

Start_Proc1:            '2015.03.26ok

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
                
                
                
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                If sts > 3000 Or sts = 3 Then


                    Call File_Error(sts, BtOpGetEqual, "受払先マスタ", 0)
                    
                    
                    
                    '>>>>>>>>>>>>>  2015.04.24
                    'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                    'If sts Then
                    '    Call File_Error(sts, BtOpReset, "")
                    'End If
                    'Call File_Open_Proc
                    Do
                        If Not File_Open_Proc() Then
                            Exit Do
                        End If
                    Loop
                    '>>>>>>>>>>>>>  2015.04.24
    
                
                    GoTo Start_Proc1
                End If
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                
                
                
                
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
    chenge_F = False

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

Dim wkHin_Gai   As String * 20          '2019.03.14


    P_COMPO_Disp_Proc = True
    Call Input_Lock             '2008.01.15

Start_Proc1:    '2013.05.26ok


    For i = ptxK_HIN_GAI01 To ptxD_BIKOU06
        Text1(i).text = ""
    Next i

                                '2008.07.30
    For i = pcmbD_SYUBETSU01 To pcmbD_SYUBETSU06

            Combo1(i).ListIndex = -1

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

    Text1(ptxLabel_QTY).text = ""       '2008.02.27



    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
        
    
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHIN_GAI).text)

    '2019.03.14
    If Trim(PI000104_OLD_HIN_GAI) <> "" Then
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, PI000104_OLD_HIN_GAI)
    End If
    '2019.03.14


    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_HEAD)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")

    sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)

    Select Case sts
        Case BtNoErr

        Case Else


                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "構成マスタ", 0)
                                                        
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc1
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26


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


            '2012.03.21
            RichTextBox1(prchBIKOU) = ""


            Call Input_UnLock           '2008.01.15
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

                '2019.03.14
                wkHin_Gai = Text1(ptxHIN_GAI).text
                If Trim(PI000104_OLD_HIN_GAI) <> "" Then
                    wkHin_Gai = PI000104_OLD_HIN_GAI
                End If
                '2019.03.14

                If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                    StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                    StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                    Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(wkHin_Gai) Then

                    Exit Do

                End If

            Case BtErrEOF
                Exit Do
            Case Else
                
                
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                        If sts > 3000 Or sts = 3 Then
    
        
                            Call File_Error(sts, BtOpGetEqual, "構成マスタ", 0)
                            
                            '>>>>>>>>>>>>>  2015.04.24
                            'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            'If sts Then
                            '    Call File_Error(sts, BtOpReset, "")
                            'End If
                            'Call File_Open_Proc
                            Do
                                If Not File_Open_Proc() Then
                                    Exit Do
                                End If
                            Loop
                            '>>>>>>>>>>>>>  2015.04.24
            
                        
                            GoTo Start_Proc1
                        End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                
                Call Input_UnLock             '2008.01.15
                Call File_Error(sts, BtOpGetNext, "構成マスタ")
                Exit Function


        End Select

        Select Case StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode)

            Case P_KOSOU    '個装資材

                k = k + 1

                If k > 36 Then
                    MsgBox "個装資材登録件数がオーバーしています。削除してください"
                Else
                    K_Item_Tbl(k).JGYOBU = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
                    K_Item_Tbl(k).NAIGAI = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
                                '品番
                    Text1(K_Index).text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)

                    Call UniCode_Conv(K0_ITEM.JGYOBU, K_Item_Tbl(k).JGYOBU)
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
                                    
                                    
                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                    If sts > 3000 Or sts = 3 Then
                
                    
                                        Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                        '>>>>>>>>>>>>>  2015.04.24
                                        'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                        'If sts Then
                                        '    Call File_Error(sts, BtOpReset, "")
                                        'End If
                                        'Call File_Open_Proc
                                        Do
                                            If Not File_Open_Proc() Then
                                                Exit Do
                                            End If
                                        Loop
                                        '>>>>>>>>>>>>>  2015.04.24
                        
                                    
                                        GoTo Start_Proc1
                                    End If
                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                    
                                    
                                    Call Input_UnLock             '2008.01.15
                                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                    Exit Function
    
                            End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21
                        Case Else
                            
                            
                            
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc1
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            
                            
                            Call Input_UnLock             '2008.01.15
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
                End If


            Case P_GAISOU   '外装資材
                g = g + 1


                If g > 51 Then
                    MsgBox "外装資材登録件数がオーバーしています。削除してください"
                Else

                    G_Item_Tbl(g).JGYOBU = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
                    G_Item_Tbl(g).NAIGAI = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
                                '品番
                    Text1(G_Index).text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)

                    Call UniCode_Conv(K0_ITEM.JGYOBU, G_Item_Tbl(g).JGYOBU)
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
                                    
                                    
                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                    If sts > 3000 Or sts = 3 Then
                
                    
                                        Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                        '>>>>>>>>>>>>>  2015.04.24
                                        'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                        'If sts Then
                                        '    Call File_Error(sts, BtOpReset, "")
                                        'End If
                                       'Call File_Open_Proc
                                        Do
                                            If Not File_Open_Proc() Then
                                                Exit Do
                                            End If
                                        Loop
                                        '>>>>>>>>>>>>>  2015.04.24
                        
                                    
                                        GoTo Start_Proc1
                                    End If
                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                    
                                    
                                    
                                    Call Input_UnLock             '2008.01.15
                                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                    Exit Function
    
                            End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21
                            
                        
                        Case Else
                            
                            
                            
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc1
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            
                            
                            
                            Call Input_UnLock             '2008.01.15
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
                End If

            Case P_DOUKON   '同梱／構成

                d = d + 1
                D_Item_Tbl(d).SYUBETSU = StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode)
                D_Item_Tbl(d).JGYOBU = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
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

                    DC_Index = DC_Index + 1

                                '品番
                    Text1(DT_Index).text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)


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


                            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                                Exit Function

                            End If

                            Text1(DT_Index + 5) = Format(Sumi_Qty + Mi_Qty, "#0")


                        Case BtErrKeyNotFound


'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21
                            
                            
                            Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
    
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
        
        
                                    If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                                            StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                                        Exit Function
        
                                    End If
        
                                    Text1(DT_Index + 5) = Format(Sumi_Qty + Mi_Qty, "#0")
                                
                                Case BtErrKeyNotFound
    

                                    Text1(DT_Index + 1) = "未登録品番"
                                    Text1(DT_Index + 4) = ""
                                    Text1(DT_Index + 5) = ""
                                Case Else
                                    
                                    
                                    
                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                    If sts > 3000 Or sts = 3 Then
                
                    
                                        Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                        '>>>>>>>>>>>>>  2015.04.24
                                        'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                        'If sts Then
                                        '    Call File_Error(sts, BtOpReset, "")
                                        'End If
                                        'Call File_Open_Proc
                                        Do
                                            If Not File_Open_Proc() Then
                                                Exit Do
                                            End If
                                        Loop
                                        '>>>>>>>>>>>>>  2015.04.24
                        
                                    
                                        GoTo Start_Proc1
                                    End If
                                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                                    
                                    
                                    
                                    
                                    
                                    
                                    Call Input_UnLock             '2008.01.15
                                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                                    Exit Function
    
                            End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>> 品番未登録表示の対応    2012.12.21








                        Case Else
                            
                            
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            If sts > 3000 Or sts = 3 Then
        
            
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                '>>>>>>>>>>>>>  2015.04.24
                                'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                                'If sts Then
                                '    Call File_Error(sts, BtOpReset, "")
                                'End If
                                'Call File_Open_Proc
                                Do
                                    If Not File_Open_Proc() Then
                                        Exit Do
                                    End If
                                Loop
                                '>>>>>>>>>>>>>  2015.04.24
                
                            
                                GoTo Start_Proc1
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                            
                            
                            
                            
                            
                            
                            Call Input_UnLock             '2008.01.15
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

    Call Input_UnLock             '2008.01.15


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
Start_Proc1:        '2015.03.26
            Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(i).JGYOBU)
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


                    If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                            StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                        Exit Function
                    End If
                    '在庫数
                    Text1(DT_Index + 5) = Format(Sumi_Qty + Mi_Qty, "#0")

                Case BtErrKeyNotFound

                    Text1(DT_Index + 1) = "未登録品番"
                    Text1(DT_Index + 4) = ""
                    Text1(DT_Index + 5) = ""

                Case Else
                    
                    
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                    If sts > 3000 Or sts = 3 Then

    
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                        '>>>>>>>>>>>>>  2015.04.24
                        'sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                        'If sts Then
                        '    Call File_Error(sts, BtOpReset, "")
                        'End If
                        'Call File_Open_Proc
                        Do
                            If Not File_Open_Proc() Then
                                Exit Do
                            End If
                        Loop
                        '>>>>>>>>>>>>>  2015.04.24
        
                    
                        GoTo Start_Proc1
                    End If
                    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.03.26
                    
                    
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

    If Den_No_Set_Proc(21, Last_JGYOBU, ID_NO) Then                         'IDNO
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
    If Den_No_Set_Proc(20, Last_JGYOBU, DEN_NO) Then
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
                sts = Den_No_Set_Proc(21, Last_JGYOBU, ID_NO)
                If sts Then
                    Exit Function
                End If

                Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO)
                Call UniCode_Conv(Y_SYUREC.ID_NO, ID_NO)

Debug.Print StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode)
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

'Private Sub Text1_LostFocus(Index As Integer)
'
'
'    Select Case Index
'        Case ptxHIN_GAI, ptxK_HIN_GAI01, ptxK_HIN_GAI02, ptxK_HIN_GAI03, ptxK_HIN_GAI04, ptxK_HIN_GAI05, _
'                ptxG_HIN_GAI01, ptxG_HIN_GAI02, ptxG_HIN_GAI03, _
'                ptxD_HIN_GAI01, ptxD_HIN_GAI02, ptxD_HIN_GAI03, ptxD_HIN_GAI04, ptxD_HIN_GAI05, ptxD_HIN_GAI06 _
'
'
'            Text1(Index).text = RTrim(StrConv(Text1(Index).text, vbUpperCase))
'
'    End Select
'
'End Sub

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
        Mesg = Mesg & "○品番   " & Text1(ptxHIN_GAI).text & Chr(13) & Chr(10)
        Mesg = Mesg & "○品名   " & Text1(ptxHIN_NAME).text & Chr(13) & Chr(10)
        Mesg = Mesg & "○品名Ｅ " & lblL_Hin_Name_E.Caption & Chr(13) & Chr(10)

    
        GENSANKOKU = lblGensankoku(1)




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
                        For i = lstGensankoku.ListCount - 1 To 0 Step -1
                            GENSANKOKU = Right(lstGensankoku.List(i), 20)
                            
                            If i = lstGensankoku.ListCount - 1 Then
                                Mesg = Mesg & GENSANKOKU
                            Else
                                Mesg = Mesg & "　　　　　   " & GENSANKOKU
                            End If
                            If i = 0 Then
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
    
    
        Mesg = Mesg & "○供給区分海外 " & "     " & lblGAI_BUHIN.Caption & Chr(13) & Chr(10)
    
        If IsNumeric(lblL_URIKIN2) Then
            Tanka2 = Format(CDbl(lblL_URIKIN2), "#0")
        Else
            Tanka2 = ""
        End If
        If IsNumeric(lblL_URIKIN3) Then
            Tanka3 = Format(CDbl(lblL_URIKIN3), "#0")
        Else
            Tanka3 = ""
        End If
            
        
        
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
            KAISYA_NAME = lblL_KAISHA_N.Caption
            JGYOBU_NAME = lblL_JGYOBU_N.Caption
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
    
    
'        Mesg = Mesg & "　　　　【ＯＫ】パーツラベルを印刷" & Chr(13) & Chr(10)     '2016.02.10
        Mesg = Mesg & "　　　　【ＯＫ】印刷／更新" & Chr(13) & Chr(10)              '2016.02.10
        Mesg = Mesg & " 【キャンセル】印刷中止" & Chr(13) & Chr(10)
    
    
    
'        Mesg_Set_Proc = MsgBox(Mesg, vbOKCancel + vbDefaultButton2 + vbExclamation, "パーツラベル項目確認")    '2016.02.10
        Mesg_Set_Proc = MsgBox(Mesg, vbOKCancel + vbDefaultButton1 + vbExclamation, "パーツラベル項目確認")     '2016.02.10


End Function


