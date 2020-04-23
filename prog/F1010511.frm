VERSION 5.00
Begin VB.Form F1010511 
   BackColor       =   &H00FFFFFF&
   Caption         =   "品目マスタメンテナンス（削除機能付き）"
   ClientHeight    =   11115
   ClientLeft      =   1920
   ClientTop       =   2295
   ClientWidth     =   15210
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
   ScaleHeight     =   11115
   ScaleWidth      =   15210
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   50
      Left            =   9120
      MaxLength       =   1
      TabIndex        =   134
      TabStop         =   0   'False
      Top             =   4440
      Width           =   300
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   49
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   4440
      Width           =   300
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   48
      Left            =   12600
      MaxLength       =   8
      TabIndex        =   130
      Top             =   9960
      Width           =   1350
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   47
      Left            =   12600
      MaxLength       =   8
      TabIndex        =   128
      Top             =   9360
      Width           =   1350
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   46
      Left            =   12600
      MaxLength       =   1
      TabIndex        =   125
      Top             =   8640
      Width           =   252
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   45
      Left            =   12600
      MaxLength       =   8
      TabIndex        =   123
      Top             =   7680
      Width           =   1350
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   315
      Index           =   7
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   121
      Top             =   10680
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   315
      Index           =   6
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   120
      Top             =   10680
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   315
      Index           =   5
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   119
      Top             =   10680
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   315
      Index           =   4
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   118
      Top             =   10680
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   117
      Top             =   10680
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   116
      Top             =   10680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   115
      Top             =   10680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   114
      Top             =   10680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   44
      Left            =   6195
      MaxLength       =   1
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   3960
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ﾗﾍﾞﾙ貼り計上なし"
      Height          =   255
      Index           =   0
      Left            =   12576
      TabIndex        =   98
      Top             =   6840
      Width           =   2295
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   43
      Left            =   2205
      TabIndex        =   109
      TabStop         =   0   'False
      Top             =   3960
      Width           =   2505
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   11280
      TabIndex        =   108
      Top             =   9840
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   4  '全角ひらがな
      Index           =   41
      Left            =   3255
      TabIndex        =   107
      Top             =   3360
      Width           =   4950
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   40
      Left            =   2205
      MaxLength       =   8
      TabIndex        =   106
      Top             =   3360
      Width           =   1065
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   42
      Left            =   9450
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1350
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
      Left            =   10290
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   39
      Left            =   300
      MaxLength       =   10
      TabIndex        =   102
      Top             =   10560
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   38
      Left            =   13200
      MaxLength       =   10
      TabIndex        =   100
      Top             =   10680
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   37
      Left            =   12540
      MaxLength       =   2
      TabIndex        =   97
      Top             =   6240
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   36
      Left            =   12600
      MaxLength       =   10
      TabIndex        =   95
      Top             =   5520
      Width           =   1350
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   35
      Left            =   12600
      MaxLength       =   10
      TabIndex        =   93
      Top             =   4800
      Width           =   1350
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   34
      Left            =   12600
      MaxLength       =   10
      TabIndex        =   91
      Top             =   4080
      Width           =   1350
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   33
      Left            =   12600
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   89
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   32
      Left            =   12900
      MaxLength       =   4
      TabIndex        =   87
      Top             =   2640
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
      Top             =   2640
      Width           =   255
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   30
      Left            =   5760
      MaxLength       =   20
      TabIndex        =   82
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   29
      Left            =   2280
      MaxLength       =   13
      TabIndex        =   80
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   21
      Left            =   10200
      MaxLength       =   8
      TabIndex        =   22
      Top             =   1680
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
      Top             =   1680
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
      Top             =   2160
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
      Top             =   2160
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
      Top             =   2160
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
      Top             =   1680
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
      Top             =   1680
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
      Top             =   1680
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
      Top             =   1680
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
      Top             =   1680
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
      Top             =   1680
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
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   13
      Left            =   13200
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
      Left            =   12480
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
      Left            =   11760
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
      Left            =   11040
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
      Left            =   8400
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
      Left            =   7560
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
      Left            =   6480
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
      MaxLength       =   20
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
      MaxLength       =   40
      TabIndex        =   3
      Top             =   600
      Width           =   4935
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   0
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   1
      Top             =   600
      Width           =   1695
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
      Left            =   9465
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   8625
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "データ"
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
      Left            =   7785
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   6465
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   5625
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   4785
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   3945
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   2625
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   1785
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   945
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   105
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   3660
      Left            =   240
      TabIndex        =   28
      Top             =   5040
      Width           =   11955
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   26
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   76
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   27
      Left            =   7680
      MaxLength       =   2
      TabIndex        =   77
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   28
      Left            =   8520
      MaxLength       =   2
      TabIndex        =   78
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "海外供給区分"
      Height          =   255
      Index           =   57
      Left            =   7200
      TabIndex        =   135
      Top             =   4560
      Width           =   1680
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      Caption         =   "0:非対象/1:対象/2:打切案内中/3:打切"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   5
      Left            =   2640
      TabIndex        =   133
      Top             =   4560
      Width           =   4200
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "国内供給区分"
      Height          =   255
      Index           =   56
      Left            =   240
      TabIndex        =   131
      Top             =   4560
      Width           =   1680
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "[商品化工数]"
      Height          =   252
      Index           =   55
      Left            =   12600
      TabIndex        =   129
      Top             =   9720
      Width           =   1572
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "生産ロット数"
      Height          =   252
      Index           =   54
      Left            =   12480
      TabIndex        =   127
      Top             =   9120
      Width           =   1572
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "(1:除外)"
      Height          =   252
      Index           =   53
      Left            =   12960
      TabIndex        =   126
      Top             =   8760
      Width           =   1092
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "「商品化計画」除外ﾌﾗｸﾞ"
      Height          =   252
      Index           =   52
      Left            =   12360
      TabIndex        =   124
      Top             =   8280
      Width           =   2652
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "入数(出荷確認計算用)"
      Height          =   252
      Index           =   51
      Left            =   12600
      TabIndex        =   122
      Top             =   7320
      Width           =   2532
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(0:請求対象外 1:PPSC請求 2:BU請求　3:PPSC/BU請求)"
      Height          =   360
      Index           =   0
      Left            =   6510
      TabIndex        =   113
      Top             =   4080
      Visible         =   0   'False
      Width           =   6270
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "商品化請求F"
      Height          =   255
      Index           =   50
      Left            =   4830
      TabIndex        =   111
      Top             =   4080
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "出荷検品ﾒｯｾｰｼﾞ"
      Height          =   255
      Index           =   49
      Left            =   315
      TabIndex        =   110
      Top             =   4080
      Width           =   1680
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "ﾒｰｶｰ"
      Height          =   255
      Index           =   48
      Left            =   1470
      TabIndex        =   105
      Top             =   3480
      Width           =   630
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "個装形態"
      Height          =   255
      Index           =   47
      Left            =   8400
      TabIndex        =   104
      Top             =   3480
      Width           =   1050
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "P2在庫"
      Height          =   252
      Index           =   46
      Left            =   300
      TabIndex        =   101
      Top             =   10320
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "S2在庫"
      Height          =   252
      Index           =   45
      Left            =   13200
      TabIndex        =   99
      Top             =   10440
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "収単/担当者"
      Height          =   252
      Index           =   44
      Left            =   12540
      TabIndex        =   96
      Top             =   6000
      Width           =   1332
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "ｸﾞﾘｯｸｽ棚番3"
      Height          =   252
      Index           =   43
      Left            =   12600
      TabIndex        =   94
      Top             =   5280
      Width           =   1452
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "ｸﾞﾘｯｸｽ棚番2"
      Height          =   252
      Index           =   26
      Left            =   12600
      TabIndex        =   92
      Top             =   4560
      Width           =   1452
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "ｸﾞﾘｯｸｽ棚番1"
      Height          =   252
      Index           =   25
      Left            =   12600
      TabIndex        =   90
      Top             =   3840
      Width           =   1452
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "月平均出荷数"
      Height          =   252
      Index           =   24
      Left            =   12600
      TabIndex        =   88
      Top             =   3120
      Width           =   1452
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "個装箱№"
      Height          =   255
      Index           =   42
      Left            =   11655
      TabIndex        =   86
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "(0:要　1:不要)"
      Height          =   255
      Index           =   41
      Left            =   9240
      TabIndex        =   85
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "商品化有無"
      Height          =   255
      Index           =   40
      Left            =   7560
      TabIndex        =   83
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "読替えコード"
      Height          =   255
      Index           =   39
      Left            =   4080
      TabIndex        =   81
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "Ｊａｎコード"
      Height          =   255
      Index           =   38
      Left            =   600
      TabIndex        =   79
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "（最新照合日付"
      Height          =   255
      Index           =   37
      Left            =   4800
      TabIndex        =   75
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "／"
      Height          =   255
      Index           =   36
      Left            =   7320
      TabIndex        =   74
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "／"
      Height          =   255
      Index           =   35
      Left            =   8160
      TabIndex        =   73
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "）"
      Height          =   255
      Index           =   34
      Left            =   9000
      TabIndex        =   72
      Top             =   2280
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
      Height          =   312
      Left            =   240
      TabIndex        =   71
      Top             =   9360
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
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
      BackStyle       =   0  '透明
      Caption         =   "備考"
      Height          =   255
      Index           =   32
      Left            =   9120
      TabIndex        =   69
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "）"
      Height          =   255
      Index           =   31
      Left            =   4680
      TabIndex        =   68
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "／"
      Height          =   255
      Index           =   30
      Left            =   3840
      TabIndex        =   67
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "／"
      Height          =   255
      Index           =   29
      Left            =   3000
      TabIndex        =   66
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "（最終入荷日付"
      Height          =   255
      Index           =   28
      Left            =   480
      TabIndex        =   65
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "サンプル数"
      Height          =   255
      Index           =   27
      Left            =   11400
      TabIndex        =   64
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "）"
      Height          =   255
      Index           =   23
      Left            =   9000
      TabIndex        =   63
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "／"
      Height          =   255
      Index           =   22
      Left            =   8160
      TabIndex        =   62
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "／"
      Height          =   255
      Index           =   21
      Left            =   7320
      TabIndex        =   61
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "（最終出庫日付"
      Height          =   255
      Index           =   20
      Left            =   4800
      TabIndex        =   60
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "）"
      Height          =   255
      Index           =   19
      Left            =   4680
      TabIndex        =   59
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "／"
      Height          =   255
      Index           =   18
      Left            =   3840
      TabIndex        =   58
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "／"
      Height          =   255
      Index           =   17
      Left            =   3000
      TabIndex        =   57
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "（最終入庫日付"
      Height          =   255
      Index           =   16
      Left            =   480
      TabIndex        =   56
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "）"
      Height          =   255
      Index           =   15
      Left            =   13680
      TabIndex        =   55
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "-"
      Height          =   255
      Index           =   13
      Left            =   12960
      TabIndex        =   54
      Top             =   1140
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "-"
      Height          =   255
      Index           =   12
      Left            =   12240
      TabIndex        =   53
      Top             =   1140
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "-"
      Height          =   255
      Index           =   11
      Left            =   11520
      TabIndex        =   52
      Top             =   1140
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "（前回入庫棚"
      Height          =   255
      Index           =   10
      Left            =   9120
      TabIndex        =   51
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "）"
      Height          =   255
      Index           =   9
      Left            =   8880
      TabIndex        =   50
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "／"
      Height          =   255
      Index           =   8
      Left            =   8040
      TabIndex        =   49
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "／"
      Height          =   255
      Index           =   7
      Left            =   7200
      TabIndex        =   48
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "品番（内部）"
      Height          =   240
      Index           =   14
      Left            =   3600
      TabIndex        =   47
      Top             =   720
      Width           =   1440
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "（設定日付"
      Height          =   255
      Index           =   6
      Left            =   5160
      TabIndex        =   46
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "-"
      Height          =   255
      Index           =   5
      Left            =   4200
      TabIndex        =   45
      Top             =   1140
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "-"
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   44
      Top             =   1140
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "-"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   43
      Top             =   1140
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "標準入庫棚"
      Height          =   240
      Index           =   2
      Left            =   600
      TabIndex        =   42
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "品 名"
      Height          =   240
      Index           =   1
      Left            =   7200
      TabIndex        =   41
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "品番（外部）"
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   40
      Top             =   720
      Width           =   1440
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


Private DEF_GOODS_F As String * 1       '2009.01.08



Private MENU_NO     As String * 2       'ﾒﾆｭｰ№　   2016.01.15
Private RIRK_ID     As String * 2       '要因　     2016.01.15
Private MEMO        As String           'メモ       2016.01.15


'Private Const LAST_UPDATE_DAY$ = "[F101051] 2016.10.27 09:30"
'Private Const LAST_UPDATE_DAY$ = "[F101051] 2017.04.18 09:45"
'Private Const LAST_UPDATE_DAY$ = "品目マスタメンテナンス[F101051] 2019.06.28 17:00" 'タイトルバー変更
Private Const LAST_UPDATE_DAY$ = "品目マスタメンテナンス[F101051] 2019.11.06 9:30 データ出力品番 trim対応"

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
            
Debug.Print (StrConv(ITEMREC.HIN_GAI, vbUnicode))
            
            Case BtErrEOF
                Exit Function
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                List_Disp = True
                Exit Function
        End Select
        
        
If Left(StrConv(ITEMREC.HIN_GAI, vbUnicode), 1) = "'" Then
Debug.Print
End If
        
        Edit = StrConv(ITEMREC.HIN_GAI, vbUnicode) & " " & StrConv(ITEMREC.HIN_NAI, vbUnicode) & " " & StrConv(ITEMREC.HIN_NAME, vbUnicode) & " "
        Edit = Edit & StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                      StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode) & " "
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

    '2009.04.28 42-->44
    '2010.12.09 44-->45
    '2011.06.30 45-->46
    
    
    '2011.06.30 46-->47
    
    
    '2011.10.02 47-->48
    For i = 1 To 48
        Text(i).Text = ""
    Next i

End Sub

'                                       入力項目のエラーチェック
Private Function Del_Chk() As Integer
            
Dim RetBuf As String
Dim i As Integer
Dim sts As Integer


    Del_Chk = False
    
    If Len(RTrim(Text(0).Text)) = 0 Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text(0).SelStart = 0
        Text(0).SelLength = Len(Text(0).Text)
        Text(0).SetFocus
        Del_Chk = True
        Exit Function
    End If
    
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
            
Dim RetBuf  As String
Dim i       As Integer
Dim sts     As Integer
Dim StrWk1  As String
Dim StrWk2  As String
Dim StrWk3  As String


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
'2006.02.14        For i = 4 To 6
'2006.02.14            If Not IsNumeric(Text(i).Text) Then
'2006.02.14                Beep
'2006.02.14                MsgBox "入力した項目はエラーです。"
'2006.02.14                Text(i).SetFocus
'2006.02.14                Err_Chk = True
'2006.02.14                Exit Function
'2006.02.14            Else
'2006.02.14                Text(i).Text = Format(CInt(Text(i).Text), "00")
'2006.02.14            End If
'2006.02.14        Next i
'2006.02.14        Call UniCode_Conv(K0_SOKO.Soko_No, Text(3).Text)
'2006.02.14        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
'2006.02.14        Select Case sts
'2006.02.14            Case BtNoErr
'2006.02.14                If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG$ Then
'2006.02.14                    If StrConv(SOKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
'2006.02.14                        Beep
'2006.02.14                        MsgBox "入力した項目はエラーです。（混載エラー）"
'2006.02.14                        Text(3).SetFocus
'2006.02.14                        Err_Chk = True
'2006.02.14                        Exit Function
'2006.02.14                    End If
'2006.02.14                End If
'2006.02.14            Case BtErrKeyNotFound
'2006.02.14                Beep
'2006.02.14                MsgBox "入力した項目はエラーです。（未登録エラー）"
'2006.02.14                Text(3).SetFocus
'2006.02.14                Err_Chk = True
'2006.02.14                Exit Function
'2006.02.14            Case Else
'2006.02.14                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
'2006.02.14                Err_Chk = SYS_ERR
'2006.02.14                Exit Function
'2006.02.14        End Select
'2006.02.14                                                '棚マスタ読み込み
'2006.02.14        Call UniCode_Conv(K0_TANA.Soko_No, Text(3).Text)
'2006.02.14        Call UniCode_Conv(K0_TANA.Retu, Text(4).Text)
'2006.02.14        Call UniCode_Conv(K0_TANA.Ren, Text(5).Text)
'2006.02.14        Call UniCode_Conv(K0_TANA.Dan, Text(6).Text)
'2006.02.14        sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
'2006.02.14        Select Case sts
'2006.02.14            Case BtNoErr
'2006.02.14            Case BtErrKeyNotFound
'2006.02.14                Beep
'2006.02.14                MsgBox "入力した項目はエラーです。（未登録エラー）"
'2006.02.14                Text(3).SetFocus
'2006.02.14                Err_Chk = True
'2006.02.14                Exit Function
'2006.02.14            Case Else
'2006.02.14                Call File_Error(sts, BtOpGetEqual, "棚マスタ")
'2006.02.14                Err_Chk = SYS_ERR
'2006.02.14                Exit Function
'2006.02.14        End Select
    
    
        For i = 3 To 6                                          '2017.04.18
            If Len(Text(i).Text) <> 2 Then                      '2017.04.18
                MsgBox "棚番は２桁で入力して下さい。"           '2017.04.18
                Text(i).SetFocus                                '2017.04.18
                Err_Chk = True                                  '2017.04.18
                Exit Function                                   '2017.04.18
           End If                                               '2017.04.18
        Next i                                                  '2017.04.18
    End If                                                      '2017.04.18


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
            Text(29).SetFocus
            Err_Chk = True
            Exit Function
        End If

        If Not IsNumeric(Text(29).Text) Then
            Beep
            MsgBox "入力した項目はエラーです。"
            Text(29).SetFocus
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
'        Text(31).Text = "0"            2009.01.08
        Text(31).Text = DEF_GOODS_F     '2009.01.08
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


                                        '収単／担当者
''    If Text(37).Text <> "" Then
''        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN05_CD)
''        Call UniCode_Conv(K0_P_CODE.C_Code, Text(37).Text)
''        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
''        Select Case sts
''            Case BtNoErr
''            Case BtErrKeyNotFound
''                Beep
''                MsgBox "入力した項目はエラーです。（未登録エラー）"
''                Text(37).SetFocus
''                Err_Chk = True
''                Exit Function
''            Case Else
''                Call File_Error(sts, BtOpGetEqual, "コードマスタ")
''                Err_Chk = SYS_ERR
''                Exit Function
''        End Select
''    End If
                                        'S2在庫
    If Text(38).Text = "" Then
        Text(38).Text = "0"
    Else
        If Not IsNumeric(Text(38).Text) Then
            Text(38).Text = "0"
        End If
    End If

                                        'P2在庫
    If Text(39).Text = "" Then
        Text(39).Text = "0"
    Else
        If Not IsNumeric(Text(39).Text) Then
            Text(39).Text = "0"
        End If
    End If



    '商品化請求ﾌﾗｸﾞ 2009.04.28
            
    If Trim(Text(44).Text) = "" Then
        Text(44).Text = "0"
    End If
                
    If Text(44).Text < "0" Or Text(44).Text > "3" Then
                
        MsgBox "入力した項目はエラーです。(商品化請求ﾌﾗｸﾞ)"
        Text(44).SetFocus
        Err_Chk = True
        Exit Function
    End If


    '入り数 2010.12.09

    If Trim(Text(45).Text) = "" Then
    Else
        If Not IsNumeric(Trim(Text(45).Text)) Then
        
            MsgBox "入力した項目はエラーです。(入数（出荷確認計算用）)"
            Text(45).SetFocus
            Err_Chk = True
            Exit Function
        End If
    End If


    '「商品化計画」除外ﾌﾗｸﾞ 2011.06.30
    If Trim(Text(46).Text) <> "" And Trim(Text(46).Text) <> "1" Then
        MsgBox "入力した項目はエラーです。(「商品化計画」除外ﾌﾗｸﾞ)"
        Text(46).SetFocus
        Err_Chk = True
        Exit Function
    End If
    

    '生産ロット数 2011.07.16
    If Trim(Text(47).Text) = "" Then
    Else
        If Not IsNumeric(Trim(Text(47).Text)) Then
        
            MsgBox "入力した項目はエラーです。(生産ロット)"
            Text(47).SetFocus
            Err_Chk = True
            Exit Function
        End If
    End If
    '生産ロット数 2011.07.16


    '[商品化工数] 2011.10.02
    If Trim(Text(48).Text) = "" Then
    Else
    
        If Not IsNumeric(Trim(Text(48).Text)) Then
        
            MsgBox "入力した項目はエラーです。([商品化工数] )"
            Text(48).SetFocus
            Err_Chk = True
            Exit Function
        End If
    End If
    '[商品化工数] 2011.10.02


End Function

Private Sub Item_Dsp()
Dim sts         As Integer
Dim Work_Date   As String * 8
Dim RetBuf      As String

    Text(0).Text = RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
    Text(1).Text = RTrim(StrConv(ITEMREC.HIN_NAI, vbUnicode))
    Text(2).Text = RTrim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
    Text(3).Text = Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode))
    Text(4).Text = Trim(StrConv(ITEMREC.ST_RETU, vbUnicode))
    Text(5).Text = Trim(StrConv(ITEMREC.ST_REN, vbUnicode))
    Text(6).Text = Trim(StrConv(ITEMREC.ST_DAN, vbUnicode))


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

    Text(29).Text = Trim(StrConv(ITEMREC.JAN_CODE, vbUnicode))
    Text(30).Text = Trim(StrConv(ITEMREC.HIN_CHANGE, vbUnicode))
    Text(31).Text = Trim(StrConv(ITEMREC.GOODS_KBN, vbUnicode))
    Text(32).Text = Trim(StrConv(ITEMREC.PACKING_NO, vbUnicode))
    If IsNumeric(StrConv(ITEMREC.AVE_SYUKA, vbUnicode)) Then
        Text(33).Text = Format(CDbl(StrConv(ITEMREC.AVE_SYUKA, vbUnicode)), "#0.0")
    Else
        Text(33).Text = ""

    End If


    Text(34).Text = Trim(StrConv(ITEMREC.GLICS1_TANA, vbUnicode))
    Text(35).Text = Trim(StrConv(ITEMREC.GLICS2_TANA, vbUnicode))
    Text(36).Text = Trim(StrConv(ITEMREC.GLICS3_TANA, vbUnicode))

    Text(37).Text = Trim(StrConv(ITEMREC.S_TANTO, vbUnicode)) '収単／担当者
                                    'ﾗﾍﾞﾙ貼り計上なし
    If StrConv(ITEMREC.G_LABEL_NON, vbUnicode) = P_G_LABEL_ON Then
        Check1(0).Value = vbUnchecked
    Else
        Check1(0).Value = vbChecked
    End If

    If IsNumeric(StrConv(ITEMREC.G_S2_ZAI_QTY, vbUnicode)) Then
        Text(38).Text = Format(CLng(StrConv(ITEMREC.G_S2_ZAI_QTY, vbUnicode)), "#0")
    Else
        Text(38).Text = "0"
    End If

    If IsNumeric(StrConv(ITEMREC.G_P2_ZAI_QTY, vbUnicode)) Then
        Text(39).Text = Format(CLng(StrConv(ITEMREC.G_P2_ZAI_QTY, vbUnicode)), "#0")
    Else
        Text(39).Text = "0"
    End If

    '2007.06.06
    Text(40).Text = Trim(StrConv(ITEMREC.MAKER_CODE, vbUnicode))
    Text(41).Text = Trim(StrConv(ITEMREC.MAKER_NAME, vbUnicode))
    
    
    '2007.06.05
    Text(42).Text = Trim(StrConv(ITEMREC.K_KEITAI, vbUnicode))

    '2009.04.17
    Text(43).Text = Trim(StrConv(ITEMREC.INSP_MESSAGE, vbUnicode))

    '2009.04.28
    Text(44).Text = Trim(StrConv(ITEMREC.S_SEIKYU_F, vbUnicode))


Text1.Text = StrConv(ITEMREC.STAT, vbUnicode)


    '2010.07.20
    Text2(0).Text = StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)
    Text2(1).Text = StrConv(ITEMREC.TORI_GEN_GENSANKOKU, vbUnicode)
    Text2(2).Text = StrConv(ITEMREC.TORI_SHIIRE_WORK_CENTER, vbUnicode)

    
    '2010.07.27
    Text2(3).Text = StrConv(ITEMREC.KANKYO_KBN, vbUnicode)
    Text2(4).Text = StrConv(ITEMREC.KANKYO_KBN_ST, vbUnicode)
    Text2(5).Text = StrConv(ITEMREC.KANKYO_KBN_SURYO, vbUnicode)

    Text2(6).Text = StrConv(ITEMREC.INS_TANTO, vbUnicode)
    Text2(7).Text = StrConv(ITEMREC.Ins_DateTime, vbUnicode)



    '2010.12.09
    If IsNumeric(StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode)) Then
        Text(45).Text = Format(Val(StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode)), "#0")
    Else
        Text(45).Text = ""
    End If


    '2011.06.30
    If StrConv(ITEMREC.GOODS_OUT_F, vbUnicode) < " " Then
        Call UniCode_Conv(ITEMREC.GOODS_OUT_F, "")
    End If
    Text(46).Text = StrConv(ITEMREC.GOODS_OUT_F, vbUnicode)
    '2011.06.30


    '2011.07.16
    If IsNumeric(StrConv(ITEMREC.SEI_LOT, vbUnicode)) Then
        Text(47).Text = Format(Val(StrConv(ITEMREC.SEI_LOT, vbUnicode)), "#0")
    Else
        Text(47).Text = ""
    End If
    '2011.07.16


    '2011.10.02
    If IsNumeric(StrConv(ITEMREC.PLN_KOUSU, vbUnicode)) Then
        Text(48).Text = Format(Val(StrConv(ITEMREC.PLN_KOUSU, vbUnicode)), "#0.00")
    Else
        Text(48).Text = ""
    End If
    '2011.10.02


    Text(49).Text = Trim(StrConv(ITEMREC.NAI_BUHIN, vbUnicode))         '2014.07.04
    Text(50).Text = Trim(StrConv(ITEMREC.GAI_BUHIN, vbUnicode))         '2014.07.04

End Sub

Private Function Update_Proc() As Integer
'品目マスタの追加／訂正
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer
Dim Sv_Naigai       As String * 1
Dim RetBuf          As String
Dim Edit            As String
Dim i               As Integer

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
                                            
                                            
                                            
                                            
                                            
                                            
                                            
    If com = BtOpInsert Then
        
        
        Call Rclr_ITEMREC               '2013.06.13
        
        Call UniCode_Conv(ITEMREC.BEF_SOKO, "")
        Call UniCode_Conv(ITEMREC.BEF_RETU, "")
        Call UniCode_Conv(ITEMREC.BEF_REN, "")
        Call UniCode_Conv(ITEMREC.BEF_DAN, "")

        Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")
        Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")
        Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")
        
'2005.11.15 DEL     Call UniCode_Conv(ITEMREC.LOCK_F, LOCK_OFF)
'2005.11.15 DEL     Call UniCode_Conv(ITEMREC.WEL_ID, "")
'2005.11.15 DEL     Call UniCode_Conv(ITEMREC.PRG_ID, "")
        Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")
        Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "00000000")
        Call UniCode_Conv(ITEMREC.FILLER, "")
    
'2005.11.15 DEL     Call UniCode_Conv(ITEMREC.SIZAI_CD, "")
'2005.11.15 DEL     Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")
        Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")

    
 '2005.11.15 DEL    Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")
 '2005.11.15 DEL    Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")
 '2005.11.15 DEL    Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")
    
    
    
    
        '業務管理項目
        Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")
        Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, "")
        Call UniCode_Conv(ITEMREC.G_SYUSHI, "")
        Call UniCode_Conv(ITEMREC.G_KUMITATE, "")
        Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")
        Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")
        For i = 0 To 2
        
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).CODE, "")
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA, "")
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, "")
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LOT, "")
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LEAD_TIME, "")
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, "")
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, "")
        
        Next i
        Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")
        Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "")
        Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")
        Call UniCode_Conv(ITEMREC.G_LABEL_NON, "")
        Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, "")

        
        '2008.11.10
        Call UniCode_Conv(ITEMREC.ZAIKO_F, P_ZAIKO_F_ON)
        
        
        '商品ラベル項目
        Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")
        Call UniCode_Conv(ITEMREC.L_BIKOU, "")
        Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")
        Call UniCode_Conv(ITEMREC.L_KISHU1, "")
        Call UniCode_Conv(ITEMREC.L_KISHU2, "")
        Call UniCode_Conv(ITEMREC.L_KISHU3, "")
        Call UniCode_Conv(ITEMREC.L_PAPER, "")
        Call UniCode_Conv(ITEMREC.L_PLASTIC, "")
        Call UniCode_Conv(ITEMREC.L_URIKIN1, "")
        Call UniCode_Conv(ITEMREC.L_URIKIN2, "")
        Call UniCode_Conv(ITEMREC.L_URIKIN3, "")
        Call UniCode_Conv(ITEMREC.L_LABEL, "")
        Call UniCode_Conv(ITEMREC.L_MAISU, "")
        Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")
        Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")
        Call UniCode_Conv(ITEMREC.L_BIKOU3, "")
        Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")
        Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")
        Call UniCode_Conv(ITEMREC.L_TANA1, "")
        Call UniCode_Conv(ITEMREC.L_TANA2, "")
        
        
        
        
        
'*------------------------------------------ 2008.08.26 新規追加項目一式 ▽
                
        Call UniCode_Conv(ITEMREC.S_TANTO, "")                  '収単／担当者コード
        Call UniCode_Conv(ITEMREC.ZAIKO_F, "1")                  '在庫管理対象有無 1:対象 0:対象外

        Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '           機種(2)

        Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "")          '           前月在庫数量
        Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, "")         '           最終出荷数

        Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, "")             'GLICS在庫(S2) 袋井用
        Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, "")             'GLICS在庫(P2) 袋井用

        Call UniCode_Conv(ITEMREC.K_KEITAI, "")                 '個装形態


        Call UniCode_Conv(ITEMREC.UNIT_BUHIN, "")               'ﾕﾆｯﾄ部品区分
        Call UniCode_Conv(ITEMREC.NAI_BUHIN, "")                '国内供給部品区分   2006.07.28
        Call UniCode_Conv(ITEMREC.GAI_BUHIN, "")                '海外供給部品区分   2006.07.28
        Call UniCode_Conv(ITEMREC.HYO_TANKA, "")                '標準単価   2006.07.28

        Call UniCode_Conv(ITEMREC.LAST_CODE, "")                '最終仕入先コード   2007.05.29
        Call UniCode_Conv(ITEMREC.LAST_TANKA, "")               '最終仕入単価       2007.05.29

        Call UniCode_Conv(ITEMREC.MAKER_CODE, "")               'ﾒｰｶｰｺｰﾄﾞ           2007.06.06
        Call UniCode_Conv(ITEMREC.MAKER_NAME, "")               'ﾒｰｶｰ名称           2007.06.06


        Call UniCode_Conv(ITEMREC.L_MARK, "")                   '再梱包ﾏｰｸ          2007.11.08

        Call UniCode_Conv(ITEMREC.SAI_SU, "")                   '才数               2008.02.14

        Call UniCode_Conv(ITEMREC.D_KEISHIKI, "")               '形式               2008.02.14
        Call UniCode_Conv(ITEMREC.D_MATERIAL, "")               '材質               2008.02.14
        Call UniCode_Conv(ITEMREC.D_THICKNESS, "")              'ﾀﾞﾝﾎﾞｰﾙ　厚さ      2008.02.14


        Call UniCode_Conv(ITEMREC.D_SIZE_W, "")                 'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（W）   2008.02.14
        Call UniCode_Conv(ITEMREC.D_SIZE_D, "")                 'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（D）   2008.02.14
        Call UniCode_Conv(ITEMREC.D_SIZE_H, "")                 'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（H）   2008.02.14

        Call UniCode_Conv(ITEMREC.D_PRINT, "")                  '印刷する／しない   2008.02.14
    

        Call UniCode_Conv(ITEMREC.S_KOUSU, "")                  '商品化　工数       2008.02.14

        Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, "")            '商品化　工数原価   2008.02.14
        Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "")            '商品化　工数売価   2008.02.14
        Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, "")         '商品化　単価設定日 2008.02.14


        Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, "")           '商品化　資材原価   2008.02.14
        Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "")           '商品化　資材売価   2008.02.14
        Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, "")        '商品化　単価設定日 2008.02.14


        Call UniCode_Conv(ITEMREC.SE_USOU_F, "")                '輸送箱　出力ﾌﾗｸﾞ   2008.02.14

        Call UniCode_Conv(ITEMREC.USE_TAPE_KIND, "")            '使用テープ種類     2008.02.14
        Call UniCode_Conv(ITEMREC.USE_TAPE_LNG, "")             '使用テープ長       2008.02.14

        Call UniCode_Conv(ITEMREC.H_TANA_MAKE, "")              '棚番マーク         2008.04.02


        Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, "")            '請求単価　メモ     2008.04.15


        Call UniCode_Conv(ITEMREC.GENSANKOKU, "")               '原産国             2008.06.11



        Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, "")            '外装単価 9(8)V99   2008.06.12
        Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, "")         'PPSC加工単価9(8)   2008.06.12
        Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, "")           'BU加工単価9(8)     2008.06.12


        Call UniCode_Conv(ITEMREC.SEI_LOT, "")                  '生産ロット         2008.07.07
        Call UniCode_Conv(ITEMREC.SEI_RATE, "")                 '分レート           2008.07.07
        Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "")              '集合梱包           2008.07.07


        Call UniCode_Conv(ITEMREC.SEI_TANKA_TANTO, "")          '単価設定担当者     2008.07.09

        Call UniCode_Conv(ITEMREC.SHIMUKE_CODE, "")             '仕向け先           2008.07.09

        Call UniCode_Conv(ITEMREC.SEI_KBN, "")                  '請求区分           2008.07.16

        Call UniCode_Conv(ITEMREC.SEI_LABEL_QTY, "")            'ラベル貼り枚数     2008.07.19

        Call UniCode_Conv(ITEMREC.SEI_SZI_CNT, "")              '資材件数     　    2008.08.20追加
        Call UniCode_Conv(ITEMREC.SEI_DKN_CNT, "")              '同梱件数           2008.08.20追加
         
'*------------------------------------------ 2008.08.26 新規追加項目一式 △
        '↓2009.02.20
        For i = 0 To 9
            Call UniCode_Conv(ITEMREC.BEF_KOUTEI(i).BEF_KOUTEI, "")
            Call UniCode_Conv(ITEMREC.MAIN_KOUTEI(i).MAIN_KOUTEI, "")
            Call UniCode_Conv(ITEMREC.AFT_KOUTEI(i).AFT_KOUTEI, "")

        Next i


        Call UniCode_Conv(ITEMREC.SE_IO_TANKA_No, "")
        '↑2009.02.20
        
                
        
        
        Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "")
        Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "")
        Call UniCode_Conv(ITEMREC.BEF_S_GAISO_TANKA, "")
        Call UniCode_Conv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, "")
        Call UniCode_Conv(ITEMREC.BEF_S_BU_KAKO_KOSU, "")
    
        Call UniCode_Conv(ITEMREC.M_BIKOU, "")
        Call UniCode_Conv(ITEMREC.SHIYOU_NO, "")
        Call UniCode_Conv(ITEMREC.MITSUMORI_KBN, "")
        Call UniCode_Conv(ITEMREC.TANKA_KIRIKAE_DT, "")
        Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, "")
        
        
        
        
        
        '追加項目
        Call UniCode_Conv(ITEMREC.INS_TANTO, "F105051")
        Call UniCode_Conv(ITEMREC.Ins_DateTime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
        
        
        '共通項目
        Call UniCode_Conv(ITEMREC.FILLER, "")
        Call UniCode_Conv(ITEMREC.UPD_TANTO, "")
        Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    
    
    End If
                                            
                                            
                                            
                                            
                                            
                                            
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


    
    Call UniCode_Conv(ITEMREC.HIN_NAI, Text(1).Text)
    Call UniCode_Conv(ITEMREC.BIKOU_SOKO, Text(20).Text)
    Call UniCode_Conv(ITEMREC.BIKOU_TANA, Text(21).Text)
    Call UniCode_Conv(ITEMREC.SAMPLE_QTY, Format(CInt(Text(22).Text), "0"))
    Call UniCode_Conv(ITEMREC.JAN_CODE, Text(29).Text)
    Call UniCode_Conv(ITEMREC.HIN_CHANGE, Text(30).Text)
    Call UniCode_Conv(ITEMREC.GOODS_KBN, Text(31).Text)
    Call UniCode_Conv(ITEMREC.PACKING_NO, Text(32).Text)
            
    Call UniCode_Conv(ITEMREC.GLICS1_TANA, Text(34).Text)
    Call UniCode_Conv(ITEMREC.GLICS2_TANA, Text(35).Text)
    Call UniCode_Conv(ITEMREC.GLICS3_TANA, Text(36).Text)
            
    Call UniCode_Conv(ITEMREC.S_TANTO, Text(37).Text)       '収単／担当者
            
        
    If Check1(0).Value = vbChecked Then                                   'ラベル貼り
        Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_OFF)
    Else
        Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_ON)
    End If
            
            
    Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, Format(CLng(Text(38).Text), "00000000"))
    Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, Format(CLng(Text(39).Text), "00000000"))
            
    '2007.06.06
    Call UniCode_Conv(ITEMREC.MAKER_CODE, Text(40).Text)
    Call UniCode_Conv(ITEMREC.MAKER_NAME, Text(41).Text)
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    '2010.08.17
    Call UniCode_Conv(ITEMREC.UPD_TANTO, "05051")
    '2010.08.17
    
    
    Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    
    
    
    Call UniCode_Conv(ITEMREC.STAT, Text1.Text)
    
    
    '2009.04.28
    Call UniCode_Conv(ITEMREC.S_SEIKYU_F, Text(44).Text)
    
    '2009.04.17
    Call UniCode_Conv(ITEMREC.INSP_MESSAGE, Text(43).Text)
    
    
    '2010.12.09
    If Text(45).Text = "" Then
        Call UniCode_Conv(ITEMREC.GAISO_IRI_QTY, "")
    Else
        Call UniCode_Conv(ITEMREC.GAISO_IRI_QTY, Format(Val(Text(45).Text), "00000000"))
    End If
    
    '2011.06.30
    If Trim(Text(46).Text) = "1" Then
        Call UniCode_Conv(ITEMREC.GOODS_OUT_F, Text(46).Text)
    Else
        Call UniCode_Conv(ITEMREC.GOODS_OUT_F, "")
    End If
    '2011.06.30
    
    
    
    '2011.07.16
    If IsNumeric(Text(47).Text) Then
        Call UniCode_Conv(ITEMREC.SEI_LOT, Format(Val(Text(47).Text), "00000000"))
    Else
        Call UniCode_Conv(ITEMREC.SEI_LOT, "")
    End If
    '2011.07.16
    
    
    '2011.10.02
    If IsNumeric(Text(48).Text) Then
        Call UniCode_Conv(ITEMREC.PLN_KOUSU, Format(Val(Text(48).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEMREC.PLN_KOUSU, "")
    End If
    '2011.10.02
    
    
    
    Call UniCode_Conv(ITEMREC.NAI_BUHIN, Text(49).Text) '2014.07.04
    Call UniCode_Conv(ITEMREC.GAI_BUHIN, Text(50).Text) '2014.07.04
    
    
    
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
        
        
    If com = BtOpInsert Then
    
        Call LOG_OUT(LOG_F, "F101051 INS HIN_GAI=" & StrConv(ITEMREC.HIN_GAI, vbUnicode) & _
                            " S_KOUSU_SET_DATE=" & StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode) & _
                            " S_SHIZAI_SET_DATE=" & StrConv(ITEMREC.S_SHIZAI_SET_DATE, vbUnicode) & _
                            " SEI_TANKA_TANTO=" & StrConv(ITEMREC.SEI_TANKA_TANTO, vbUnicode) & _
                            " S_GAISO_TANKA=" & StrConv(ITEMREC.S_GAISO_TANKA, vbUnicode) & _
                            " S_PPSC_KAKO_KOSU=" & StrConv(ITEMREC.S_PPSC_KAKO_KOSU, vbUnicode) & _
                            " S_BU_KAKO_KOSU=" & StrConv(ITEMREC.S_BU_KAKO_KOSU, vbUnicode))

    
    
    End If
        
        
        
        
        
        
        
        
    Edit = StrConv(ITEMREC.HIN_GAI, vbUnicode) & " " & StrConv(ITEMREC.HIN_NAI, vbUnicode) & " " & _
           StrConv(ITEMREC.HIN_NAME, vbUnicode) & " "
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
    
    
        '作業ﾛｸﾞ出力    '2016.01.15
        If Trim(MENU_NO) <> "" Then
        
            If P_SAGYO_LOG_OUTPUT_PROC("", _
                                        "", _
                                        Last_JGYOBU, _
                                        Sv_Naigai, _
                                        MENU_NO, _
                                        RIRK_ID, _
                                        Text(0).Text, _
                                        0, _
                                        0, _
                                        "", _
                                        "", , , , , , , , , , MEMO) Then
                Exit Function
            End If
    
        End If
                                        
    End If
                                        'リストボックス削除
    For i = 0 To List1.ListCount - 1
        If RTrim(Left$(List1.List(i), 13)) = RTrim(Text(0).Text) Then
            List1.RemoveItem i
        End If
    Next i
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
                                            
            sts = Err_Chk()                 'エラーチェック
                
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
'    PrintForm                      2017.04.18
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
        
    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            Me.Caption = "(" + RTrim(JGYOBU_T(i).NAME) + ")" & LAST_UPDATE_DAY
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
                                
    Unload SubMenu(i)
                                
                                'リストボックス最大表示件数表示
'    If GetIni(App.EXEName, "LISTMAX", "SYS", c) Then               '2016.01.15
    If GetIni(App.EXEName, "LISTMAX", App.EXEName, c) Then          '2016.01.15
        Beep
        MsgBox "最大表示件数の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LIST_MAX = CInt(RTrim(c))
                                
                                
                                '商品化ﾌﾗｸﾞﾃﾞﾌｫﾙﾄ   2008.01.08
'    If GetIni(App.EXEName, "DEF_GOODS_F", "SYS", c) Then           '2016.01.15
    If GetIni(App.EXEName, "DEF_GOODS_F", App.EXEName, c) Then      '2016.01.15
        DEF_GOODS_F = "0"
    Else
        If Trim(c) = "1" Then
            DEF_GOODS_F = "1"
        Else
            DEF_GOODS_F = "0"
        End If
    End If
                                
'>>>>>>>>>>>>>>>>>>>>   作業ログ情報    2016.01.15
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
'>>>>>>>>>>>>>>>>>>>>   作業ログ情報    2016.01.15
                                
                                
                                
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
                                
'    If ITEM_B_Open(BtOpenNomal) Then
'        Unload Me
'    End If
'    If ITEM_C_Open(BtOpenNomal) Then
'        Unload Me
'    End If
                                
                                '個装箱マスタＯＰＥＮ
    If PACKING_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '作業ログＯＰＥＮ   '2016.01.15
    If P_SAGYO_LOG_Open(BtOpenNomal) Then
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
                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "コードマスタ")
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

'2010.12.09
''    Call UniCode_Conv(K0_ITEM.HIN_GAI, Mid$(List1.List(List1.ListIndex), 1, 13))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Mid$(List1.List(List1.ListIndex), 1, 20))
'2010.12.09
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
'2010.12.09
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Mid$(List1.List(List1.ListIndex), 1, 13))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Mid$(List1.List(List1.ListIndex), 1, 20))
'2010.12.09
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
    
    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1010511.Caption = "品目マスタメンテナンス（削除機能付き）（" + RTrim(JGYOBU_T(Index).NAME) + ") " & LAST_UPDATE_DAY
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)
End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = RTrim(Text(Index).Text)
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
                    
                    Text(0).Text = StrConv(RTrim(Text(0).Text), vbUpperCase)
                                
                    
                    If Len(Text(Index).Text) <> 0 Then
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
                    End If
                Case 1
                    
                    Text(1).Text = StrConv(RTrim(Text(1).Text), vbUpperCase)
                    
                    If Len(Text(0).Text) = 0 Then
                                                '内部品番で読み込み
                        Call UniCode_Conv(K2_ITEM.JGYOBU, Last_JGYOBU)
                        If Combo(0).Text = NAIGAI1$ Then
                            Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI_NAI$)
                        Else
                            Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI_GAI$)
                        End If
                        Call UniCode_Conv(K2_ITEM.HIN_NAI, RTrim(Text(1).Text))
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                        Select Case sts
                            Case BtNoErr
                                Call Item_Dsp
                                If List_Disp() Then
                                    Unload Me
                                End If
                            Case BtErrKeyNotFound
                                Call UniCode_Conv(K2_ITEM.JGYOBU, Last_JGYOBU)
                                If Combo(0).Text = NAIGAI1$ Then
                                    Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI_NAI$)
                                Else
                                    Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI_GAI$)
                                End If
                                Call UniCode_Conv(K2_ITEM.HIN_NAI, RTrim(Text(1).Text))
                                sts = BTRV(BtOpGetGreaterEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                                If sts = BtNoErr Then
                                    Text(0).Text = RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                End If
                                If List_Disp() Then
                                    Unload Me
                                End If
                                MsgBox "入力したコードは登録されていません。"
                                Call Item_Dsp
                                Exit Sub
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                Unload Me
                        End Select
                    End If
                Case 29             'Janコード
                    
                    If Len(Trim(Text(0).Text)) = 0 Then
                        If Len(Trim(Text(29).Text)) <> 0 Then
                            Call UniCode_Conv(K4_ITEM.JGYOBU, Last_JGYOBU)
                            If Combo(0).Text = NAIGAI1$ Then
                                Call UniCode_Conv(K4_ITEM.NAIGAI, NAIGAI_NAI$)
                            Else
                                Call UniCode_Conv(K4_ITEM.NAIGAI, NAIGAI_GAI$)
                            End If
                            Call UniCode_Conv(K4_ITEM.JAN_CODE, RTrim(Text(29).Text))
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K4_ITEM, Len(K4_ITEM), 4)
                            Select Case sts
                                Case BtNoErr
                                    Call Item_Dsp
                                Case BtErrKeyNotFound
                                    MsgBox "入力したコードは登録されていません。"
                                    Call Item_Dsp
                                    Exit Sub
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                    Unload Me
                            End Select
                        End If
                    End If
                Case 30             '読替えコード
                    
                    If Len(Trim(Text(0).Text)) = 0 Then
                        If Len(Trim(Text(30).Text)) <> 0 Then
                            Call UniCode_Conv(K5_ITEM.JGYOBU, Last_JGYOBU)
                            If Combo(0).Text = NAIGAI1$ Then
                                Call UniCode_Conv(K5_ITEM.NAIGAI, NAIGAI_NAI$)
                            Else
                                Call UniCode_Conv(K5_ITEM.NAIGAI, NAIGAI_GAI$)
                            End If
                            Call UniCode_Conv(K5_ITEM.HIN_CHANGE, RTrim(Text(30).Text))
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K5_ITEM, Len(K5_ITEM), 5)
                            Select Case sts
                                Case BtNoErr
                                    Call Item_Dsp
                                Case BtErrKeyNotFound
                                    MsgBox "入力したコードは登録されていません。"
                                    Call Item_Dsp
                                    Exit Sub
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                    Unload Me
                            End Select
                        End If
                    End If
            
            
            
            
            
                Case 44             '商品化請求ﾌﾗｸﾞ 2009.04.28
            
                    If Trim(Text(44).Text) = "" Then
                        Text(44).Text = "0"
                    End If
                                
                    If Text(44).Text < "0" Or Text(44).Text > "3" Then
                                
                        MsgBox "入力した項目はエラーです。(商品化請求ﾌﾗｸﾞ)"
                        Exit Sub
                    End If
            
            
            
            
            
            
            End Select
            'For i = Index + 1 To 44                                '2013.07.30
            For i = Index + 1 To 48                                 '2013.07.30
                'If Text(i).Enabled And Not Text(i).Locked Then     '2013.07.30
                If Text(i).Enabled And Not Text(i).Locked And Text(i).Visible Then
                    Text(i).SetFocus
                    Exit For
                End If
            Next i


    End Select
End Sub



Private Function CSV_OUTPUT_Proc() As Integer

Dim FileNo          As Integer
Dim FileName        As String
Dim ret             As Integer

Dim com             As Integer
Dim sts             As Integer

Dim c               As String * 128

Dim Soko_No         As String * 2

    Call Input_Lock
    
    FileNo = FreeFile
    FileName = ITEM_CSV
    
'    Ret = InStr(1, Trim(fileName), ".") - 1
    
    
    ret = InStrRev(Trim(FileName), ".") - 1
    
    FileName = Left(Trim(FileName), ret) & "_" & Last_JGYOBU & Right(Trim(FileName), Len(Trim(FileName)) - ret)

    On Error GoTo Error_Proc

    Open (FileName) For Output As FileNo
    
'    Write #FileNo, "事業部", "内外", "品番（外部）", "品名", "標準棚番設定日", "標準棚番", "前回棚番", "最終入庫日", "最終出庫日", "品番（内部）", "ホスト倉庫", "ホスト棚番", "サンプル数", "最終入荷日", "最終照合日", "照合時在庫数", "印刷備考", "印刷入り数", "ＪＡＮ", "読み替え", "商品化有無", "個装箱№"

'    Write #FileNo, "内外", "品番（外部）", "品名", "標準棚番設定日", "標準棚番", "前回棚番", "最終入庫日", "最終出庫日", "品番（内部）", "ホスト倉庫", "ホスト棚番", "サンプル数", "最終入荷日", "最終照合日", "照合時在庫数", "印刷備考", "印刷入り数", "ＪＡＮ", "読み替え", "商品化有無", "個装箱№", "G棚番1", "G棚番2", "G棚番3", "在庫対象"
    Write #FileNo, "内外", "品番（外部）", "品名", "標準棚番設定日", "標準棚番", "前回棚番", "最終入庫日", "最終出庫日", "品番（内部）", "ホスト倉庫", "ホスト棚番", "サンプル数", "最終入荷日", "最終照合日", "照合時在庫数", "印刷備考", "印刷入り数", "ＪＡＮ", "読み替え", "商品化有無", "個装箱№", "G棚番1", "G棚番2", "G棚番3", "外装箱入り数"
    
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, "")
    Call UniCode_Conv(K0_ITEM.HIN_GAI, "")

    com = BtOpGetGreaterEqual
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
    
    
'    If Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) = "APB01H413-CU" Or _
'        Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) = "ARB01-C54W9S" Or _
'        Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) = "ARE60-B78" Or _
'        Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) = "AVE39-172-H" Or _
'        Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) = "NC-CF1" Or _
'        Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) = "XTN4+10BFN" Or _
'        Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) = "AZB03-813-0S" Or _
'        Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) = "KZ-JJ112-566" Then



'    If Trim(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode)) >= "20110114" Then
    
    
    
'    If Not IsNumeric(StrConv(ITEMREC.SAI_SU, vbUnicode)) Or Not IsNumeric(StrConv(ITEMREC.KUTI_SU, vbUnicode)) Then
    
    
'        Write #FileNo, StrConv(ITEMREC.JGYOBU, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.NAIGAI, vbUnicode),
        Write #FileNo, Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)),    '2019/11/06 trim
        Write #FileNo, StrConv(ITEMREC.HIN_NAME, vbUnicode),

''        If IsDate(StrConv(ITEMREC.ST_SET_DT, vbUnicode)) Then
''            Write #FileNo, Format(StrConv(ITEMREC.ST_SET_DT, vbUnicode), "YYYY/MM/DD"),
''        Else
''            Write #FileNo, ,
''        End If
        Write #FileNo, StrConv(ITEMREC.ST_SET_DT, vbUnicode),


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

''        If IsDate(StrConv(ITEMREC.LAST_NYU_DT, vbUnicode)) Then
''            Write #FileNo, Format(StrConv(ITEMREC.LAST_NYU_DT, vbUnicode), "YYYY/MM/DD"),
''        Else
''            Write #FileNo, ,
''        End If
        
        Write #FileNo, StrConv(ITEMREC.LAST_NYU_DT, vbUnicode),
        
        
''        If IsDate(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode)) Then
''            Write #FileNo, Format(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode), "YYYY/MM/DD"),
''        Else
''            Write #FileNo, ,
''        End If
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
If Trim(StrConv(ITEMREC.JAN_CODE, vbUnicode)) <> "" Then
Debug.Print Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) & "-" & Trim(StrConv(ITEMREC.JAN_CODE, vbUnicode))
End If
        
        Write #FileNo, StrConv(ITEMREC.HIN_CHANGE, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.GOODS_KBN, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.PACKING_NO, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.GLICS1_TANA, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.GLICS2_TANA, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.GLICS3_TANA, vbUnicode),
        Write #FileNo, StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode),
        
        Write #FileNo, StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode),
    
'        Write #FileNo, "ZAIKO_F=" & StrConv(ITEMREC.ZAIKO_F, vbUnicode),
'        Write #FileNo, StrConv(ITEMREC.G_SYUSHI, vbUnicode),
        
'        Write #FileNo, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode),
'        Write #FileNo, StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode),
        
'        Write #FileNo, StrConv(ITEMREC.G_SYUSHI, vbUnicode),
        
        
'        Write #FileNo, StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode),
'        Write #FileNo, StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode),
        
'        Write #FileNo, "ZAIKO_F=" & StrConv(ITEMREC.ZAIKO_F, vbUnicode),
    
'        Write #FileNo, "L_MARK=" & StrConv(ITEMREC.L_MARK, vbUnicode),
    
    
    
'        Write #FileNo, StrConv(ITEMREC.K_KEITAI, vbUnicode),
        
'        Write #FileNo, StrConv(ITEMREC.SHIMUKE_CODE, vbUnicode),
'        Write #FileNo, StrConv(ITEMREC.SEI_SZI_CNT, vbUnicode),
'        Write #FileNo, StrConv(ITEMREC.SEI_DKN_CNT, vbUnicode),
        
'Dim i As Integer
'        For i = 0 To 9
'            Write #FileNo, StrConv(ITEMREC.BEF_KOUTEI(i).BEF_KOUTEI, vbUnicode),
'        Next i
'        For i = 0 To 9
'            Write #FileNo, StrConv(ITEMREC.MAIN_KOUTEI(i).MAIN_KOUTEI, vbUnicode),
'        Next i
'        For i = 0 To 9
'            Write #FileNo, StrConv(ITEMREC.AFT_KOUTEI(i).AFT_KOUTEI, vbUnicode),
'        Next i
'
'        Write #FileNo, "才＝" & StrConv(ITEMREC.SAI_SU, vbUnicode),
'        Write #FileNo, "口＝" & StrConv(ITEMREC.KUTI_SU, vbUnicode),
'        Write #FileNo, "INS＝" & StrConv(ITEMREC.Ins_DateTime, vbUnicode),
'        Write #FileNo, "梱包＝" & StrConv(ITEMREC.KONPOU_F, vbUnicode),
        
'        Write #FileNo, StrConv(ITEMREC.SE_IO_TANKA_No, vbUnicode)
        
        
'        Write #FileNo, StrConv(ITEMREC.L_PAPER, vbUnicode);
'        Write #FileNo, StrConv(ITEMREC.L_PLASTIC, vbUnicode);
        
        
        
        
        
        
        
'        Call UniCode_Conv(K0_ITEM_C.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
'        Call UniCode_Conv(K0_ITEM_C.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
'        Call UniCode_Conv(K0_ITEM_C.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
'        sts = BTRV(BtOpGetEqual, ITEM_C_POS, ITEM_CREC, Len(ITEM_CREC), K0_ITEM_C, Len(K0_ITEM_C), 0)
'        Select Case sts
'            Case BtNoErr
'                Write #FileNo, StrConv(ITEM_CREC.L_PAPER, vbUnicode);
'                Write #FileNo, StrConv(ITEM_CREC.L_PLASTIC, vbUnicode);
'            Case BtErrKeyNotFound
'                Write #FileNo, "未";
'                Write #FileNo, "未";
'        End Select
        
        
        
        
        
        
'        Call UniCode_Conv(K0_ITEM_B.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
'        Call UniCode_Conv(K0_ITEM_B.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
'        Call UniCode_Conv(K0_ITEM_B.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
'        sts = BTRV(BtOpGetEqual, ITEM_B_POS, ITEM_BREC, Len(ITEM_BREC), K0_ITEM_B, Len(K0_ITEM_B), 0)
'        Select Case sts
'            Case BtNoErr
'                Write #FileNo, StrConv(ITEM_BREC.L_PAPER, vbUnicode);
'                Write #FileNo, StrConv(ITEM_BREC.L_PLASTIC, vbUnicode);
'            Case BtErrKeyNotFound
'                Write #FileNo, "未";
'                Write #FileNo, "未";
'        End Select
        
'Write #FileNo, StrConv(ITEMREC.G_SYUSHI, vbUnicode),
        
        
'Write #FileNo, StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode),
'Write #FileNo, StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode),
        
        
'Write #FileNo, StrConv(ITEMREC.GOODS_KBN, vbUnicode),
'Write #FileNo, StrConv(ITEMREC.NAI_BUHIN, vbUnicode),


'Write #FileNo, "ZAIKO_F=" & StrConv(ITEMREC.ZAIKO_F, vbUnicode),

'Write #FileNo, "SAI_SU=" & StrConv(ITEMREC.SAI_SU, vbUnicode),
'Write #FileNo, "KUTI_SU=" & StrConv(ITEMREC.KUTI_SU, vbUnicode),
        
        
'Write #FileNo, "GENSANKOKU", StrConv(ITEMREC.GENSANKOKU, vbUnicode),
'Write #FileNo, "TORI_GENSANKOKU", StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode),

'Write #FileNo, "NAI_BUHIN=" & StrConv(ITEMREC.NAI_BUHIN, vbUnicode),
'Write #FileNo, "L_KISHU1=" & StrConv(ITEMREC.L_KISHU1, vbUnicode),
'Write #FileNo, "L_KISHU2=" & StrConv(ITEMREC.L_KISHU2, vbUnicode),
'Write #FileNo, "L_KISHU3=" & StrConv(ITEMREC.L_KISHU3, vbUnicode),

'Write #FileNo, "PLUS_KOUSU=" & StrConv(ITEMREC.PLUS_KOUSU, vbUnicode),
'
'Write #FileNo, "INS_TANTO=" & StrConv(ITEMREC.INS_TANTO, vbUnicode),
'Write #FileNo, "INS_DateTime=" & StrConv(ITEMREC.Ins_DateTime, vbUnicode),
        
'Write #FileNo, "UPD_TANTO=" & StrConv(ITEMREC.UPD_TANTO, vbUnicode),
'Write #FileNo, "UPD_DateTime=" & StrConv(ITEMREC.UPD_DATETIME, vbUnicode),
        
'Write #FileNo, "CATE_ST_FUKA=" & StrConv(ITEMREC.CATE_ST_FUKA, vbUnicode),
'Write #FileNo, "CATE_AD_FUKA=" & StrConv(ITEMREC.CATE_AD_FUKA, vbUnicode),
        
'Write #FileNo, StrConv(ITEMREC.GENSANKOKU, vbUnicode),
'Write #FileNo, StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode),
        

'Write #FileNo, StrConv(ITEMREC.INSP_MESSAGE, vbUnicode),
        

'Write #FileNo, "IRI_QTY=" & Format(Val(StrConv(ITEMREC.IRI_QTY, vbUnicode)), "#"),
'Write #FileNo, "GAISO_IRI_QTY=" & Format(Val(StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode)), "#"),
        
        
'Write #FileNo, StrConv(ITEMREC.INS_TANTO, vbUnicode),
'Write #FileNo, StrConv(ITEMREC.Ins_DateTime, vbUnicode),
        
'Write #FileNo, StrConv(ITEMREC.UPD_TANTO, vbUnicode),
'Write #FileNo, StrConv(ITEMREC.UPD_DATETIME, vbUnicode),
        
        
'Write #FileNo, "ZAIKO_F=" & Format(Val(StrConv(ITEMREC.ZAIKO_F, vbUnicode)), "#"),
        
'Write #FileNo, "PLN_SAGYOU_KOUSU=" & Format(Val(StrConv(ITEMREC.PLN_SAGYOU_KOUSU, vbUnicode)), "#0.0"),
        
        
'Write #FileNo, "S_KOUSU_SET_DATE=" & Trim(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode)),
'
'
'Write #FileNo, "S_KOUSU=" & Trim(StrConv(ITEMREC.S_KOUSU, vbUnicode)),
'
'Write #FileNo, "BEF_S_KOUSU=" & Trim(StrConv(ITEMREC.BEF_S_KOUSU, vbUnicode)),
'
'Write #FileNo, "CATE_AD_FUN=" & Trim(StrConv(ITEMREC.CATE_AD_FUN, vbUnicode)),
'
'Write #FileNo, "UPD_TANTO=" & Trim(StrConv(ITEMREC.UPD_TANTO, vbUnicode)),
'
'Write #FileNo, "UPD_DateTime=" & Trim(StrConv(ITEMREC.UPD_DATETIME, vbUnicode)),
'
''2013.03.26 DEBUG
'Write #FileNo, "L_KAISHA_CODE=" & Trim(StrConv(ITEMREC.L_KAISHA_CODE, vbUnicode)),
'Write #FileNo, "L_JGYOBU_CODE=" & Trim(StrConv(ITEMREC.L_JGYOBU_CODE, vbUnicode)),
''2013.03.26 DEBUG
'
'
'Write #FileNo, "JAN_CODE=", StrConv(ITEMREC.JAN_CODE, vbUnicode),
'Write #FileNo, "HIN_CHANGE=", StrConv(ITEMREC.HIN_CHANGE, vbUnicode),
        
        
        
        
        Write #FileNo,
        
        
        
        
'        End If
    
        com = BtOpGetNext
    Loop

    Close #FileNo
    
    Call Input_UnLock
    
    
    Beep
    MsgBox "「" & FileName & "」は正常に出力されました。"


    Exit Function
Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox FileName & "が使用中です。"
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

Private Sub Text_LostFocus(Index As Integer)

    If Index = 0 Or Index = 1 Then
        Text(Index).Text = StrConv(RTrim(Text(Index).Text), vbUpperCase)
    End If

    '>>>>>>>>>>>>   2015.12.15
    If Index = 3 Or Index = 4 Or Index = 5 Or Index = 6 Then
        Text(Index).Text = StrConv(RTrim(Text(Index).Text), vbUpperCase)
    End If
    '>>>>>>>>>>>>   2015.12.15

End Sub
