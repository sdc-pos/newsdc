VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form CNV_SEI00101 
   Caption         =   "[請求システム]見積書情報コンバート処理　2009.01.21"
   ClientHeight    =   9975
   ClientLeft      =   2025
   ClientTop       =   -3210
   ClientWidth     =   15240
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
   ScaleHeight     =   9975
   ScaleWidth      =   15240
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   166
      Left            =   7245
      TabIndex        =   31
      Top             =   1920
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   165
      Left            =   7245
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1680
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   164
      Left            =   9135
      TabIndex        =   33
      Top             =   1920
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   162
      Left            =   9135
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1680
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   163
      Left            =   8190
      TabIndex        =   32
      Top             =   1920
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
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
      Index           =   159
      Left            =   11025
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   720
      Width           =   1905
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
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
      Index           =   158
      Left            =   10605
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   720
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   154
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   164
      TabStop         =   0   'False
      Top             =   9240
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   153
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   163
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   152
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   162
      TabStop         =   0   'False
      Top             =   8760
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   151
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   161
      TabStop         =   0   'False
      Top             =   8520
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   85
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   9240
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   84
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   9000
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   83
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   8760
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   82
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   8520
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   117
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   127
      TabStop         =   0   'False
      Top             =   9240
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   116
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   126
      TabStop         =   0   'False
      Top             =   9000
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   115
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   125
      TabStop         =   0   'False
      Top             =   8760
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   114
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   8520
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   149
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   159
      TabStop         =   0   'False
      Top             =   9240
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   148
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   158
      TabStop         =   0   'False
      Top             =   9000
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   147
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   157
      TabStop         =   0   'False
      Top             =   8760
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   146
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   156
      TabStop         =   0   'False
      Top             =   8520
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   157
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   168
      TabStop         =   0   'False
      Top             =   8640
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   156
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   167
      TabStop         =   0   'False
      Top             =   8280
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   155
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   166
      TabStop         =   0   'False
      Top             =   7920
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   145
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   155
      TabStop         =   0   'False
      Top             =   8280
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   144
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   154
      TabStop         =   0   'False
      Top             =   8040
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   143
      Left            =   9135
      Locked          =   -1  'True
      TabIndex        =   153
      TabStop         =   0   'False
      Top             =   8040
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   142
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   152
      TabStop         =   0   'False
      Top             =   8040
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   141
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   151
      TabStop         =   0   'False
      Top             =   7800
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   140
      Left            =   9135
      Locked          =   -1  'True
      TabIndex        =   150
      TabStop         =   0   'False
      Top             =   7800
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   139
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   149
      TabStop         =   0   'False
      Top             =   7800
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   138
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   148
      TabStop         =   0   'False
      Top             =   7560
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   137
      Left            =   9135
      Locked          =   -1  'True
      TabIndex        =   147
      TabStop         =   0   'False
      Top             =   7560
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   136
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   146
      TabStop         =   0   'False
      Top             =   7560
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   135
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   145
      TabStop         =   0   'False
      Top             =   7320
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   134
      Left            =   9135
      Locked          =   -1  'True
      TabIndex        =   144
      TabStop         =   0   'False
      Top             =   7320
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   133
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   143
      TabStop         =   0   'False
      Top             =   7320
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   132
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   142
      TabStop         =   0   'False
      Top             =   7080
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   131
      Left            =   9135
      Locked          =   -1  'True
      TabIndex        =   141
      TabStop         =   0   'False
      Top             =   7080
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   130
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   140
      TabStop         =   0   'False
      Top             =   7080
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   129
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   139
      TabStop         =   0   'False
      Top             =   6840
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   128
      Left            =   9135
      Locked          =   -1  'True
      TabIndex        =   138
      TabStop         =   0   'False
      Top             =   6840
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   127
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   137
      TabStop         =   0   'False
      Top             =   6840
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   126
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   136
      TabStop         =   0   'False
      Top             =   6600
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   125
      Left            =   9135
      Locked          =   -1  'True
      TabIndex        =   135
      TabStop         =   0   'False
      Top             =   6600
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   124
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   134
      TabStop         =   0   'False
      Top             =   6600
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   123
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   133
      TabStop         =   0   'False
      Top             =   6360
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   122
      Left            =   9135
      Locked          =   -1  'True
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   6360
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   121
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   131
      TabStop         =   0   'False
      Top             =   6360
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   120
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   130
      TabStop         =   0   'False
      Top             =   6120
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   119
      Left            =   9135
      Locked          =   -1  'True
      TabIndex        =   129
      TabStop         =   0   'False
      Top             =   6120
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   118
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   128
      TabStop         =   0   'False
      Top             =   6120
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   113
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   123
      TabStop         =   0   'False
      Top             =   8280
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   112
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   8040
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   111
      Left            =   6195
      Locked          =   -1  'True
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   8040
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   110
      Left            =   5460
      Locked          =   -1  'True
      TabIndex        =   120
      TabStop         =   0   'False
      Top             =   8040
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   109
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   119
      TabStop         =   0   'False
      Top             =   7800
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   108
      Left            =   6195
      Locked          =   -1  'True
      TabIndex        =   118
      TabStop         =   0   'False
      Top             =   7800
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   107
      Left            =   5460
      Locked          =   -1  'True
      TabIndex        =   117
      TabStop         =   0   'False
      Top             =   7800
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   106
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   116
      TabStop         =   0   'False
      Top             =   7560
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   105
      Left            =   6195
      Locked          =   -1  'True
      TabIndex        =   115
      TabStop         =   0   'False
      Top             =   7560
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   104
      Left            =   5460
      Locked          =   -1  'True
      TabIndex        =   114
      TabStop         =   0   'False
      Top             =   7560
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   103
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   113
      TabStop         =   0   'False
      Top             =   7320
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   101
      Left            =   5460
      Locked          =   -1  'True
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   7320
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   102
      Left            =   6195
      Locked          =   -1  'True
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   7320
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   100
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   7080
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   99
      Left            =   6195
      Locked          =   -1  'True
      TabIndex        =   109
      TabStop         =   0   'False
      Top             =   7080
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   98
      Left            =   5460
      Locked          =   -1  'True
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   7080
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   97
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   6840
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   96
      Left            =   6195
      Locked          =   -1  'True
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   6840
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   95
      Left            =   5460
      Locked          =   -1  'True
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   6840
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   94
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   6600
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   93
      Left            =   6195
      Locked          =   -1  'True
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   6600
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   92
      Left            =   5460
      Locked          =   -1  'True
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   6600
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   91
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   6360
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   90
      Left            =   6195
      Locked          =   -1  'True
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   6360
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   89
      Left            =   5460
      Locked          =   -1  'True
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   6360
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   86
      Left            =   5460
      Locked          =   -1  'True
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   6120
      Width           =   750
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1575
      Index           =   0
      Left            =   10500
      TabIndex        =   165
      Top             =   6120
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   2778
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"CNV_SEI00101.frx":0000
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   81
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   8280
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   80
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   8040
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   77
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   7800
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   74
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   7560
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   71
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   7320
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   68
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   7080
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   65
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   6840
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   62
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   6600
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   59
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   6360
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   56
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   6120
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   79
      Left            =   3045
      Locked          =   -1  'True
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   8040
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   76
      Left            =   3045
      Locked          =   -1  'True
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   7800
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   73
      Left            =   3045
      Locked          =   -1  'True
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   7560
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   70
      Left            =   3045
      Locked          =   -1  'True
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   7320
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   67
      Left            =   3045
      Locked          =   -1  'True
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   7080
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   64
      Left            =   3045
      Locked          =   -1  'True
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   6840
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   61
      Left            =   3045
      Locked          =   -1  'True
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   6600
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   58
      Left            =   3045
      Locked          =   -1  'True
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   6360
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   55
      Left            =   3045
      Locked          =   -1  'True
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   6120
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   78
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   8040
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   75
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   7800
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   72
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   7560
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   69
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   7320
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   66
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   7080
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   63
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   6840
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   60
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   6600
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   57
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   6360
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   54
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   6120
      Width           =   750
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1470
      MaxLength       =   20
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
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
      Index           =   3
      Left            =   4095
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   4320
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
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
      Index           =   4
      Left            =   8715
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
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
      Index           =   5
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
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
      Index           =   6
      Left            =   9660
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   720
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
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
      Index           =   7
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   720
      Width           =   330
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1470
      MaxLength       =   5
      TabIndex        =   0
      Top             =   480
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   53
      Left            =   13755
      Locked          =   -1  'True
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   52
      Left            =   12705
      Locked          =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   51
      Left            =   11655
      Locked          =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   50
      Left            =   10605
      Locked          =   -1  'True
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   49
      Left            =   9555
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   48
      Left            =   8505
      Locked          =   -1  'True
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   47
      Left            =   7455
      Locked          =   -1  'True
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   46
      Left            =   6405
      Locked          =   -1  'True
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   45
      Left            =   5355
      Locked          =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   44
      Left            =   4305
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   43
      Left            =   3255
      Locked          =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   42
      Left            =   2205
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   40
      Left            =   13755
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   39
      Left            =   12705
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   38
      Left            =   11655
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   37
      Left            =   10605
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   36
      Left            =   9555
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   35
      Left            =   8505
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   34
      Left            =   7455
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   33
      Left            =   6405
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   32
      Left            =   5355
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   31
      Left            =   4305
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   30
      Left            =   3255
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   29
      Left            =   2205
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   27
      Left            =   11655
      TabIndex        =   36
      Top             =   1920
      Width           =   3585
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   26
      Left            =   11025
      TabIndex        =   35
      Top             =   1920
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   25
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1920
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   24
      Left            =   6300
      TabIndex        =   30
      Top             =   1920
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   23
      Left            =   5355
      TabIndex        =   29
      Top             =   1920
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   22
      Left            =   4410
      TabIndex        =   28
      Top             =   1920
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   21
      Left            =   3465
      TabIndex        =   27
      Top             =   1920
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
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
      Index           =   20
      Left            =   2730
      MaxLength       =   5
      TabIndex        =   26
      Top             =   1920
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   19
      Left            =   1785
      MaxLength       =   8
      TabIndex        =   25
      Top             =   1920
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   18
      Left            =   840
      MaxLength       =   8
      TabIndex        =   24
      Top             =   1920
      Width           =   960
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   11655
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1680
      Width           =   3585
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   11025
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1680
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFC0&
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
      Index           =   15
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1680
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   6300
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1680
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   5355
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1680
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   4410
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1680
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   3465
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1680
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
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
      Index           =   10
      Left            =   2730
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1680
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   1785
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1680
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
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
      Index           =   8
      Left            =   840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1680
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "単価設定"
      Enabled         =   0   'False
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
      Index           =   5
      Left            =   8775
      TabIndex        =   175
      ToolTipText     =   "商品化単価を品目マスターに登録します"
      Top             =   0
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "見積書発行"
      Enabled         =   0   'False
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
      Index           =   4
      Left            =   7050
      TabIndex        =   174
      ToolTipText     =   "商品化単価見積書(EXCEL)を作成します"
      Top             =   0
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
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
      Index           =   1
      Left            =   2310
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   2325
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1470
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   10
      Top             =   960
      Width           =   2220
   End
   Begin VB.CommandButton Command1 
      Caption         =   "単価計算"
      Enabled         =   0   'False
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
      Index           =   3
      Left            =   5325
      TabIndex        =   173
      ToolTipText     =   "商品化単価を計算します(F9)"
      Top             =   0
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "構成−保存"
      Enabled         =   0   'False
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
      Left            =   3450
      TabIndex        =   172
      ToolTipText     =   "商品化構成を保存します"
      Top             =   0
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   11025
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   170
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton Command1 
      Caption         =   "構成−読込"
      Enabled         =   0   'False
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
      Left            =   1845
      TabIndex        =   171
      ToolTipText     =   "商品化構成を読み込みます（Ｆ5）"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "閉じる"
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
      Index           =   0
      Left            =   315
      TabIndex        =   169
      Top             =   0
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   41
      Left            =   1155
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   28
      Left            =   1155
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   87
      Left            =   6195
      TabIndex        =   97
      Top             =   6120
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   88
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   6120
      Width           =   750
   End
   Begin TrueDBGrid80.TDBDropDown TDBDropDown1 
      Height          =   2415
      Left            =   3885
      TabIndex        =   268
      Top             =   3480
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   4260
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).ValueItems(0)._DefaultItem=   0
      Columns(0).ValueItems(0).Value=   "aaaa"
      Columns(0).ValueItems(0).Value.vt=   8
      Columns(0).ValueItems(0).DisplayValue=   "aaaa"
      Columns(0).ValueItems(0).DisplayValue.vt=   8
      Columns(0).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(0).ValueItems.Count=   1
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   1
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=1"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4366"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   1
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowDividerStyle =   2
      LayoutName      =   ""
      LayoutFileName  =   ""
      LayoutURL       =   ""
      EmptyRows       =   0   'False
      ListField       =   ""
      DataField       =   "ub_grid2"
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   0   'False
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   14215660
      ValueTranslate  =   0   'False
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=ＭＳ ゴシック"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ＭＳ ゴシック"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Named:id=33:Normal"
      _StyleDefs(41)  =   ":id=33,.parent=0"
      _StyleDefs(42)  =   "Named:id=34:Heading"
      _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(44)  =   ":id=34,.wraptext=-1"
      _StyleDefs(45)  =   "Named:id=35:Footing"
      _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   "Named:id=36:Selected"
      _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(49)  =   "Named:id=37:Caption"
      _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(51)  =   "Named:id=38:HighlightRow"
      _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=39:EvenRow"
      _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(55)  =   "Named:id=40:OddRow"
      _StyleDefs(56)  =   ":id=40,.parent=33"
      _StyleDefs(57)  =   "Named:id=41:RecordSelector"
      _StyleDefs(58)  =   ":id=41,.parent=34"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=33"
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   2655
      Index           =   0
      Left            =   105
      TabIndex        =   63
      Top             =   3120
      Width           =   14715
      _ExtentX        =   25956
      _ExtentY        =   4683
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "事業部"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "国内外"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   1
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "種別"
      Columns(2).DataField=   ""
      Columns(2).DropDown=   "TDBDropDown1"
      Columns(2).DropDown.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "品番"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "品名"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "員数"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "仕入＠"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "販売＠"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "仕入金額計"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "販売金額計"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "作業時間（秒）"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "集合梱包（秒）"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "備考"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "販売金額　草津用"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   14
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=14"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1217"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1111"
      Splits(0)._ColumnProps(4)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1032"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=926"
      Splits(0)._ColumnProps(9)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2196"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2090"
      Splits(0)._ColumnProps(14)=   "Column(2).Button=1"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=1905"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1799"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=4710"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=4604"
      Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=8196"
      Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(25)=   "Column(5).Width=1164"
      Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=1058"
      Splits(0)._ColumnProps(28)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(30)=   "Column(6).Width=1879"
      Splits(0)._ColumnProps(31)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(6)._WidthInPix=1773"
      Splits(0)._ColumnProps(33)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(34)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(35)=   "Column(7).Width=2143"
      Splits(0)._ColumnProps(36)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(7)._WidthInPix=2037"
      Splits(0)._ColumnProps(38)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(39)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(40)=   "Column(8).Width=2117"
      Splits(0)._ColumnProps(41)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(8)._WidthInPix=2011"
      Splits(0)._ColumnProps(43)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(44)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(45)=   "Column(9).Width=2249"
      Splits(0)._ColumnProps(46)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(9)._WidthInPix=2143"
      Splits(0)._ColumnProps(48)=   "Column(9)._ColStyle=2"
      Splits(0)._ColumnProps(49)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(50)=   "Column(10).Width=3096"
      Splits(0)._ColumnProps(51)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(10)._WidthInPix=2990"
      Splits(0)._ColumnProps(53)=   "Column(10)._ColStyle=2"
      Splits(0)._ColumnProps(54)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(55)=   "Column(11).Width=3201"
      Splits(0)._ColumnProps(56)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(11)._WidthInPix=3096"
      Splits(0)._ColumnProps(58)=   "Column(11)._ColStyle=2"
      Splits(0)._ColumnProps(59)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(60)=   "Column(12).Width=3810"
      Splits(0)._ColumnProps(61)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(12)._WidthInPix=3704"
      Splits(0)._ColumnProps(63)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(64)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(65)=   "Column(13).Width=3810"
      Splits(0)._ColumnProps(66)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(13)._WidthInPix=3704"
      Splits(0)._ColumnProps(68)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(69)=   "Column(13).Order=14"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=ＭＳ ゴシック"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ＭＳ ゴシック"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=82,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=78,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=75,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=76,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=77,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.locked=-1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=1"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=58,.parent=13,.alignment=1"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=62,.parent=13,.alignment=1"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=66,.parent=13,.alignment=1"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=70,.parent=13,.alignment=1"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=67,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=68,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=69,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=74,.parent=13,.alignment=1"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=71,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=72,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=73,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=90,.parent=13"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
      _StyleDefs(92)  =   "Named:id=33:Normal"
      _StyleDefs(93)  =   ":id=33,.parent=0"
      _StyleDefs(94)  =   "Named:id=34:Heading"
      _StyleDefs(95)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(96)  =   ":id=34,.wraptext=-1"
      _StyleDefs(97)  =   "Named:id=35:Footing"
      _StyleDefs(98)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(99)  =   "Named:id=36:Selected"
      _StyleDefs(100) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(101) =   "Named:id=37:Caption"
      _StyleDefs(102) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(103) =   "Named:id=38:HighlightRow"
      _StyleDefs(104) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(105) =   "Named:id=39:EvenRow"
      _StyleDefs(106) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(107) =   "Named:id=40:OddRow"
      _StyleDefs(108) =   ":id=40,.parent=33"
      _StyleDefs(109) =   "Named:id=41:RecordSelector"
      _StyleDefs(110) =   ":id=41,.parent=34"
      _StyleDefs(111) =   "Named:id=42:FilterBar"
      _StyleDefs(112) =   ":id=42,.parent=33"
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   150
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   160
      TabStop         =   0   'False
      Top             =   8280
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   161
      Left            =   8190
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1680
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "外装"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   99
      Left            =   7350
      TabIndex        =   280
      Top             =   1500
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "BU加工"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   98
      Left            =   9135
      TabIndex        =   279
      Top             =   1500
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "PP加工"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   97
      Left            =   8190
      TabIndex        =   278
      Top             =   1500
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "(円／個)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   79
      Left            =   2310
      TabIndex        =   256
      Top             =   9360
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "(円／個)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   91
      Left            =   -105
      TabIndex        =   267
      Top             =   9360
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "(円／個)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   87
      Left            =   8400
      TabIndex        =   262
      Top             =   9360
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "(円／個)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   84
      Left            =   5145
      TabIndex        =   259
      Top             =   9360
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "(分／個)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   90
      Left            =   -105
      TabIndex        =   266
      Top             =   9120
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "(分／個)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   86
      Left            =   8400
      TabIndex        =   261
      Top             =   9120
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "(分／個)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   83
      Left            =   5145
      TabIndex        =   258
      Top             =   9120
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "(分／個)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   78
      Left            =   2310
      TabIndex        =   255
      Top             =   9120
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "(秒／個)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   85
      Left            =   8400
      TabIndex        =   260
      Top             =   8880
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "(秒／個)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   80
      Left            =   5145
      TabIndex        =   257
      Top             =   8880
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "(秒／個)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   77
      Left            =   2310
      TabIndex        =   254
      Top             =   8880
      Width           =   1185
   End
   Begin VB.Label YOYU_RITU 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   9135
      TabIndex        =   276
      Top             =   8640
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "余裕率(　　)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   95
      Left            =   8400
      TabIndex        =   275
      Top             =   8640
      Width           =   1185
   End
   Begin VB.Label YOYU_RITU 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6195
      TabIndex        =   274
      Top             =   8640
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "余裕率(　　)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   94
      Left            =   5460
      TabIndex        =   273
      Top             =   8640
      Width           =   1185
   End
   Begin VB.Label YOYU_RITU 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3045
      TabIndex        =   272
      Top             =   8640
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "余裕率(　　)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   93
      Left            =   2310
      TabIndex        =   271
      Top             =   8640
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "(秒／個)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   89
      Left            =   -105
      TabIndex        =   265
      Top             =   8880
      Width           =   1185
   End
   Begin VB.Label YOYU_RITU 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   735
      TabIndex        =   270
      Top             =   8640
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "余裕率(　　)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   92
      Left            =   0
      TabIndex        =   269
      Top             =   8640
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "工程計(秒)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   88
      Left            =   0
      TabIndex        =   264
      Top             =   8400
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "後工程計(秒)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   61
      Left            =   8400
      TabIndex        =   263
      Top             =   8400
      Width           =   1185
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   58
      Left            =   7455
      TabIndex        =   253
      Top             =   8040
      Width           =   960
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   57
      Left            =   7455
      TabIndex        =   252
      Top             =   7800
      Width           =   960
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   56
      Left            =   7455
      TabIndex        =   251
      Top             =   7560
      Width           =   960
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   55
      Left            =   7455
      TabIndex        =   250
      Top             =   7320
      Width           =   960
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   54
      Left            =   7455
      TabIndex        =   249
      Top             =   7080
      Width           =   960
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   45
      Left            =   7455
      TabIndex        =   248
      Top             =   6840
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "後工程"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   52
      Left            =   7455
      TabIndex        =   247
      Top             =   5880
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "単位（秒"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   51
      Left            =   8400
      TabIndex        =   246
      Top             =   5880
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "数量"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   50
      Left            =   9135
      TabIndex        =   245
      Top             =   5880
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "工数（秒"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   48
      Left            =   9660
      TabIndex        =   244
      Top             =   5880
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "作業工程計(秒)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   76
      Left            =   5145
      TabIndex        =   243
      Top             =   8400
      Width           =   1500
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   75
      Left            =   4305
      TabIndex        =   242
      Top             =   8040
      Width           =   1170
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   74
      Left            =   4305
      TabIndex        =   241
      Top             =   7800
      Width           =   1170
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   73
      Left            =   4305
      TabIndex        =   240
      Top             =   7560
      Width           =   1170
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   53
      Left            =   4305
      TabIndex        =   239
      Top             =   7320
      Width           =   1170
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "検査表記入"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   49
      Left            =   7455
      TabIndex        =   238
      Top             =   6120
      Width           =   960
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "部品搬入"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   47
      Left            =   7455
      TabIndex        =   237
      Top             =   6360
      Width           =   960
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "後片付け"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   46
      Left            =   7455
      TabIndex        =   236
      Top             =   6600
      Width           =   960
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "ラベル貼り"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   44
      Left            =   4305
      TabIndex        =   235
      Top             =   6120
      Width           =   1170
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "個装作業"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   43
      Left            =   4305
      TabIndex        =   234
      Top             =   6360
      Width           =   1170
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "同梱作業"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   42
      Left            =   4305
      TabIndex        =   233
      Top             =   6600
      Width           =   1170
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "加工作業"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   41
      Left            =   4305
      TabIndex        =   232
      Top             =   6840
      Width           =   1170
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "集合梱包作業"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   37
      Left            =   4305
      TabIndex        =   231
      Top             =   7080
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "作業工程"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   82
      Left            =   4305
      TabIndex        =   230
      Top             =   5880
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "単位（秒"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   40
      Left            =   5460
      TabIndex        =   229
      Top             =   5880
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "数量"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   39
      Left            =   6195
      TabIndex        =   228
      Top             =   5880
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "工数（秒"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   38
      Left            =   6720
      TabIndex        =   227
      Top             =   5880
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "メモ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   16
      Left            =   11655
      TabIndex        =   192
      Top             =   1500
      Width           =   3585
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "工料"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   9
      Left            =   3465
      TabIndex        =   186
      Top             =   1500
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "担当者"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   15
      Left            =   11025
      TabIndex        =   191
      Top             =   1500
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "設定日"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   14
      Left            =   9975
      TabIndex        =   190
      Top             =   1500
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "工料"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   13
      Left            =   4410
      TabIndex        =   189
      Top             =   1500
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "箱代"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   11
      Left            =   6300
      TabIndex        =   188
      Top             =   1500
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "箱代"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   10
      Left            =   5355
      TabIndex        =   187
      Top             =   1500
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "工数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   8
      Left            =   2835
      TabIndex        =   185
      Top             =   1500
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "分レート"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   4
      Left            =   1995
      TabIndex        =   184
      Top             =   1500
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "ロット数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   3
      Left            =   945
      TabIndex        =   183
      Top             =   1500
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "指図票備考"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   81
      Left            =   10500
      TabIndex        =   226
      Top             =   5880
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "前工程計(秒)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   60
      Left            =   2310
      TabIndex        =   225
      Top             =   8400
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "工数（秒"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   72
      Left            =   3570
      TabIndex        =   224
      Top             =   5880
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "数量"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   71
      Left            =   3045
      TabIndex        =   223
      Top             =   5880
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "単位(秒"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   59
      Left            =   2310
      TabIndex        =   222
      Top             =   5880
      Width           =   750
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "見本確認"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   70
      Left            =   105
      TabIndex        =   221
      Top             =   8040
      Width           =   2220
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "指図票／ラベル発行"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   69
      Left            =   105
      TabIndex        =   220
      Top             =   7800
      Width           =   2220
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "ラベル発行"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   68
      Left            =   105
      TabIndex        =   219
      Top             =   7560
      Width           =   2220
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "同梱部品準備"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   67
      Left            =   105
      TabIndex        =   218
      Top             =   7320
      Width           =   2220
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "副資材準備"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   66
      Left            =   105
      TabIndex        =   217
      Top             =   7080
      Width           =   2220
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "部品準備"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   65
      Left            =   105
      TabIndex        =   216
      Top             =   6840
      Width           =   2220
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "在庫有無確認（同梱部品）"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   64
      Left            =   105
      TabIndex        =   215
      Top             =   6600
      Width           =   2220
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "在庫有無確認（副資材）"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   63
      Left            =   105
      TabIndex        =   214
      Top             =   6360
      Width           =   2220
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "事前商品化部品　数量選定"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   62
      Left            =   105
      TabIndex        =   213
      Top             =   6120
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "出荷数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   20
      Left            =   105
      TabIndex        =   212
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "前工程"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   36
      Left            =   105
      TabIndex        =   211
      Top             =   5880
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "３"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   35
      Left            =   13755
      TabIndex        =   210
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "２"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   34
      Left            =   12705
      TabIndex        =   209
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "１"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   33
      Left            =   11655
      TabIndex        =   208
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "１２"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   32
      Left            =   10605
      TabIndex        =   207
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "１１"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   31
      Left            =   9555
      TabIndex        =   206
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "１０"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   30
      Left            =   8505
      TabIndex        =   205
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "９"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   29
      Left            =   7455
      TabIndex        =   204
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "８"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   28
      Left            =   6405
      TabIndex        =   203
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "７"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   27
      Left            =   5355
      TabIndex        =   202
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "６"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   26
      Left            =   4305
      TabIndex        =   201
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "５"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   25
      Left            =   3255
      TabIndex        =   200
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "４"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   24
      Left            =   2205
      TabIndex        =   199
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "平均"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   23
      Left            =   1155
      TabIndex        =   198
      Top             =   2280
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "今年度"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   22
      Left            =   105
      TabIndex        =   197
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "前年度"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   21
      Left            =   105
      TabIndex        =   196
      Top             =   2520
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "変更後"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   17
      Left            =   105
      TabIndex        =   193
      Top             =   1920
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "変更前"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   105
      TabIndex        =   182
      Top             =   1680
      Width           =   765
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   14385
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "担当者"
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
      Index           =   12
      Left            =   750
      TabIndex        =   181
      Top             =   480
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "-"
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
      Index           =   7
      Left            =   9975
      TabIndex        =   180
      Top             =   720
      Width           =   225
   End
   Begin VB.Label Label1 
      Caption         =   "-"
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
      Index           =   6
      Left            =   9555
      TabIndex        =   179
      Top             =   720
      Width           =   225
   End
   Begin VB.Label Label1 
      Caption         =   "-"
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
      Index           =   5
      Left            =   9030
      TabIndex        =   178
      Top             =   720
      Width           =   225
   End
   Begin VB.Label Label1 
      Caption         =   "品目コード"
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
      Index           =   1
      Left            =   315
      TabIndex        =   177
      Top             =   720
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "仕向先"
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
      Left            =   525
      TabIndex        =   176
      Top             =   1080
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "(原価)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   19
      Left            =   5355
      TabIndex        =   195
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "(原価)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   18
      Left            =   3540
      TabIndex        =   194
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "工程計"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   96
      Left            =   1050
      TabIndex        =   277
      Top             =   8520
      Width           =   1185
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "閉じる"
         Index           =   0
         Shortcut        =   {F12}
      End
      Begin VB.Menu SHORI 
         Caption         =   "検索"
         Index           =   1
         Shortcut        =   {F5}
      End
      Begin VB.Menu SHORI 
         Caption         =   "保存"
         Index           =   2
      End
      Begin VB.Menu SHORI 
         Caption         =   "単価計算"
         Index           =   3
         Shortcut        =   {F9}
      End
      Begin VB.Menu SHORI 
         Caption         =   "見積書発行"
         Index           =   4
      End
      Begin VB.Menu SHORI 
         Caption         =   "単価登録"
         Index           =   5
      End
      Begin VB.Menu SHORI 
         Caption         =   "画面印刷"
         Index           =   6
      End
   End
End
Attribute VB_Name = "CNV_SEI00101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'------------------------------------   'テキスト定義

Private Const ptxTanto_Code% = 0            '担当者コード
Private Const ptxTanto_Name% = 1            '担当者名称
Private Const ptxHin_Gai% = 2               '品番
Private Const ptxHin_Name% = 3              '品名

Private Const ptxST_SOKO% = 4               '標準棚番　 倉庫
Private Const ptxST_RETU% = 5               '標準棚番   列
Private Const ptxST_REN% = 6                '標準棚番　 連
Private Const ptxST_DAN% = 7                '標準棚番　 段

Private Const ptxBEF_SEI_LOT% = 8           '変更前　   ロット数
Private Const ptxBEF_SEI_RATE% = 9          '           分レート
Private Const ptxBEF_S_KOUSU% = 10          '           分レート
Private Const ptxBEF_S_KOUSU_GENKA% = 11    '           (原価)商品化工料
Private Const ptxBEF_S_KOUSU_BAIKA% = 12    '           (売価)商品化工料
Private Const ptxBEF_S_SHIZAI_GENKA% = 13   '           (原価)箱代
Private Const ptxBEF_S_SHIZAI_BAIKA% = 14   '           (売価)箱代

Private Const ptxBEF_S_GAISO_TANKA% = 165   '           外装単価
Private Const ptxBEF_S_PPSC_KAKO_KOSU% = 161 '          PPSC加工単価
Private Const ptxBEF_S_BU_KAKO_KOSU% = 162  '           BU加工単価




Private Const ptxBEF_S_KOUSU_SET_DATE% = 15 '          設定日
Private Const ptxBEF_SEI_TANKA_TANTO% = 16  '          担当者
Private Const ptxBEF_SE_TANKA_MEMO% = 17    '          メモ

Private Const ptxAFT_SEI_LOT% = 18          '変更後　   ロット数
Private Const ptxAFT_SEI_RATE% = 19         '           分レート
Private Const ptxAFT_S_KOUSU% = 20          '           工数
Private Const ptxAFT_S_KOUSU_GENKA% = 21    '           (原価)商品化工料
Private Const ptxAFT_S_KOUSU_BAIKA% = 22    '           (売価)商品化工料
Private Const ptxAFT_S_SHIZAI_GENKA% = 23   '           (原価)箱代
Private Const ptxAFT_S_SHIZAI_BAIKA% = 24   '           (売価)箱代




Private Const ptxAFT_S_GAISO_TANKA% = 166   '           外装単価
Private Const ptxAFT_S_PPSC_KAKO_KOSU% = 163 '          PPSC加工単価
Private Const ptxAFT_S_BU_KAKO_KOSU% = 164  '           BU加工単価


Private Const ptxAFT_S_KOUSU_SET_DATE% = 25 '          設定日
Private Const ptxAFT_SEI_TANKA_TANTO% = 26  '          担当者
Private Const ptxAFT_SE_TANKA_MEMO% = 27    '          メモ


Private Const ptxZEN_AVE% = 28              '月平均出荷数   前年度　平均
Private Const ptxZEN_SYUKAQTY04% = 29       '月平均出荷数   前年度　4月
Private Const ptxZEN_SYUKAQTY05% = 30       '　                     5月
Private Const ptxZEN_SYUKAQTY06% = 31       '　                     6月
Private Const ptxZEN_SYUKAQTY07% = 32       '　                     7月
Private Const ptxZEN_SYUKAQTY08% = 33       '　                     8月
Private Const ptxZEN_SYUKAQTY09% = 34       '　                     9月
Private Const ptxZEN_SYUKAQTY10% = 35       '　                     10月
Private Const ptxZEN_SYUKAQTY11% = 36       '　                     11月
Private Const ptxZEN_SYUKAQTY12% = 37       '　                     12月
Private Const ptxZEN_SYUKAQTY01% = 38       '　                     1月
Private Const ptxZEN_SYUKAQTY02% = 39       '　                     2月
Private Const ptxZEN_SYUKAQTY03% = 40       '　                     3月

Private Const ptxTOU_AVE% = 41              '月平均出荷数   今年度　平均
Private Const ptxTOU_SYUKAQTY04% = 42       '月平均出荷数   今年度　4月
Private Const ptxTOU_SYUKAQTY05% = 43       '　                     5月
Private Const ptxTOU_SYUKAQTY06% = 44       '　                     6月
Private Const ptxTOU_SYUKAQTY07% = 45       '　                     7月
Private Const ptxTOU_SYUKAQTY08% = 46       '　                     8月
Private Const ptxTOU_SYUKAQTY09% = 47       '　                     9月
Private Const ptxTOU_SYUKAQTY10% = 48       '　                     10月
Private Const ptxTOU_SYUKAQTY11% = 49       '　                     11月
Private Const ptxTOU_SYUKAQTY12% = 50       '　                     12月
Private Const ptxTOU_SYUKAQTY01% = 51       '　                     1月
Private Const ptxTOU_SYUKAQTY02% = 52       '　                     2月
Private Const ptxTOU_SYUKAQTY03% = 53       '　                     3月





Private Const ptxBEF_KOUTEI_TANI01% = 54    '前工程01　 単位
Private Const ptxBEF_KOUTEI_QTY01% = 55     '           数量
Private Const ptxBEF_KOUTEI_KOUSU01% = 56   '           工数
Private Const ptxBEF_KOUTEI_TANI02% = 57    '前工程02　 単位
Private Const ptxBEF_KOUTEI_QTY02% = 58     '           数量
Private Const ptxBEF_KOUTEI_KOUSU02% = 59   '           工数
Private Const ptxBEF_KOUTEI_TANI03% = 60    '前工程03　 単位
Private Const ptxBEF_KOUTEI_QTY03% = 61     '           数量
Private Const ptxBEF_KOUTEI_KOUSU03% = 62   '           工数
Private Const ptxBEF_KOUTEI_TANI04% = 63    '前工程04　 単位
Private Const ptxBEF_KOUTEI_QTY04% = 64     '           数量
Private Const ptxBEF_KOUTEI_KOUSU04% = 65   '           工数
Private Const ptxBEF_KOUTEI_TANI05% = 66    '前工程05　 単位
Private Const ptxBEF_KOUTEI_QTY05% = 67     '           数量
Private Const ptxBEF_KOUTEI_KOUSU05% = 68   '           工数
Private Const ptxBEF_KOUTEI_TANI06% = 69    '前工程06　 単位
Private Const ptxBEF_KOUTEI_QTY06% = 70     '           数量
Private Const ptxBEF_KOUTEI_KOUSU06% = 71   '           工数
Private Const ptxBEF_KOUTEI_TANI07% = 72    '前工程07　 単位
Private Const ptxBEF_KOUTEI_QTY07% = 73     '           数量
Private Const ptxBEF_KOUTEI_KOUSU07% = 74   '           工数
Private Const ptxBEF_KOUTEI_TANI08% = 75    '前工程08　 単位
Private Const ptxBEF_KOUTEI_QTY08% = 76     '           数量
Private Const ptxBEF_KOUTEI_KOUSU08% = 77   '           工数
Private Const ptxBEF_KOUTEI_TANI09% = 78    '前工程09　 単位
Private Const ptxBEF_KOUTEI_QTY09% = 79     '           数量
Private Const ptxBEF_KOUTEI_KOUSU09% = 80   '           工数

Private Const ptxBEF_KOUTEI_KEI1% = 81      '前工程計   計

Private Const ptxBEF_KOUTEI_R_RATE% = 82    '前工程計   余裕率

Private Const ptxBEF_KOUTEI_KEI2% = 83      '前工程計   (秒／個)
Private Const ptxBEF_KOUTEI_KEI3% = 84      '前工程計   (分／個)
Private Const ptxBEF_KOUTEI_KEI4% = 85      '前工程計   (円／個)

Private Const ptxMAIN_KOUTEI_TANI01% = 86   '作業工程01 単位
Private Const ptxMAIN_KOUTEI_QTY01% = 87    '           数量
Private Const ptxMAIN_KOUTEI_KOUSU01% = 88  '           工数
Private Const ptxMAIN_KOUTEI_TANI02% = 89   '作業工程02 単位
Private Const ptxMAIN_KOUTEI_QTY02% = 90    '           数量
Private Const ptxMAIN_KOUTEI_KOUSU02% = 91  '           工数
Private Const ptxMAIN_KOUTEI_TANI03% = 92   '作業工程03 単位
Private Const ptxMAIN_KOUTEI_QTY03% = 93    '           数量
Private Const ptxMAIN_KOUTEI_KOUSU03% = 94  '           工数
Private Const ptxMAIN_KOUTEI_TANI04% = 95   '作業工程04 単位
Private Const ptxMAIN_KOUTEI_QTY04% = 96    '           数量
Private Const ptxMAIN_KOUTEI_KOUSU04% = 97  '           工数
Private Const ptxMAIN_KOUTEI_TANI05% = 98   '作業工程05 単位
Private Const ptxMAIN_KOUTEI_QTY05% = 99    '           数量
Private Const ptxMAIN_KOUTEI_KOUSU05% = 100 '           工数
Private Const ptxMAIN_KOUTEI_TANI06% = 101  '作業工程06 単位
Private Const ptxMAIN_KOUTEI_QTY06% = 102   '           数量
Private Const ptxMAIN_KOUTEI_KOUSU06% = 103 '           工数
Private Const ptxMAIN_KOUTEI_TANI07% = 104  '作業工程07 単位
Private Const ptxMAIN_KOUTEI_QTY07% = 105   '           数量
Private Const ptxMAIN_KOUTEI_KOUSU07% = 106 '           工数
Private Const ptxMAIN_KOUTEI_TANI08% = 107  '作業工程08 単位
Private Const ptxMAIN_KOUTEI_QTY08% = 108   '           数量
Private Const ptxMAIN_KOUTEI_KOUSU08% = 109 '           工数
Private Const ptxMAIN_KOUTEI_TANI09% = 110  '作業工程09 単位
Private Const ptxMAIN_KOUTEI_QTY09% = 111   '           数量
Private Const ptxMAIN_KOUTEI_KOUSU09% = 112 '           工数

Private Const ptxMAIN_KOUTEI_KEI1% = 113    '作業工程計 計

Private Const ptxMAIN_KOUTEI_R_RATE% = 114  '作業工程計   余裕率


Private Const ptxMAIN_KOUTEI_KEI2% = 115    '作業工程計  (秒／個)
Private Const ptxMAIN_KOUTEI_KEI3% = 116    '作業工程計  (分／個)
Private Const ptxMAIN_KOUTEI_KEI4% = 117    '作業工程計  (円／個)

Private Const ptxAFT_KOUTEI_TANI01% = 118   '後工程01   単位
Private Const ptxAFT_KOUTEI_QTY01% = 119    '           数量
Private Const ptxAFT_KOUTEI_KOUSU01% = 120  '           工数
Private Const ptxAFT_KOUTEI_TANI02% = 121   '後工程02   単位
Private Const ptxAFT_KOUTEI_QTY02% = 122    '           数量
Private Const ptxAFT_KOUTEI_KOUSU02% = 123  '           工数
Private Const ptxAFT_KOUTEI_TANI03% = 124   '後工程03   単位
Private Const ptxAFT_KOUTEI_QTY03% = 125    '           数量
Private Const ptxAFT_KOUTEI_KOUSU03% = 126  '           工数
Private Const ptxAFT_KOUTEI_TANI04% = 127   '後工程04   単位
Private Const ptxAFT_KOUTEI_QTY04% = 128    '           数量
Private Const ptxAFT_KOUTEI_KOUSU04% = 129  '           工数
Private Const ptxAFT_KOUTEI_TANI05% = 130   '後工程05   単位
Private Const ptxAFT_KOUTEI_QTY05% = 131    '           数量
Private Const ptxAFT_KOUTEI_KOUSU05% = 132  '           工数
Private Const ptxAFT_KOUTEI_TANI06% = 133   '後工程06   単位
Private Const ptxAFT_KOUTEI_QTY06% = 134    '           数量
Private Const ptxAFT_KOUTEI_KOUSU06% = 135  '           工数
Private Const ptxAFT_KOUTEI_TANI07% = 136   '後工程07   単位
Private Const ptxAFT_KOUTEI_QTY07% = 137    '           数量
Private Const ptxAFT_KOUTEI_KOUSU07% = 138  '           工数
Private Const ptxAFT_KOUTEI_TANI08% = 139   '後工程08   単位
Private Const ptxAFT_KOUTEI_QTY08% = 140    '           数量
Private Const ptxAFT_KOUTEI_KOUSU08% = 141  '           工数
Private Const ptxAFT_KOUTEI_TANI09% = 142   '後工程09   単位
Private Const ptxAFT_KOUTEI_QTY09% = 143    '           数量
Private Const ptxAFT_KOUTEI_KOUSU09% = 144  '           工数

Private Const ptxAFT_KOUTEI_KEI1% = 145     '後工程計   計

Private Const ptxAFT_KOUTEI_R_RATE% = 146   '後工程計   余裕率



Private Const ptxAFT_KOUTEI_KEI2% = 147     '後工程計   (秒／個)
Private Const ptxAFT_KOUTEI_KEI3% = 148     '後工程計   (分／個)
Private Const ptxAFT_KOUTEI_KEI4% = 149     '後工程計   (円／個)


Private Const ptxKOUTEI_KEI1% = 150         '工程計   計

Private Const ptxKOUTEI_R_RATE% = 151       '工程計   余裕率


Private Const ptxKOUTEI_KEI2% = 152         '工程計   (秒／個)
Private Const ptxKOUTEI_KEI3% = 153         '工程計   (分／個)
Private Const ptxKOUTEI_KEI4% = 154         '工程計   (円／個)


Private Const ptxS_CLASS_CODE% = 155        '商品化ｸﾗｽ
Private Const ptxF_CLASS_CODE% = 156        '付加ｸﾗｽ
Private Const ptxN_CLASS_CODE% = 157        '内職ｸﾗｽ

Private Const ptxIO_TANKA_No% = 158         '棚区分
Private Const ptxSE_Name% = 159             '棚区分名称



'------------------------------------   'コンボ定義
Private Const pcmbSHIMUKE% = 0          '仕向け先


'------------------------------------   'リッチテキストボックス定義
Private Const prchBIKOU% = 0            '備考



'------------------------------------   '構成品
Private Const pGrdKOUSEI% = 0

Dim KOUSEI      As New XArrayDB

Private Const Min_Row% = 1              '最小行数

Dim Max_Row   As Integer                'グリッド最大表示件数

Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 13             '最大列数

Private Const ColKO_JGYOBU% = 0         '事業部
Private Const ColKO_NAIGAI% = 1         '国内外


Private Const ColKO_SYUBETSU% = 2       '種別
Private Const ColKO_HIN_GAI% = 3        '品番
Private Const ColKO_HIN_NAME% = 4       '品名
Private Const ColKO_QTY% = 5            '員数
Private Const ColG_ST_SHITAN% = 6       '仕入＠
Private Const ColG_ST_URITAN% = 7       '売上＠
Private Const ColG_ST_SHIKIN% = 8       '仕入金額
Private Const ColG_ST_URIKIN% = 9       '売上金額
Private Const ColS_KOUSU% = 10          '作業時間
Private Const ColSEI_SYU_KON% = 11      '集合梱包
Private Const ColKO_BIKOU% = 12         '備考


                                        '草津 金額出力用
Private Const ColG_ST_URIKIN_KUSATU% = 13



'-----------------------------------    ドロップダウン
Dim SYUBETSU        As New XArrayDB


'-----------------------------------

Dim KOSOU_KBN       As String * 2       '個装区分
Dim GAISO_KBN       As String * 2       '外装区分


Dim INV_IO_TANKA_No As String * 2       '標準棚未登録時の出庫区分
Dim HIN_INV         As Boolean          '未登録品番の登録可否


Dim KUSATU_F        As Boolean          '対象センター　草津 OR 草津以外


Dim SHIZAI_T        As Variant          '資材対象
Dim DOUKON_T        As Variant          '同梱対象
Dim KAKOU_T         As Variant          '加工対象

Dim BU_T            As Variant          'BU付加対象
Dim PPSC_T          As Variant          'PPSC付加対象

Private Const KUSATU_ETC$ = "その他"


Dim svHin_Gai       As String           '品番
Dim svSHIMUKE_CODE  As String           '仕向け先

'-----------------------------------    ＥＸＣＥＬ 宛名＆住所

Dim EX_NAME1        As String           '宛名１
Dim EX_NAME2        As String           '宛名２

Dim EX_SYAMEI       As String           '自社　名称
Dim EX_ADDR1        As String           '自社　住所１
Dim EX_ADDR2        As String           '自社　住所２
Dim EX_BIKOU1       As String           '自社　備考


Dim EX_CENTER_NAME  As String           'センター   名称
Dim EX_CENTER_ADDR1 As String           'センター   名称１
Dim EX_CENTER_BIKOU1    As String       'センター   備考１


Private Const LAST_UPDATE_DAY$ = "2008.09.03"






Private Sub Combo1_Change(Index As Integer)
Dim i   As Integer
    
    
    If ptxHin_Gai = Index Then
        If Trim(svHin_Gai) = (Text1(ptxHin_Gai).Text) Then
        Else
            For i = 2 To 5
                Command1(i).Enabled = False
            Next i
        End If
    End If

End Sub

Private Sub Combo1_GotFocus(Index As Integer)

    If Index = pcmbSHIMUKE Then
        svSHIMUKE_CODE = Right(Combo1(pcmbSHIMUKE).Text, 2)
    End If

End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim i   As Integer
    
    If Index = pcmbSHIMUKE Then
        
        If Trim(svSHIMUKE_CODE) = Right(Combo1(pcmbSHIMUKE).Text, 2) Then
        Else
            For i = 2 To 5
                Command1(i).Enabled = False
            Next i
        End If

    End If
End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans     As Integer
Dim i       As Integer

Dim MESG    As String

Dim com     As Integer
Dim sts     As Integer



    Select Case Index
    
        Case 0      '終了
            Unload Me
    
        Case 5      '単価登録
            
            
            
            
            
            
            
            
            
            
            MESG = "単価を登録します。よろしいですか？" & vbCrLf

            
            
            
            ans = MsgBox(MESG, vbYesNo + vbDefaultButton1 + vbExclamation, "確認入力")
            If ans = vbYes Then
                
                Call Input_Lock
                
                com = BtOpGetFirst
                
                
                Do
                    
                    DoEvents
                    
                    sts = Detail_Disp_Proc(com)
                        
                    Select Case sts
                        
                        Case False, SYS_CANCEL
                        
                        Case BtErrEOF
                            Exit Do
                        Case True
                                
                            Unload Me
                    End Select
                
                                    
                    If sts = False Then
                
                        If Tanka_Update_Proc() Then
                            Unload Me
                        End If
                    End If
                
                    com = BtOpGetNext
                
                
                Loop
                
                
                
                
                
                
                
                
                
                
            
            
                Call Input_UnLock
            
            
            
            
                MsgBox "終了しました！！"
            
            
            
            End If
        
                    
    
    End Select






End Sub

Private Sub Form_Activate()

Dim com As Integer
Dim sts As Integer


        Call Log_Out(LOG_F, "見積書情報コンバート処理　開始")


        Call Input_Lock
        
        
'        Call UniCode_Conv(K0_ITEM.JGYOBU, "D")
'        Call UniCode_Conv(K0_ITEM.NAIGAI, "")
'        Call UniCode_Conv(K0_ITEM.HIN_GAI, "")

        
        com = BtOpGetFirst
'        com = BtOpGetGreaterEqual
        
        
        Do
            
            DoEvents
            
            sts = Detail_Disp_Proc(com)
                
            Select Case sts
                
                Case False, SYS_CANCEL
                
                Case BtErrEOF
                    Exit Do
                Case True
                        
                    Unload Me
            End Select
        
                            
            If sts = False Then
        
                If Tanka_Update_Proc() Then
                    Unload Me
                End If
            End If
        
            com = BtOpGetNext
        
        
        Loop
                
                
                        
                
                
                
                
                
                
                
            
            
        Call Input_UnLock




        Call Log_Out(LOG_F, "見積書情報コンバート処理　終了")

        Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim c       As String * 128
Dim sts     As Integer



'    If App.PrevInstance Then
'        Beep
'        MsgBox "同一プログラム実行中です。"
'        End
'    End If


    
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]商品化単価見積作成処理", Me.hwnd, 0)
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
                                'センターの識別
    If GetIni(App.EXEName, "KUSATU", App.EXEName, c) Then
        KUSATU_F = False
    Else
        If Trim(c) = "1" Then
            KUSATU_F = True
        Else
            KUSATU_F = False
        End If
    End If
                                
                                
                                
                                
                                
                                
                                
                                
                                '個装資材区分の獲得
    If GetIni(App.EXEName, "KOSOU", App.EXEName, c) Then
        Beep
        MsgBox "個装資材区分の獲得に失敗しました。処理を中止して下さい。"
        End
    Else
        KOSOU_KBN = Trim(c)
    End If
                                '外装資材区分の獲得
    If GetIni(App.EXEName, "GAISO", App.EXEName, c) Then
        Beep
        MsgBox "外装資材区分の獲得に失敗しました。処理を中止して下さい。"
        End
    Else
        GAISO_KBN = Trim(c)
    End If
                                '未登録時の出庫区分の獲得
    If GetIni(App.EXEName, "INV_IO_TANKA_No", App.EXEName, c) Then
        INV_IO_TANKA_No = ""
    Else
        INV_IO_TANKA_No = Trim(c)
    End If
                                '未登録品番の登録可否の獲得
    If GetIni(App.EXEName, "HIN_INV", App.EXEName, c) Then
        HIN_INV = False
    Else
        If Trim(c) = "0" Then
            HIN_INV = False
        Else
            HIN_INV = True
        End If
    End If
                                
                                '資材対象種別
    If GetIni(App.EXEName, "SHIZAI", App.EXEName, c) Then
        Beep
        MsgBox "資材対象の獲得に失敗しました。処理を中止して下さい。"
        End
    Else
        SHIZAI_T = Split(Trim(c), ",", -1)
    End If
                                '同梱対象種別
    If GetIni(App.EXEName, "DOUKON", App.EXEName, c) Then
        Beep
        MsgBox "同梱対象の獲得に失敗しました。処理を中止して下さい。"
        End
    Else
        DOUKON_T = Split(Trim(c), ",", -1)
    End If
                                '加工対象種別
   If GetIni(App.EXEName, "KAKOU", App.EXEName, c) Then
        Beep
        MsgBox "加工対象の獲得に失敗しました。処理を中止して下さい。"
        End
    Else
        KAKOU_T = Split(Trim(c), ",", -1)
    End If
                                
                                
                                'PPSC対象種別
    If GetIni(App.EXEName, "PPSC", App.EXEName, c) Then
        Beep
        MsgBox "PPSC対象の獲得に失敗しました。処理を中止して下さい。"
        End
    Else
        PPSC_T = Split(Trim(c), ",", -1)
    End If
                                'BU対象種別
    If GetIni(App.EXEName, "BU", App.EXEName, c) Then
        Beep
        MsgBox "BU対象の獲得に失敗しました。処理を中止して下さい。"
        End
    Else
        BU_T = Split(Trim(c), ",", -1)
    End If
                                
                                
                                
                                
                                '見積書 宛名１
    If GetIni(App.EXEName, "NAME1", App.EXEName, c) Then
        EX_NAME1 = ""
    Else
        EX_NAME1 = Trim(c)
    End If
                                '見積書 宛名２
    If GetIni(App.EXEName, "NAME2", App.EXEName, c) Then
        EX_NAME2 = ""
    Else
        EX_NAME2 = Trim(c)
    End If
                                '見積書 自社　名称
    If GetIni(App.EXEName, "SYAMEI", App.EXEName, c) Then
        EX_SYAMEI = ""
    Else
        EX_SYAMEI = Trim(c)
    End If
                                '見積書 自社　住所１
    If GetIni(App.EXEName, "ADDR1", App.EXEName, c) Then
        EX_ADDR1 = ""
    Else
        EX_ADDR1 = Trim(c)
    End If
                                '見積書 自社　住所２
    If GetIni(App.EXEName, "ADDR2", App.EXEName, c) Then
        EX_ADDR2 = ""
    Else
        EX_ADDR2 = Trim(c)
    End If
                                '見積書 自社　備考
    If GetIni(App.EXEName, "BIKOU1", App.EXEName, c) Then
        EX_BIKOU1 = ""
    Else
        EX_BIKOU1 = Trim(c)
    End If
                                '見積書 センター   名称
    If GetIni(App.EXEName, "CENTER_NAME", App.EXEName, c) Then
        EX_CENTER_NAME = ""
    Else
        EX_CENTER_NAME = Trim(c)
    End If
                                '見積書 センター   住所１
    If GetIni(App.EXEName, "CENTER_ADDR1", App.EXEName, c) Then
        EX_CENTER_ADDR1 = ""
    Else
        EX_CENTER_ADDR1 = Trim(c)
    End If
                                '見積書 センター   備考１
    If GetIni(App.EXEName, "CENTER_BIKOU1", App.EXEName, c) Then
        EX_CENTER_BIKOU1 = ""
    Else
        EX_CENTER_BIKOU1 = Trim(c)
    End If
                                
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '構成マスタＯＰＥＮ
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '月平均出荷数(月別集計)ＯＰＥＮ
    If MONTHLYQTY_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '入出庫単価マスタＯＰＥＮ
    If SE_LOC_TANKA_M_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目単価変更履歴ＯＰＥＮ
    If ITEM_HST_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '管理マスタＯＰＥＮ
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If

    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case Else
            Call File_Error(sts, BtOpGetEqual, "管理マスタ(KEY=01)")
        Unload Me
    End Select

    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_DEF_No)
    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC02, Len(P_KANRIREC02), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case Else
            Call File_Error(sts, BtOpGetEqual, "管理マスタ(KEY=02)")
        Unload Me
    End Select


    If IsNumeric(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) Then
        YOYU_RITU(0) = Format(CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)), "#0.00")
        YOYU_RITU(1) = Format(CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)), "#0.00")
        YOYU_RITU(2) = Format(CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)), "#0.00")
        YOYU_RITU(3) = Format(CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)), "#0.00")
    Else
        YOYU_RITU(0) = "1.00"
        YOYU_RITU(1) = "1.00"
        YOYU_RITU(2) = "1.00"
        YOYU_RITU(3) = "1.00"
    End If

    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc
    
    '仕向け先のセット
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD, 0) Then
        Unload Me
    End If
    Combo1(pcmbSHIMUKE).ListIndex = 0

    '種別セット
    If SYUBETSU_Set_Proc() Then
        Unload Me
    End If

    CNV_SEI00101.Caption = CNV_SEI00101.Caption
    

    Call Init_Proc

    Show

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
                                            '構成マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "構成マスタ")
        End If
    End If
                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "構成マスタ")
        End If
    End If
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
                                            '管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "管理マスタ")
        End If
    End If
    
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    End
End Sub

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    CNV_SEI00101.MousePointer = vbHourglass

    Call Ctrl_Lock(CNV_SEI00101)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(CNV_SEI00101)


    CNV_SEI00101.MousePointer = vbDefault

End Sub


Private Sub SHORI_Click(Index As Integer)
    Select Case Index
        Case 0 To 5
            Command1(Index).Value = True

        Case 6      '画面印刷
        
            Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)

    End Select
                    
    
    


End Sub






Private Function Init_Proc() As Integer
'----------------------------------------------------------------------------
'                   画面初期化
'----------------------------------------------------------------------------
Dim i           As Integer

Dim Row         As Integer
Dim KOTEI_NO    As Integer

Dim c           As String * 128
                                
Dim wkKOTEI As Variant
                                
                                
                                
                                
                                
    Init_Proc = True
                                
                                
    If SYUBETSU_Set_Proc() Then
        Exit Function
    End If
                                
                                
                                
                                '作業工程情報取り込み
'    Set SAGYO = Nothing
    
    
    
    
    
'    Text1(ptxDEF_LOT).Text = Format(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode), "#0")
    
    
    
    
    Row = 0
    KOTEI_NO = 0
    For i = 1 To 10
        
        If GetIni("KOUTEI", "BEF" & Format(i, "00"), "SEI_SYS", c) Then
        Else
            If Trim(c) = "*" Then
            Else
                Row = Row + 1
'                SAGYO.ReDim S_Min_Row, Row, S_Min_Col, S_Max_Col
                
                KOTEI_NO = KOTEI_NO + 10
'                SAGYO(Row, ColS_No) = Format(KOTEI_NO, "00")
                
                
                
                wkKOTEI = Split(Trim(c), ",", -1)
                
'                SAGYO(Row, ColS_NAME) = Trim(wkKOTEI(0))
'                SAGYO(Row, ColS_KOUSU) = Trim(wkKOTEI(1))
                
                If IsNumeric(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)) Then
'                    SAGYO(Row, ColS_TANKA) = Format(SAGYO(Row, ColS_KOUSU) * CDbl(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#,##0.00")
                Else
'                    SAGYO(Row, ColS_TANKA) = 0
                End If
                
                If IsNumeric(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)) Then
'                    SAGYO(Row, ColS_KIN) = Format(SAGYO(Row, ColS_TANKA) * CLng(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)), "#,##0")
                Else
'                    SAGYO(Row, ColS_KIN) = 0
                End If
            End If
        End If
    
    Next i
    
    For i = 1 To 10
        
        If GetIni("KOUTEI", "MAIN" & Format(i, "00"), "SEI_SYS", c) Then
        Else
            If Trim(c) = "*" Then
            Else
                Row = Row + 1
'                SAGYO.ReDim S_Min_Row, Row, S_Min_Col, S_Max_Col
                
                KOTEI_NO = KOTEI_NO + 10
'                SAGYO(Row, ColS_No) = Format(KOTEI_NO, "00")
                
                wkKOTEI = Split(Trim(c), ",", -1)
                
'                SAGYO(Row, ColS_NAME) = Trim(wkKOTEI(0))
'                SAGYO(Row, ColS_KOUSU) = Trim(wkKOTEI(1))
                
                If IsNumeric(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)) Then
'                    SAGYO(Row, ColS_TANKA) = Format(SAGYO(Row, ColS_KOUSU) * CDbl(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#,##0.00")
                Else
'                    SAGYO(Row, ColS_TANKA) = 0
                End If
                If IsNumeric(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)) Then
'                    SAGYO(Row, ColS_KIN) = Format(SAGYO(Row, ColS_TANKA) * CLng(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)), "#,##0")
                Else
'                    SAGYO(Row, ColS_KIN) = 0
                End If
            End If
        End If
    
    Next i
                                
    For i = 1 To 10
        
        If GetIni("KOUTEI", "AFT" & Format(i, "00"), "SEI_SYS", c) Then
        Else
            If Trim(c) = "*" Then
            Else
                Row = Row + 1
'                SAGYO.ReDim S_Min_Row, Row, S_Min_Col, S_Max_Col
                
                KOTEI_NO = KOTEI_NO + 10
'                SAGYO(Row, ColS_No) = Format(KOTEI_NO, "00")
                
                wkKOTEI = Split(Trim(c), ",", -1)
                
'                SAGYO(Row, ColS_NAME) = Trim(wkKOTEI(0))
'                SAGYO(Row, ColS_KOUSU) = Trim(wkKOTEI(1))
                
                If IsNumeric(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)) Then
'                    SAGYO(Row, ColS_TANKA) = Format(SAGYO(Row, ColS_KOUSU) * CDbl(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#,##0.00")
                Else
'                    SAGYO(Row, ColS_TANKA) = 0
                End If
                
                If IsNumeric(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)) Then
'                    SAGYO(Row, ColS_KIN) = Format(SAGYO(Row, ColS_TANKA) * CLng(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)), "#,##0")
                Else
'                    SAGYO(Row, ColS_KIN) = 0
                End If
            End If
        End If
    
    Next i
                                
                                
                                
'    Set TDBGrid1(pGrdSAGYO).Array = SAGYO
    
    
'    TDBGrid1(pGrdSAGYO).Bookmark = Null
    
'    TDBGrid1(pGrdSAGYO).ReBind
'    TDBGrid1(pGrdSAGYO).Update
'    TDBGrid1(pGrdSAGYO).ScrollBars = dbgAutomatic

    Init_Proc = True


End Function
Private Function SYUBETSU_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   コードマスタをドロップダウンリストにセットする。
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer



Dim i           As Integer
    
    SYUBETSU_Set_Proc = True
    
    Set SYUBETSU = Nothing
    
    
    
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")

    com = BtOpGetGreater

    i = 0
    Do
        DoEvents
    
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> P_KBN06_CD Then
                    
                    Exit Do
                
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "コードマスタ")
                Exit Function
        
        End Select

        
        i = i + 1
        SYUBETSU.ReDim 1, i, 0, 0
        
        
        SYUBETSU(i, 0) = Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode)) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), 2)
        
        
        com = BtOpGetNext
    
    Loop

    Set TDBDropDown1.Array = SYUBETSU
    TDBDropDown1.ReBind

    SYUBETSU_Set_Proc = False
    



End Function



Private Sub TDBGrid1_AfterColUpdate(Index As Integer, ByVal ColIndex As Integer)

Dim sts         As Integer
Dim Bookmark    As Variant
    
    
Dim i           As Integer
    
    
Dim wkDouble    As Double
    
    
    Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
    TDBGrid1(pGrdKOUSEI).Update
    
    
    
    If TDBGrid1(pGrdKOUSEI).Bookmark = Null Then
        Exit Sub
    End If
    
    If TDBGrid1(pGrdKOUSEI).Bookmark <= 0 Then
        Exit Sub
    End If
    
                
        
        
        
        
    If Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_GAI)) = "" Then
        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU) = ""
        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_NAIGAI) = ""
        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_NAME) = ""
        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = ""
        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN) = ""
        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = ""
        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN) = ""
        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = ""
        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColS_KOUSU) = ""
        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColSEI_SYU_KON) = ""
    
    
    
    Else
        '品番
        If Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU)) = "" And _
            Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_NAIGAI)) = "" Then
            Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
            Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
        Else
            Call UniCode_Conv(K0_ITEM.JGYOBU, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU))
            Call UniCode_Conv(K0_ITEM.NAIGAI, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_NAIGAI))
        End If
        
        Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_GAI))
    
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                '資材品で読み替え
                Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_GAI))
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        
                        If HIN_INV Then
                            '未登録品番　可　資材としておく
                            Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                            Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                            Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番")
                            Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                        Else
                            MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]行目 入力した項目はエラーです。(品番)"
                            Exit Sub
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                        Unload Me
                
                End Select
        
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                Unload Me
        
        End Select
    
        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU) = StrConv(ITEMREC.JGYOBU, vbUnicode)
        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_NAIGAI) = StrConv(ITEMREC.NAIGAI, vbUnicode)
        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    
    
        '員数
        If Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) = "" Then
            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = "1.00"
        End If
        If IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) Then
            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = Format(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)), "#0.00")
        Else
            MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]行目 入力した項目はエラーです。(員数)"
            Exit Sub
        End If
    
    
        '仕入＠
        If Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN)) = "" Then
            If IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)), "#0.00")
            Else
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN) = "0.00"
            End If
        Else
            If IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN)) Then
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN) = Format(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN)), "#0.00")
            Else
                MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]行目 入力した項目はエラーです。(仕入＠)"
                Exit Sub
        
            End If
        End If
        
        '仕入金額計
        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN) = 0
    
        For i = 0 To UBound(SHIZAI_T)
            If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = SHIZAI_T(i) Then
                
                
                If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                    
                    
                    If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN)) / CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY))), 2), "#,##0.00")
                    
                    Else
                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN))), 2), "#,##0.00")
                    End If
                    
                End If
                Exit For
            End If
        
        Next i
        If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN)) = 0 Then
            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN) = ""
        End If
    
        '販売＠
        If Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) = "" Then
            If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
            Else
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = "0.00"
            End If
        Else
            If IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) Then
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = Format(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)), "#0.00")
            Else
                MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]行目 入力した項目はエラーです。(販売＠)"
                Exit Sub
            End If
        End If
        '売上金額計
        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = 0
    
        For i = 0 To UBound(SHIZAI_T)
        
            If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = SHIZAI_T(i) Then
                If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                    If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = GAISO_KBN Then
                        
                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) / CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY))), 2), "#,##0.00")
                    Else
                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN))), 2), "#,##0.00")
                    End If
                End If
            Else
            
                If KUSATU_F Then
                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
            
                        If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = GAISO_KBN Then
                        
                            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN_KUSATU) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) / CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY))), 2), "#,##0.00")
                        Else
                            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN_KUSATU) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN))), 2), "#,##0.00")
                        End If
                    End If
                
                End If
                
            End If
        
        Next i
        
        If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN)) = 0 Then
            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = ""
        End If
        
        
        
        
        
        
        
        
        
        '作業時間
        If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = GAISO_KBN Then
            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColS_KOUSU) = ""
            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColSEI_SYU_KON) = ""
        Else
        
            If Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColS_KOUSU)) = "" Then
                If IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColS_KOUSU) = Format(CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode)), "#")
                Else
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColS_KOUSU) = ""
                End If
            Else
                If IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColS_KOUSU)) Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColS_KOUSU) = Format(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColS_KOUSU)), "#")
                Else
                    MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]行目 入力した項目はエラーです。(作業時間)"
                End If
            End If
            
            
            '集合梱包時間
            If Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColSEI_SYU_KON)) = "" Then
                If IsNumeric(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)) Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColSEI_SYU_KON) = Format(CDbl(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)), "#")
                Else
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColSEI_SYU_KON) = ""
                End If
            Else
                If IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColSEI_SYU_KON)) Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColSEI_SYU_KON) = Format(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColSEI_SYU_KON)), "#")
                Else
                    MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]行目 入力した項目はエラーです。(集合梱包時間)"
                End If
            End If
    
        End If
    End If
        
    Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
        
    
    TDBGrid1(pGrdKOUSEI).Refresh
    TDBGrid1(pGrdKOUSEI).Update
'    TDBGrid1.ScrollBars = dbgAutomatic
    
    TDBGrid1(pGrdKOUSEI).SetFocus



End Sub


Private Sub TDBGrid1_BeforeInsert(Index As Integer, Cancel As Integer)
    
    KOUSEI.ReDim Min_Row, KOUSEI.Count(1), Min_Col, Max_Col

End Sub

Private Sub Text1_Change(Index As Integer)
Dim i   As Integer
    
    
    If ptxHin_Gai = Index Then
        If Trim(svHin_Gai) = (Text1(ptxHin_Gai).Text) Then
        Else
            For i = 2 To 5
                Command1(i).Enabled = False
            Next i
        
        
            Text1(ptxMAIN_KOUTEI_QTY01).Text = ""
        
        
        End If
    
    
    
    
    End If



End Sub

Private Sub Text1_GotFocus(Index As Integer)
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If


    If Index = ptxHin_Gai Then
        svHin_Gai = Text1(ptxHin_Gai).Text
    End If



End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    If Error_Check_Proc(Index) Then   'エラーチェック
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動
End Sub
Private Function Error_Check_Proc(Mode As Integer) As Integer
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
        
Dim INV_F       As Boolean
    
    
    Error_Check_Proc = True
    
    Select Case Mode
    
    
        Case ptxTanto_Code     '担当者
        
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTanto_Code).Text)

            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    Text1(ptxTanto_Name).Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    Text1(ptxTanto_Name).Text = ""
            
                    MsgBox "入力した項目はエラーです。(担当者)"
                    Text1(Mode).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Exit Function
                
            
            
            End Select
        Case ptxHin_Gai         '品番
    
            
            Text1(ptxHin_Gai).Text = Trim(StrConv(Text1(ptxHin_Gai).Text, vbUpperCase))
            
            Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
            Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)


            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                    Text1(ptxHin_Name).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                
                    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                        Text1(ptxST_SOKO).Text = ""
                        Text1(ptxST_RETU).Text = ""
                        Text1(ptxST_REN).Text = ""
                        Text1(ptxST_DAN).Text = ""
                    Else
                        Text1(ptxST_SOKO).Text = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                        Text1(ptxST_RETU).Text = StrConv(ITEMREC.ST_RETU, vbUnicode)
                        Text1(ptxST_REN).Text = StrConv(ITEMREC.ST_REN, vbUnicode)
                        Text1(ptxST_DAN).Text = StrConv(ITEMREC.ST_DAN, vbUnicode)
                    End If
                
                
                Case BtErrKeyNotFound

                    Text1(ptxHin_Name).Text = ""

                    MsgBox "入力した項目はエラーです。(品番)"
                    Text1(Mode).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function

            End Select
        
        
        
        
            INV_F = False
            Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
            Select Case sts
                Case BtNoErr
        
                    Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, StrConv(SOKOREC.IO_TANKA_No, vbUnicode))
                    sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
                    Select Case sts
                        Case BtNoErr
                        
                            Text1(ptxIO_TANKA_No).Text = StrConv(SE_LOC_TANKA_M_REC.SE_IO_TANKA_No, vbUnicode)
                            Text1(ptxSE_Name).Text = StrConv(SE_LOC_TANKA_M_REC.SE_Name, vbUnicode)
                        
                        
                        Case BtErrKeyNotFound
                
                            INV_F = True
                
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                        Exit Function
                    End Select
        
                Case BtErrKeyNotFound
        
                    INV_F = True
        
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                    Exit Function
    
            End Select
    
    
            If INV_F Then
                
                Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, INV_IO_TANKA_No)
                sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    
                            Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_Name, "")
                    
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                        Exit Function
                End Select
            
            
                Text1(ptxIO_TANKA_No).Text = INV_IO_TANKA_No
                Text1(ptxSE_Name).Text = ""
            
            End If
        
        
        
        
        
        
        
        
        
        
        Case ptxAFT_SEI_LOT         'ロット数
            
            If Text1(ptxAFT_SEI_LOT).Text = "" Then
                Text1(ptxAFT_SEI_LOT).Text = "1"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_SEI_LOT).Text) Then
                MsgBox "入力した項目はエラーです。(ロット数)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_SEI_LOT).Text = Format(CLng(Text1(ptxAFT_SEI_LOT).Text), "#0")
            End If
        
        Case ptxAFT_SEI_RATE        '分レート
            
            If Text1(ptxAFT_SEI_RATE).Text = "" Then
                Text1(ptxAFT_SEI_RATE).Text = "0"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_SEI_RATE).Text) Then
                MsgBox "入力した項目はエラーです。(分レート)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_SEI_RATE).Text = Format(CLng(Text1(ptxAFT_SEI_RATE).Text), "#0")
            End If
    
        Case ptxAFT_S_KOUSU         '工数
            
            If Text1(ptxAFT_S_KOUSU).Text = "" Then
                Text1(ptxAFT_S_KOUSU).Text = "0.0"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_SEI_RATE).Text) Then
                MsgBox "入力した項目はエラーです。(工数)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_S_KOUSU).Text = Format(CDbl(Text1(ptxAFT_S_KOUSU).Text), "#0.0")
            End If
    
    
        Case ptxAFT_S_KOUSU_GENKA   '工料原価
            
            If Text1(ptxAFT_S_KOUSU_GENKA).Text = "" Then
                Text1(ptxAFT_S_KOUSU_GENKA).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_S_KOUSU_GENKA).Text) Then
                MsgBox "入力した項目はエラーです。(工料原価)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_S_KOUSU_GENKA).Text = Format(CDbl(Text1(ptxAFT_S_KOUSU_GENKA).Text), "#0.00")
            End If
        
        Case ptxAFT_S_KOUSU_BAIKA   '工料売価
            
            If Text1(ptxAFT_S_KOUSU_BAIKA).Text = "" Then
                Text1(ptxAFT_S_KOUSU_BAIKA).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_S_KOUSU_BAIKA).Text) Then
                MsgBox "入力した項目はエラーです。(工料売価)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxAFT_S_KOUSU_BAIKA).Text), "#0.00")
            End If
    
    
    
    
        Case ptxAFT_S_SHIZAI_GENKA   '資材原価
            
            If Text1(ptxAFT_S_SHIZAI_GENKA).Text = "" Then
                Text1(ptxAFT_S_SHIZAI_GENKA).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_S_SHIZAI_GENKA).Text) Then
                MsgBox "入力した項目はエラーです。(資材原価)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_S_SHIZAI_GENKA).Text = Format(CDbl(Text1(ptxAFT_S_SHIZAI_GENKA).Text), "#0.00")
            End If
    
    
        Case ptxAFT_S_SHIZAI_BAIKA  '資材売価
            
            If Text1(ptxAFT_S_SHIZAI_BAIKA).Text = "" Then
                Text1(ptxAFT_S_SHIZAI_BAIKA).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_S_SHIZAI_BAIKA).Text) Then
                MsgBox "入力した項目はエラーです。(資材売価)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_S_SHIZAI_BAIKA).Text = Format(CDbl(Text1(ptxAFT_S_SHIZAI_BAIKA).Text), "#0.00")
            End If
    
    
    
        Case ptxAFT_S_PPSC_KAKO_KOSU    'PPSC加工単価
        
        
            If Text1(ptxAFT_S_PPSC_KAKO_KOSU).Text = "" Then
                Text1(ptxAFT_S_PPSC_KAKO_KOSU).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_S_PPSC_KAKO_KOSU).Text) Then
                MsgBox "入力した項目はエラーです。(PPSC加工単価)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_S_PPSC_KAKO_KOSU).Text = Format(CDbl(Text1(ptxAFT_S_PPSC_KAKO_KOSU).Text), "#0.00")
            End If
        
        
        
        
        Case ptxAFT_S_BU_KAKO_KOSU      'BU加工単価
    
            If Text1(ptxAFT_S_BU_KAKO_KOSU).Text = "" Then
                Text1(ptxAFT_S_BU_KAKO_KOSU).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_S_BU_KAKO_KOSU).Text) Then
                MsgBox "入力した項目はエラーです。(PPSC加工単価)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_S_BU_KAKO_KOSU).Text = Format(CDbl(Text1(ptxAFT_S_BU_KAKO_KOSU).Text), "#0.00")
            End If
    
    
    
        Case ptxAFT_SEI_TANKA_TANTO     '担当者
        
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxAFT_SEI_TANKA_TANTO).Text)

            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
            
                    MsgBox "入力した項目はエラーです。(担当者)"
                    Text1(Mode).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Exit Function
                
            
            
            End Select
    
        Case ptxAFT_SE_TANKA_MEMO       'メモ
        
        Case ptxMAIN_KOUTEI_QTY01       'ラベル貼り付け枚数
            
            If Not IsNumeric(Text1(ptxMAIN_KOUTEI_QTY01).Text) Then
                Text1(ptxMAIN_KOUTEI_QTY01).Text = "1"
            Else
                Text1(ptxMAIN_KOUTEI_QTY01).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_QTY01).Text), "#0")
            End If
    
    

    
            If IsNumeric(Text1(ptxMAIN_KOUTEI_TANI01)) Then
                Text1(ptxMAIN_KOUTEI_KOUSU01).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI01).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY01).Text), "#0")
            End If
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function


Private Function Code_Set_Proc(Index As Integer, KBN As String, Mode As Integer) As Integer
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
            wkOption = Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
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


Private Function P_COMPO_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   構成マスタの読み込み＆表示
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
    
Dim Row         As Long
    
Dim FAST_FLG    As Boolean
    
    P_COMPO_Disp_Proc = True
    Call Input_Lock             '2008.01.15
    
        
    
    
            

    

    Set KOUSEI = Nothing

    
    
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHin_Gai).Text)
       
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_HEAD)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
       
    sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        
    Select Case sts
        Case BtNoErr
        
            FAST_FLG = True
        
            '備考
            RichTextBox1(prchBIKOU).Text = RTrim(StrConv(P_COMPO_O_REC.BIKOU, vbUnicode))
        
            '商品化ｸﾗｽ
            Text1(ptxS_CLASS_CODE).Text = Trim(StrConv(P_COMPO_O_REC.CLASS_CODE, vbUnicode))
            '付加ｸﾗｽ
            Text1(ptxF_CLASS_CODE).Text = Trim(StrConv(P_COMPO_O_REC.F_CLASS_CODE, vbUnicode))
            '内職ｸﾗｽ
            Text1(ptxN_CLASS_CODE).Text = Trim(StrConv(P_COMPO_O_REC.N_CLASS_CODE, vbUnicode))

        
        Case BtErrKeyNotFound
            
            FAST_FLG = False
            
            '備考
            RichTextBox1(prchBIKOU).Text = ""
        
            '商品化ｸﾗｽ
            Text1(ptxS_CLASS_CODE).Text = ""
            '付加ｸﾗｽ
            Text1(ptxF_CLASS_CODE).Text = ""
            '内職ｸﾗｽ
            Text1(ptxN_CLASS_CODE).Text = ""
        
        
        Case Else
            
            Set KOUSEI = Nothing
            
            
            Call Input_UnLock           '2008.01.15
            P_COMPO_Disp_Proc = sts
            Exit Function
    End Select

    '--------------------------------   「子」情報
    
    Set KOUSEI = Nothing
    
    
    
    If FAST_FLG Then
    
        Row = Min_Row - 1
        
        Do
            DoEvents
            
            sts = BTRV(BtOpGetNext, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                
                                
                    If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                        Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHin_Gai).Text) Then
                    
                        Exit Do
                
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call Input_UnLock             '2008.01.15
                    Call File_Error(sts, BtOpGetNext, "構成マスタ")
                    Exit Function
            End Select
            
            If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_KOSOU Then
                Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, KOSOU_KBN)
            End If
            If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_GAISOU Then
                Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, GAISO_KBN)
            End If
            
            Row = Row + 1
                        
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
            
            
            
        Loop
    End If

    Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
    
    
    TDBGrid1(pGrdKOUSEI).Bookmark = Null
    
    TDBGrid1(pGrdKOUSEI).ReBind
    TDBGrid1(pGrdKOUSEI).Update
    TDBGrid1(pGrdKOUSEI).ScrollBars = dbgAutomatic
    
    If KOUSEI.Count(1) > 0 Then
        TDBGrid1(pGrdKOUSEI).MoveFirst
    End If















    Call Input_UnLock

    
    
    P_COMPO_Disp_Proc = False

End Function
Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'                   構成マスタ==>Gridテーブル
'----------------------------------------------------------------------------

Dim sts As Integer
Dim i   As Integer
Dim j   As Integer
    
    Grid_Set_Proc = True

    

    KOUSEI.ReDim Min_Row, Row, Min_Col, Max_Col
    
    
    '事業部
    KOUSEI(Row, ColKO_JGYOBU) = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
    '国内外
    KOUSEI(Row, ColKO_NAIGAI) = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
    
    '種別
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        
    Select Case sts
        Case BtNoErr
            KOUSEI(Row, ColKO_SYUBETSU) = Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode)) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), 2)
        
        
        
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "コードマスタ")
            Exit Function
    
    End Select
    '品番
    KOUSEI(Row, ColKO_HIN_GAI) = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
    '品名
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Select Case sts
        Case BtNoErr
            KOUSEI(Row, ColKO_HIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
        Case BtErrKeyNotFound
            KOUSEI(Row, ColKO_HIN_NAME) = "未登録品番"
            
            Call UniCode_Conv(ITEMREC.S_KOUSU, "00000000")
        
            Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "00000000.00")
            Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
        
        
            Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "000.00")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    
    End Select
    '員数
    If IsNumeric(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) Then
        KOUSEI(Row, ColKO_QTY) = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
    Else
        KOUSEI(Row, ColKO_QTY) = "1.00"
    End If
    
    '仕入単価
    If IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
        KOUSEI(Row, ColG_ST_SHITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)), "#0.00")
    Else
        KOUSEI(Row, ColG_ST_SHITAN) = "0.00"
    End If
    '販売単価
    If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
        KOUSEI(Row, ColG_ST_URITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
    Else
        KOUSEI(Row, ColG_ST_URITAN) = "0.00"
    End If
    '仕入金額計
    KOUSEI(Row, ColG_ST_SHIKIN) = 0

    For i = 0 To UBound(SHIZAI_T)
        If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = SHIZAI_T(i) Then
            
            
            If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                
                If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                    
                    If CDbl(KOUSEI(Row, ColKO_QTY)) = 0 Then
                        KOUSEI(Row, ColG_ST_SHIKIN) = 0
                    Else
                        KOUSEI(Row, ColG_ST_SHIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(Row, ColG_ST_SHITAN)) / CDbl(KOUSEI(Row, ColKO_QTY))), 2), "#,##0.00")
                    End If
                Else
                    KOUSEI(Row, ColG_ST_SHIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(Row, ColKO_QTY)) * CDbl(KOUSEI(Row, ColG_ST_SHITAN))), 2), "#,##0.00")
                End If
            End If
            Exit For
        End If
    
    Next i
    If CDbl(KOUSEI(Row, ColG_ST_SHIKIN)) = 0 Then
        KOUSEI(Row, ColG_ST_SHIKIN) = ""
    End If
    
    '売上金額計
    KOUSEI(Row, ColG_ST_URIKIN) = 0
    KOUSEI(Row, ColG_ST_URIKIN_KUSATU) = 0

    For i = 0 To UBound(SHIZAI_T)
    
        If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
    
    
            If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = SHIZAI_T(i) Then
                If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                
                    If CDbl(KOUSEI(Row, ColKO_QTY)) = 0 Then
                        KOUSEI(Row, ColG_ST_URIKIN) = 0
                    Else
                        KOUSEI(Row, ColG_ST_URIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(Row, ColG_ST_URITAN)) / CDbl(KOUSEI(Row, ColKO_QTY))), 2), "#,##0.00")
                    End If
                Else
                    KOUSEI(Row, ColG_ST_URIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(Row, ColKO_QTY)) * CDbl(KOUSEI(Row, ColG_ST_URITAN))), 2), "#,##0.00")
                End If
    
            
            Else
            
                If KUSATU_F Then
            
                    If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                    
                        If CDbl(KOUSEI(Row, ColKO_QTY)) = 0 Then
                            KOUSEI(Row, ColG_ST_URIKIN_KUSATU) = 0
                        Else
                            KOUSEI(Row, ColG_ST_URIKIN_KUSATU) = Format(ToRoundUp(CCur(CDbl(KOUSEI(Row, ColG_ST_URITAN)) / CDbl(KOUSEI(Row, ColKO_QTY))), 2), "#,##0.00")
                        End If
                    Else
                        KOUSEI(Row, ColG_ST_URIKIN_KUSATU) = Format(ToRoundUp(CCur(CDbl(KOUSEI(Row, ColKO_QTY)) * CDbl(KOUSEI(Row, ColG_ST_URITAN))), 2), "#,##0.00")
                    End If
                
                
                End If
            
            
            
            End If
        End If
    Next i
    
    
    If CDbl(KOUSEI(Row, ColG_ST_URIKIN)) = 0 Then
        KOUSEI(Row, ColG_ST_URIKIN) = ""
    End If
    
    
    If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
        KOUSEI(Row, ColS_KOUSU) = ""
        KOUSEI(Row, ColSEI_SYU_KON) = ""
    Else
        '作業時間
        If IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
            KOUSEI(Row, ColS_KOUSU) = Format(CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode)), "#")
        Else
            KOUSEI(Row, ColS_KOUSU) = ""
        End If
        '集合梱包
        If IsNumeric(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)) Then
            KOUSEI(Row, ColSEI_SYU_KON) = Format(CInt(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)), "#")
        Else
            KOUSEI(Row, ColSEI_SYU_KON) = ""
        End If
    End If
    
    
    '備考
    KOUSEI(Row, ColKO_BIKOU) = StrConv(P_COMPO_K_REC.KO_BIKOU, vbUnicode)
    
    
    
    
    Grid_Set_Proc = False
End Function

' ------------------------------------------------------------------------
'       指定した精度の数値に切り上げします。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に切り上げられた数値。
' ------------------------------------------------------------------------
Private Function ToRoundUp(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
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





' ------------------------------------------------------------------------
'       指定した精度の数値に四捨五入します。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に四捨五入された数値。
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






Private Function COVER_Proc() As Integer
'----------------------------------------------------------------------------
'                   ＥＸＣＥＬ（御見積書）出力
'----------------------------------------------------------------------------
Dim excelApplication    As Excel.Application
Dim excelWorkBook       As Excel.Workbook
Dim excelSheet          As Excel.Worksheet

    

    COVER_Proc = True
    
    Call Input_Lock
    



    
    Set excelApplication = CreateObject("Excel.Application")
    excelApplication.Visible = True
    
    Set excelWorkBook = excelApplication.Workbooks.Add
    Set excelSheet = excelWorkBook.Worksheets(1)
    

    
    
    excelApplication.StandardFontSize = 13
    
    excelApplication.StandardFont = "ＭＳ Ｐゴシック"

    
    
    'ページ設定
    With excelSheet.Application.ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    
    '列の幅
    excelSheet.Application.Cells.Select
    excelSheet.Application.Selection.ColumnWidth = 6.25
    excelSheet.Application.Columns(11).Select
    excelSheet.Application.Selection.ColumnWidth = 7.13

    '１行目
    excelSheet.Application.Rows(1).Select
    excelSheet.Application.Selection.RowHeight = 28.5
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 5), excelSheet.Application.Cells(1, 9)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 5), excelSheet.Application.Cells(1, 9)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 5), excelSheet.Application.Cells(1, 9)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 5), excelSheet.Application.Cells(1, 9)).Select
    With excelSheet.Application.Selection.Font
        .Size = 24
    End With
    excelSheet.Application.Cells(1, 5).Value = "御 見 積 書"
    
    '２行目
    excelSheet.Application.Rows(2).Select
    excelSheet.Application.Selection.RowHeight = 13.5
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 11), excelSheet.Application.Cells(2, 14)).HorizontalAlignment = xlRight
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 11), excelSheet.Application.Cells(2, 14)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 11), excelSheet.Application.Cells(2, 14)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 11), excelSheet.Application.Cells(2, 14)).Select
    
    With excelSheet.Application.Selection.Font
        .Size = 11
        .Underline = xlUnderlineStyleSingle
    End With
    excelSheet.Application.Cells(2, 11).NumberFormatLocal = "yyyy""年""m""月""d""日"";@"
    excelSheet.Application.Cells(2, 11).Value = Date
    
    '３行目
    excelSheet.Application.Rows(3).Select
    excelSheet.Application.Selection.RowHeight = 17.25
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(3, 6)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    excelSheet.Application.Cells(3, 1).Value = EX_NAME1
    
    
    '４行目
    excelSheet.Application.Rows(4).Select
    excelSheet.Application.Selection.RowHeight = 17.25
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(4, 6)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    excelSheet.Application.Cells(4, 1).Value = EX_NAME2
'    excelSheet.Application.Cells(4, 5).Value = "西谷ＧＭ様"
    
    '５行目
    excelSheet.Application.Rows(5).Select
    excelSheet.Application.Selection.RowHeight = 13.5
    With excelSheet.Application.Selection.Font
        .Size = 9
    End With
    excelSheet.Application.Cells(5, 1).Value = "　下記のとおり御見積りいたしましたので、何卒ご用命"
    excelSheet.Application.Cells(5, 14).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(5, 14).Select
    With excelSheet.Application.Selection.Font
        .Size = 11
    End With
    excelSheet.Application.Cells(5, 14).Value = EX_SYAMEI
    '６行目
    excelSheet.Application.Rows(6).Select
    excelSheet.Application.Selection.RowHeight = 13.5
    With excelSheet.Application.Selection.Font
        .Size = 9
    End With
    excelSheet.Application.Cells(6, 1).Value = "賜りますよう宜しくお願い申し上げます。"
    excelSheet.Application.Cells(6, 14).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(6, 14).Select
    With excelSheet.Application.Selection.Font
        .Size = 8
    End With
    excelSheet.Application.Cells(6, 14).Value = EX_ADDR1
    '７行目
    excelSheet.Application.Rows(7).Select
    excelSheet.Application.Selection.RowHeight = 13.5
    excelSheet.Application.Cells(7, 14).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(7, 14).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ Ｐゴシック"
        .Size = 8
    End With
    excelSheet.Application.Cells(7, 14).Value = EX_BIKOU1
    '８行目
    excelSheet.Application.Rows(8).Select
    excelSheet.Application.Selection.RowHeight = 13.5
    excelSheet.Application.Cells(8, 11).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(8, 11).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ Ｐゴシック"
        .Size = 11
    End With
    excelSheet.Application.Cells(8, 11).Value = EX_CENTER_NAME
    '９行目
    excelSheet.Application.Rows(9).Select
    excelSheet.Application.Selection.RowHeight = 13.5
    excelSheet.Application.Cells(9, 11).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(9, 11).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ Ｐゴシック"
        .Size = 7
    End With
    excelSheet.Application.Cells(9, 11).Value = EX_CENTER_ADDR1
    '10行目
    excelSheet.Application.Rows(10).Select
    excelSheet.Application.Selection.RowHeight = 29.5
    excelSheet.Application.Cells(10, 11).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(10, 11).VerticalAlignment = xlTop
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ Ｐゴシック"
        .Size = 7
    End With
    excelSheet.Application.Cells(10, 11).Value = EX_CENTER_BIKOU1
    '９〜10行目
    ActiveSheet.Shapes.AddShape(1, 525#, 117.75, 70.5, 13.5). _
        Select
    Selection.Characters.Text = "承認印"
    With Selection.Characters(Start:=1, Length:=3).Font
        .NAME = "ＭＳ Ｐゴシック"
        .FontStyle = "標準"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .Orientation = xlHorizontal
        .AutoSize = False
        .AddIndent = False
    End With
    
    ActiveSheet.Shapes.AddShape(1, 525#, 131.25, 70.5, 37.5).Select


    ActiveSheet.Shapes.AddShape(1, 595#, 117.75, 70.5, 13.5). _
        Select
    Selection.Characters.Text = "担当印"
    With Selection.Characters(Start:=1, Length:=3).Font
        .NAME = "ＭＳ Ｐゴシック"
        .FontStyle = "標準"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .Orientation = xlHorizontal
        .AutoSize = False
        .AddIndent = False
    End With
    
    ActiveSheet.Shapes.AddShape(1, 595#, 131.25, 70.5, 37.5).Select


    

    Set excelSheet = Nothing
    Set excelWorkBook = Nothing
    Set excelApplication = Nothing

End Function

Private Function Detail_Disp_Proc(com As Integer) As Integer
'----------------------------------------------------------------------------
'                   現在値画面表示
'----------------------------------------------------------------------------
Dim sts         As Integer

Dim i           As Integer
Dim j           As Integer
Dim wkInt       As Integer
Dim wkDouble    As Double

Dim wkKUSATU    As Variant
Dim c           As String * 128


Dim INV_F   As Boolean


    Detail_Disp_Proc = True
    
    '品目マスタ読み込み

    sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrEOF
            Detail_Disp_Proc = sts
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function

    End Select
If RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) = "AXR30-100" Then

Debug.Print
End If
    For i = 0 To Combo1(pcmbSHIMUKE).ListCount - 1
    
        If StrConv(ITEMREC.JGYOBU, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE).List(i), 4), 3, 1) And _
            StrConv(ITEMREC.NAIGAI, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE).List(i), 4), 4, 1) Then
            Combo1(pcmbSHIMUKE).ListIndex = i
            Exit For
        End If
    Next i
    
    
    If i > Combo1(pcmbSHIMUKE).ListCount - 1 Then
        Detail_Disp_Proc = SYS_CANCEL
        Exit Function
    End If
    
    '品番
    Text1(ptxHin_Gai).Text = RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
    
    
    '品名
    Text1(ptxHin_Name).Text = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
    '標準棚番
    Text1(ptxST_SOKO).Text = StrConv(ITEMREC.ST_SOKO, vbUnicode)
    Text1(ptxST_RETU).Text = StrConv(ITEMREC.ST_RETU, vbUnicode)
    Text1(ptxST_REN).Text = StrConv(ITEMREC.ST_REN, vbUnicode)
    Text1(ptxST_DAN).Text = StrConv(ITEMREC.ST_DAN, vbUnicode)
    '-----------------------------------    変更前
    
    
    If IsNumeric(StrConv(ITEMREC.SEI_LOT, vbUnicode)) Then
        Text1(ptxBEF_SEI_LOT).Text = Format(Val(StrConv(ITEMREC.SEI_LOT, vbUnicode)), "#0")
    Else
        Text1(ptxBEF_SEI_LOT).Text = "1"
    End If
    '分ﾚｰﾄ
    If IsNumeric(StrConv(ITEMREC.SEI_RATE, vbUnicode)) Then
        Text1(ptxBEF_SEI_RATE).Text = Format(Val(StrConv(ITEMREC.SEI_RATE, vbUnicode)), "#0")
    Else
        
        If IsNumeric(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)) Then
            Text1(ptxBEF_SEI_RATE).Text = Format(Val(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#0")
        Else
            Text1(ptxBEF_SEI_RATE).Text = ""
        End If
    End If
    
    
    
    
    
    'ﾛｯﾄ数
'    If IsNumeric(StrConv(ITEMREC.SEI_LOT, vbUnicode)) Then
'        Text1(ptxBEF_SEI_LOT).Text = Format(Val(StrConv(ITEMREC.SEI_LOT, vbUnicode)), "#0")
'    Else
'        Text1(ptxBEF_SEI_LOT).Text = ""
'    End If
    '分ﾚｰﾄ
'    If IsNumeric(StrConv(ITEMREC.SEI_RATE, vbUnicode)) Then
'        Text1(ptxBEF_SEI_RATE).Text = Format(Val(StrConv(ITEMREC.SEI_RATE, vbUnicode)), "#0")
'    Else
'        Text1(ptxBEF_SEI_RATE).Text = ""
'    End If
    '工数
    If IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
        Text1(ptxBEF_S_KOUSU).Text = Format(CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode)), "#0.0")
    Else
        Text1(ptxBEF_S_KOUSU).Text = "0.0"
    End If
    '(原価)工料
    If IsNumeric(StrConv(ITEMREC.S_KOUSU_GENKA, vbUnicode)) Then
        Text1(ptxBEF_S_KOUSU_GENKA).Text = Format(CDbl(StrConv(ITEMREC.S_KOUSU_GENKA, vbUnicode)), "#0.00")
    Else
        Text1(ptxBEF_S_KOUSU_GENKA).Text = "0.00"
    End If
    '工料
    If IsNumeric(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) Then
        Text1(ptxBEF_S_KOUSU_BAIKA).Text = Format(CDbl(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)), "#0.00")
    Else
        Text1(ptxBEF_S_KOUSU_BAIKA).Text = "0.00"
    End If
    '(原価)資材
    If IsNumeric(StrConv(ITEMREC.S_SHIZAI_GENKA, vbUnicode)) Then
        Text1(ptxBEF_S_SHIZAI_GENKA).Text = Format(CDbl(StrConv(ITEMREC.S_SHIZAI_GENKA, vbUnicode)), "#0.00")
    Else
        Text1(ptxBEF_S_SHIZAI_GENKA).Text = "0.00"
    End If
    '資材
    If IsNumeric(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)) Then
        Text1(ptxBEF_S_SHIZAI_BAIKA).Text = Format(CDbl(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)), "#0.00")
    Else
        Text1(ptxBEF_S_SHIZAI_BAIKA).Text = "0.00"
    End If
    
    
    
    '外装費
    If IsNumeric(StrConv(ITEMREC.S_GAISO_TANKA, vbUnicode)) Then
        Text1(ptxBEF_S_GAISO_TANKA).Text = Format(CDbl(StrConv(ITEMREC.S_GAISO_TANKA, vbUnicode)), "#0.00")
    Else
        Text1(ptxBEF_S_GAISO_TANKA).Text = "0.00"
    End If
    
    
    'PPSC加工単価
    If IsNumeric(StrConv(ITEMREC.S_PPSC_KAKO_KOSU, vbUnicode)) Then
        Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text = Format(CDbl(StrConv(ITEMREC.S_PPSC_KAKO_KOSU, vbUnicode)), "#0.00")
    Else
        Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text = "0.00"
    End If
    'BU加工単価
    If IsNumeric(StrConv(ITEMREC.S_BU_KAKO_KOSU, vbUnicode)) Then
        Text1(ptxBEF_S_BU_KAKO_KOSU).Text = Format(CDbl(StrConv(ITEMREC.S_BU_KAKO_KOSU, vbUnicode)), "#0.00")
    Else
        Text1(ptxBEF_S_BU_KAKO_KOSU).Text = "0.00"
    End If
    
    
    
    
    
    
    '設定日
    Text1(ptxBEF_S_KOUSU_SET_DATE).Text = Trim(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode))
    '担当者
    Text1(ptxBEF_SEI_TANKA_TANTO).Text = Trim(StrConv(ITEMREC.SEI_TANKA_TANTO, vbUnicode))
    'メモ
    Text1(ptxBEF_SE_TANKA_MEMO).Text = Trim(StrConv(ITEMREC.SE_TANKA_MEMO, vbUnicode))


    '-----------------------------------    変更後
    
    'ﾛｯﾄ数
    If IsNumeric(StrConv(ITEMREC.SEI_LOT, vbUnicode)) Then
        Text1(ptxAFT_SEI_LOT).Text = Format(Val(StrConv(ITEMREC.SEI_LOT, vbUnicode)), "#0")
    Else
        Text1(ptxAFT_SEI_LOT).Text = "1"
    End If
    '分ﾚｰﾄ
    If IsNumeric(StrConv(ITEMREC.SEI_RATE, vbUnicode)) Then
        Text1(ptxAFT_SEI_RATE).Text = Format(Val(StrConv(ITEMREC.SEI_RATE, vbUnicode)), "#0")
    Else
        
        If IsNumeric(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)) Then
            Text1(ptxAFT_SEI_RATE).Text = Format(Val(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#0")
        Else
'2009.01.21            Text1(ptxAFT_SEI_RATE).Text = ""
            Text1(ptxAFT_SEI_RATE).Text = "27"
        End If
    End If
    '設定日
    Text1(ptxAFT_S_KOUSU_SET_DATE).Text = ""
    '担当者
    Text1(ptxAFT_SEI_TANKA_TANTO).Text = Text1(ptxTanto_Code).Text
    'メモ
    Text1(ptxAFT_SE_TANKA_MEMO).Text = ""
    
    '-----------------------------------    月平均出荷数
    If MONTHLYQTY_Disp_Proc() Then
        Exit Function
    End If
    
    '-----------------------------------    構成品表示
    If P_COMPO_Disp_Proc() Then
        Exit Function
    End If
    
    '-----------------------------------    前工程
    
    
    '品目マスタ再読み込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound

            Text1(ptxHin_Name).Text = ""
            Text1(ptxST_SOKO).Text = ""
            Text1(ptxST_RETU).Text = ""
            Text1(ptxST_REN).Text = ""
            Text1(ptxST_DAN).Text = ""
            MsgBox "入力した項目はエラーです。(品番)"
            Detail_Disp_Proc = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function

    End Select
    
    
    
    '�@
    If IsNumeric(StrConv(P_KANRIREC02.BEF_KOTEI(0).KOTEI, vbUnicode)) Then
        Text1(ptxBEF_KOUTEI_TANI01).Text = Format(CInt(StrConv(P_KANRIREC02.BEF_KOTEI(0).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxBEF_KOUTEI_TANI01).Text = "0"
    End If
    Text1(ptxBEF_KOUTEI_QTY01).Text = "1"
    Text1(ptxBEF_KOUTEI_KOUSU01).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI01).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY01).Text), "#0")
    '�A
    wkInt = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    wkInt = wkInt + 1
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    
    If IsNumeric(StrConv(P_KANRIREC02.BEF_KOTEI(1).KOTEI, vbUnicode)) Then
        Text1(ptxBEF_KOUTEI_TANI02).Text = Format(CInt(StrConv(P_KANRIREC02.BEF_KOTEI(1).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxBEF_KOUTEI_TANI02).Text = "0"
    End If
    Text1(ptxBEF_KOUTEI_QTY02).Text = Format(wkInt, "#0")
    Text1(ptxBEF_KOUTEI_KOUSU02).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI02).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY02).Text), "#0")
    '�B
    wkInt = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(DOUKON_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = DOUKON_T(j) Then
                    wkInt = wkInt + 1
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    If IsNumeric(StrConv(P_KANRIREC02.BEF_KOTEI(2).KOTEI, vbUnicode)) Then
        Text1(ptxBEF_KOUTEI_TANI03).Text = Format(CInt(StrConv(P_KANRIREC02.BEF_KOTEI(2).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxBEF_KOUTEI_TANI03).Text = "0"
    End If
    Text1(ptxBEF_KOUTEI_QTY03).Text = Format(wkInt, "#0")
    Text1(ptxBEF_KOUTEI_KOUSU03).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI03).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY03).Text), "#0")
    '�C
    
    
    If KUSATU_F Then
    
        '草津はINI参照
            
        wkInt = 0
        If GetIni(Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1), KUSATU_ETC, App.EXEName, c) Then
        Else
            wkKUSATU = Split(Trim(c), ",", -1)
                    
            If IsNumeric(wkKUSATU(0)) Then
                wkInt = CInt(wkKUSATU(0))
            End If
        End If
        If GetIni(Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1), Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), App.EXEName, c) Then
            If GetIni(App.EXEName, Trim(StrConv(StrConv(ITEMREC.HIN_NAME, vbUnicode), vbWide)), App.EXEName, c) Then
            Else
                wkKUSATU = Split(Trim(c), ",", -1)
                        
                If IsNumeric(wkKUSATU(0)) Then
                    wkInt = CInt(wkKUSATU(0))
                End If
            End If
        Else
            wkKUSATU = Split(Trim(c), ",", -1)
            If IsNumeric(wkKUSATU(0)) Then
                wkInt = CInt(wkKUSATU(0))
            End If
        End If
        Text1(ptxBEF_KOUTEI_TANI04).Text = Format(wkInt, "#0")
        Text1(ptxBEF_KOUTEI_QTY04).Text = "1"
        Text1(ptxBEF_KOUTEI_KOUSU04).Text = Format(CDbl(Text1(ptxBEF_KOUTEI_TANI04).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY04).Text), "#0")

    
    
    
    
    
    
    Else
        '草津以外はマスタ参照
    
        INV_F = False
        Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
            
                Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, StrConv(SOKOREC.IO_TANKA_No, vbUnicode))
                sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    
                        INV_F = True
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                        Exit Function
                End Select
            
            Case BtErrKeyNotFound
            
                INV_F = True
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                Exit Function
        
        End Select
        
        
        If INV_F Then
            
            Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, INV_IO_TANKA_No)
            sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                
                    Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_S_OUT_KOUSU, "000.00")
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                    Exit Function
            End Select
        End If
        If IsNumeric(StrConv(SE_LOC_TANKA_M_REC.SE_S_OUT_KOUSU, vbUnicode)) Then
            Text1(ptxBEF_KOUTEI_TANI04).Text = Format(CInt(StrConv(SE_LOC_TANKA_M_REC.SE_S_OUT_KOUSU, vbUnicode)), "#0")
        Else
            Text1(ptxBEF_KOUTEI_TANI04).Text = "0"
        End If
        Text1(ptxBEF_KOUTEI_QTY04).Text = "1"
        Text1(ptxBEF_KOUTEI_KOUSU04).Text = Format(CDbl(Text1(ptxBEF_KOUTEI_TANI04).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY04).Text), "#0")
    End If
    '�D
    wkInt = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    wkInt = wkInt + 1
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    If IsNumeric(StrConv(P_KANRIREC02.BEF_KOTEI(4).KOTEI, vbUnicode)) Then
        Text1(ptxBEF_KOUTEI_TANI05).Text = Format(CInt(StrConv(P_KANRIREC02.BEF_KOTEI(4).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxBEF_KOUTEI_TANI05).Text = "0"
    End If
    Text1(ptxBEF_KOUTEI_QTY05).Text = Format(wkInt, "#0")
    Text1(ptxBEF_KOUTEI_KOUSU05).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI05).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY05).Text), "#0")
    '�E
    wkInt = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(DOUKON_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = DOUKON_T(j) Then
                    wkInt = wkInt + 1
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    If IsNumeric(StrConv(P_KANRIREC02.BEF_KOTEI(5).KOTEI, vbUnicode)) Then
        Text1(ptxBEF_KOUTEI_TANI06).Text = Format(CInt(StrConv(P_KANRIREC02.BEF_KOTEI(5).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxBEF_KOUTEI_TANI06).Text = "0"
    End If
    Text1(ptxBEF_KOUTEI_QTY06).Text = Format(wkInt, "#0")
    Text1(ptxBEF_KOUTEI_KOUSU06).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI06).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY06).Text), "#0")
    
    '�F
    If IsNumeric(StrConv(P_KANRIREC02.BEF_KOTEI(6).KOTEI, vbUnicode)) Then
        Text1(ptxBEF_KOUTEI_TANI07).Text = Format(CInt(StrConv(P_KANRIREC02.BEF_KOTEI(6).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxBEF_KOUTEI_TANI07).Text = "0"
    End If
    Text1(ptxBEF_KOUTEI_QTY07).Text = "1"
    Text1(ptxBEF_KOUTEI_KOUSU07).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI07).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY07).Text), "#0")
    '�G
    If IsNumeric(StrConv(P_KANRIREC02.BEF_KOTEI(7).KOTEI, vbUnicode)) Then
        Text1(ptxBEF_KOUTEI_TANI08).Text = Format(CInt(StrConv(P_KANRIREC02.BEF_KOTEI(7).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxBEF_KOUTEI_TANI08).Text = "0"
    End If
    Text1(ptxBEF_KOUTEI_QTY08).Text = "1"
    Text1(ptxBEF_KOUTEI_KOUSU08).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI08).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY08).Text), "#0")
    '�H
    If IsNumeric(StrConv(P_KANRIREC02.BEF_KOTEI(8).KOTEI, vbUnicode)) Then
        Text1(ptxBEF_KOUTEI_TANI09).Text = Format(CInt(StrConv(P_KANRIREC02.BEF_KOTEI(8).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxBEF_KOUTEI_TANI09).Text = "0"
    End If
    Text1(ptxBEF_KOUTEI_QTY09).Text = "1"
    Text1(ptxBEF_KOUTEI_KOUSU09).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI09).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY09).Text), "#0")
    '計
    wkInt = 0
    For i = ptxBEF_KOUTEI_KOUSU01 To ptxBEF_KOUTEI_KOUSU09 Step 3
    
        wkInt = wkInt + CInt(Text1(i).Text)
    
    Next i
    Text1(ptxBEF_KOUTEI_KEI1).Text = Format(wkInt, "#0")
    
    
    Text1(ptxBEF_KOUTEI_R_RATE).Text = ToHalfAdjust(CCur(wkInt * CDbl(YOYU_RITU(0).Caption)), 0)
    
    
    
    
    '(秒／個)
    If IsNumeric(Text1(ptxBEF_SEI_LOT).Text) Then
        
        If CInt(Text1(ptxBEF_SEI_LOT).Text) = 0 Then
            Text1(ptxBEF_KOUTEI_KEI2).Text = "0"
        Else
        
            Text1(ptxBEF_KOUTEI_KEI2).Text = Format(ToHalfAdjust(CCur(CInt(Text1(ptxBEF_KOUTEI_R_RATE).Text) / CInt(Text1(ptxBEF_SEI_LOT).Text)), 0), "#0")
        End If
    Else
        Text1(ptxBEF_KOUTEI_KEI2).Text = "0"
    End If
    '(分／個)
    Text1(ptxBEF_KOUTEI_KEI3).Text = Format(ToRoundUp(CCur(CInt(Text1(ptxBEF_KOUTEI_KEI2).Text) / 60), 1), "#0.0")
    '(円／個)
    If IsNumeric(Text1(ptxBEF_SEI_RATE).Text) Then
        Text1(ptxBEF_KOUTEI_KEI4).Text = Format(ToRoundUp(CCur(CDbl(Text1(ptxBEF_KOUTEI_KEI3).Text) * CInt((Text1(ptxBEF_SEI_RATE).Text))), 2), "#0.00")
    Else
        Text1(ptxBEF_KOUTEI_KEI4).Text = "0.00"
    End If
        
    
    
    '-----------------------------------    作業工程
    '�@
    
    If IsNumeric(StrConv(P_KANRIREC02.MAIN_KOTEI(0).KOTEI, vbUnicode)) Then
        Text1(ptxMAIN_KOUTEI_TANI01).Text = Format(CInt(StrConv(P_KANRIREC02.MAIN_KOTEI(0).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxMAIN_KOUTEI_TANI01).Text = "0"
    End If
    
    
''    Text1(ptxMAIN_KOUTEI_QTY01).Text = "1"
    
    
    If IsNumeric(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode)) Then
            Text1(ptxMAIN_KOUTEI_QTY01).Text = Format(CInt(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode)), "#0")
    Else
            Text1(ptxMAIN_KOUTEI_QTY01).Text = "1"
    End If
    Text1(ptxMAIN_KOUTEI_KOUSU01).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI01).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY01).Text), "#0")
    
    
    
    
    
    '�A
    wkInt = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                    
                        wkInt = wkInt + CInt(ToRoundUp(CCur(KOUSEI(i, ColS_KOUSU)) * CDbl(KOUSEI(i, ColKO_QTY)), 0))
                    
                    End If
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI02).Text = Format(wkInt, "#0")
    Text1(ptxMAIN_KOUTEI_QTY02).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU02).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI02).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY02).Text), "#0")
    '�B
    wkInt = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(DOUKON_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = DOUKON_T(j) Then
                    
                    
                    If IsNumeric(KOUSEI(i, ColKO_QTY)) Then
                        wkInt = wkInt + CInt(ToRoundUp(CCur(KOUSEI(i, ColKO_QTY)), 0))
                    End If
                    
                    
                    
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    If IsNumeric(StrConv(P_KANRIREC02.MAIN_KOTEI(3).KOTEI, vbUnicode)) Then
        Text1(ptxMAIN_KOUTEI_TANI03).Text = Format(CInt(StrConv(P_KANRIREC02.MAIN_KOTEI(3).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxMAIN_KOUTEI_TANI03).Text = "0"
    End If
    Text1(ptxMAIN_KOUTEI_QTY03).Text = Format(wkInt, "#0")
    Text1(ptxMAIN_KOUTEI_KOUSU03).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI03).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY03).Text), "#0")
    '�C
    wkInt = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(KAKOU_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = KAKOU_T(j) Then
                    
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkInt = wkInt + CInt(KOUSEI(i, ColS_KOUSU))
                    End If
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI04).Text = Format(wkInt, "#0")
    Text1(ptxMAIN_KOUTEI_QTY04).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU04).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI04).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY04).Text), "#0")
    '�D
    wkInt = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
            
            
            For j = 0 To UBound(SHIZAI_T)
            
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    If IsNumeric(KOUSEI(i, ColSEI_SYU_KON)) Then
                        wkInt = wkInt + CInt(KOUSEI(i, ColSEI_SYU_KON))
                    End If
                End If
            
            Next j
            
            
            
            
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI05).Text = Format(wkInt, "#0")
    Text1(ptxMAIN_KOUTEI_QTY05).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU05).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI05).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY05).Text), "#0")
    '計
    wkInt = 0
    For i = ptxMAIN_KOUTEI_KOUSU01 To ptxMAIN_KOUTEI_KOUSU09 Step 3
    
        If IsNumeric(Text1(i).Text) Then
            wkInt = wkInt + CInt(Text1(i).Text)
        End If
    Next i
    Text1(ptxMAIN_KOUTEI_KEI1).Text = Format(wkInt, "#0")
    
    
    Text1(ptxMAIN_KOUTEI_R_RATE).Text = ToHalfAdjust(CCur(wkInt * CDbl(YOYU_RITU(0).Caption)), 0)
    
    
    '(秒／個)
    If IsNumeric(Text1(ptxBEF_SEI_LOT).Text) Then
        Text1(ptxMAIN_KOUTEI_KEI2).Text = Format(ToHalfAdjust(CCur(Text1(ptxMAIN_KOUTEI_R_RATE).Text), 0), "#0")
    Else
        Text1(ptxMAIN_KOUTEI_KEI2).Text = "0"
    End If
    '(分／個)
    Text1(ptxMAIN_KOUTEI_KEI3).Text = Format(ToRoundUp(CCur(CInt(Text1(ptxMAIN_KOUTEI_KEI2).Text) / 60), 1), "#0.0")
    '(円／個)
    If IsNumeric(Text1(ptxBEF_SEI_RATE).Text) Then
        Text1(ptxMAIN_KOUTEI_KEI4).Text = Format(ToRoundUp(CCur(CDbl(Text1(ptxMAIN_KOUTEI_KEI3).Text) * CInt((Text1(ptxBEF_SEI_RATE).Text))), 2), "#0.00")
    Else
        Text1(ptxMAIN_KOUTEI_KEI4).Text = "0.00"
    End If
    
    '-----------------------------------    後工程
    '�@
    If IsNumeric(StrConv(P_KANRIREC02.AFT_KOTEI(0).KOTEI, vbUnicode)) Then
        Text1(ptxAFT_KOUTEI_TANI01).Text = Format(CInt(StrConv(P_KANRIREC02.AFT_KOTEI(0).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxAFT_KOUTEI_TANI01).Text = "0"
    End If
    Text1(ptxAFT_KOUTEI_QTY01).Text = "1"
    Text1(ptxAFT_KOUTEI_KOUSU01).Text = Format(CInt(Text1(ptxAFT_KOUTEI_TANI01).Text) * CInt(Text1(ptxAFT_KOUTEI_QTY01).Text), "#0")
    '�A
    
    If KUSATU_F Then
        '草津はINIより
    
        wkInt = 0
        If GetIni(Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1), KUSATU_ETC, App.EXEName, c) Then
        Else
            wkKUSATU = Split(Trim(c), ",", -1)
                    
            If IsNumeric(wkKUSATU(1)) Then
                wkInt = CInt(wkKUSATU(1))
            End If
        End If
        If GetIni(Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1), Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), App.EXEName, c) Then
            If GetIni(App.EXEName, Trim(StrConv(StrConv(ITEMREC.HIN_NAME, vbUnicode), vbWide)), App.EXEName, c) Then
            Else
                wkKUSATU = Split(Trim(c), ",", -1)
                        
                If IsNumeric(wkKUSATU(0)) Then
                    wkInt = CInt(wkKUSATU(1))
                End If
            End If
        Else
            wkKUSATU = Split(Trim(c), ",", -1)
                    
            If IsNumeric(wkKUSATU(1)) Then
                wkInt = CInt(wkKUSATU(1))
            End If
        End If
    
        Text1(ptxAFT_KOUTEI_TANI02).Text = Format(wkInt, "#0")
    
        Text1(ptxAFT_KOUTEI_QTY02).Text = "1"
        Text1(ptxAFT_KOUTEI_KOUSU02).Text = Format(CDbl(Text1(ptxAFT_KOUTEI_TANI02).Text) * CInt(Text1(ptxAFT_KOUTEI_QTY02).Text), "#0")
    
    Else
    
        INV_F = False
        Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
            
                Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, StrConv(SOKOREC.IO_TANKA_No, vbUnicode))
                sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    
                        INV_F = True
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                        Exit Function
                End Select
            
            Case BtErrKeyNotFound
            
                INV_F = True
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                Exit Function
        
        End Select
        
        
        If INV_F Then
            
            Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, INV_IO_TANKA_No)
            sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                
                    Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_S_IN_KOUSU, "000.00")
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                    Exit Function
            End Select
        End If
        If IsNumeric(StrConv(SE_LOC_TANKA_M_REC.SE_S_IN_KOUSU, vbUnicode)) Then
            Text1(ptxAFT_KOUTEI_TANI02).Text = Format(CInt(StrConv(SE_LOC_TANKA_M_REC.SE_S_IN_KOUSU, vbUnicode)), "#0")
        Else
            Text1(ptxAFT_KOUTEI_TANI02).Text = "0"
        End If
        Text1(ptxAFT_KOUTEI_QTY02).Text = "1"
        Text1(ptxAFT_KOUTEI_KOUSU02).Text = Format(CDbl(Text1(ptxAFT_KOUTEI_TANI02).Text) * CInt(Text1(ptxAFT_KOUTEI_QTY02).Text), "#0")
    End If
    '�B
    If IsNumeric(StrConv(P_KANRIREC02.AFT_KOTEI(5).KOTEI, vbUnicode)) Then
        Text1(ptxAFT_KOUTEI_TANI03).Text = Format(CInt(StrConv(P_KANRIREC02.AFT_KOTEI(5).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxAFT_KOUTEI_TANI03).Text = "0"
    End If
    Text1(ptxAFT_KOUTEI_QTY03).Text = "1"
    Text1(ptxAFT_KOUTEI_KOUSU03).Text = Format(CInt(Text1(ptxAFT_KOUTEI_TANI03).Text) * CInt(Text1(ptxAFT_KOUTEI_QTY03).Text), "#0")
    '計
    wkInt = 0
    For i = ptxAFT_KOUTEI_KOUSU01 To ptxAFT_KOUTEI_KOUSU09 Step 3
    
        If IsNumeric(Text1(i).Text) Then
            wkInt = wkInt + CInt(Text1(i).Text)
        End If
    Next i
    Text1(ptxAFT_KOUTEI_KEI1).Text = Format(wkInt, "#0")
    
    Text1(ptxAFT_KOUTEI_R_RATE).Text = ToHalfAdjust(CCur(wkInt * CDbl(YOYU_RITU(0).Caption)), 0)
    
    '(秒／個)
    If IsNumeric(Text1(ptxBEF_SEI_LOT).Text) Then
        
        If CInt(Text1(ptxBEF_SEI_LOT).Text) = 0 Then
            Text1(ptxAFT_KOUTEI_KEI2).Text = "0"
        Else
            Text1(ptxAFT_KOUTEI_KEI2).Text = Format(ToHalfAdjust(CCur(CInt(Text1(ptxAFT_KOUTEI_R_RATE).Text) / CInt(Text1(ptxBEF_SEI_LOT).Text)), 0), "#0")
        End If
    Else
        Text1(ptxAFT_KOUTEI_KEI2).Text = "0"
    End If
    '(分／個)
    Text1(ptxAFT_KOUTEI_KEI3).Text = Format(ToRoundUp(CCur(CInt(Text1(ptxAFT_KOUTEI_KEI2).Text) / 60), 1), "#0.0")
    '(円／個)
    If IsNumeric(Text1(ptxBEF_SEI_RATE).Text) Then
        Text1(ptxAFT_KOUTEI_KEI4).Text = Format(ToRoundUp(CCur(CDbl(Text1(ptxAFT_KOUTEI_KEI3).Text) * CInt((Text1(ptxBEF_SEI_RATE).Text))), 1), "#0.00")
    Else
        Text1(ptxAFT_KOUTEI_KEI4).Text = "0.00"
    End If
    
    
    '工程計
    Text1(ptxKOUTEI_KEI1).Text = Format(CInt(Text1(ptxBEF_KOUTEI_KEI1).Text) + CInt(Text1(ptxMAIN_KOUTEI_KEI1).Text) + CInt(Text1(ptxAFT_KOUTEI_KEI1).Text), "#0")
    
    Text1(ptxKOUTEI_R_RATE).Text = Format(CInt(Text1(ptxBEF_KOUTEI_R_RATE).Text) + CInt(Text1(ptxMAIN_KOUTEI_R_RATE).Text) + CInt(Text1(ptxAFT_KOUTEI_R_RATE).Text), "#0")
    
    Text1(ptxKOUTEI_KEI2).Text = Format(CInt(Text1(ptxBEF_KOUTEI_KEI2).Text) + CInt(Text1(ptxMAIN_KOUTEI_KEI2).Text) + CInt(Text1(ptxAFT_KOUTEI_KEI2).Text), "#0")
    
'''    Text1(ptxKOUTEI_KEI3).Text = Format(CDbl(Text1(ptxBEF_KOUTEI_KEI3).Text) + CDbl(Text1(ptxMAIN_KOUTEI_KEI3).Text) + CDbl(Text1(ptxAFT_KOUTEI_KEI3).Text), "#0.0")
'''    Text1(ptxKOUTEI_KEI4).Text = Format(CDbl(Text1(ptxBEF_KOUTEI_KEI4).Text) + CDbl(Text1(ptxMAIN_KOUTEI_KEI4).Text) + CDbl(Text1(ptxAFT_KOUTEI_KEI4).Text), "#0.00")
    
    
    
    '(分／個)
    Text1(ptxKOUTEI_KEI3).Text = Format(ToRoundUp(CCur(CInt(Text1(ptxKOUTEI_KEI2).Text) / 60), 1), "#0.0")
    '(円／個)
    If IsNumeric(Text1(ptxBEF_SEI_RATE).Text) Then
        Text1(ptxKOUTEI_KEI4).Text = Format(ToRoundUp(CCur(CDbl(Text1(ptxKOUTEI_KEI3).Text) * CInt((Text1(ptxBEF_SEI_RATE).Text))), 1), "#0.00")
    Else
        Text1(ptxKOUTEI_KEI4).Text = "0.00"
    End If
    
    
    
    
    '-----------------------------------    変更前／変更後（集計値）
    
    
'    '工数
    Text1(ptxAFT_S_KOUSU).Text = Text1(ptxKOUTEI_KEI3).Text
'    '工料
    Text1(ptxAFT_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxKOUTEI_KEI4).Text), "#0.00")
'
'    '箱代
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    If IsNumeric(KOUSEI(i, ColG_ST_SHIKIN)) Then
                        wkDouble = wkDouble + CDbl(KOUSEI(i, ColG_ST_SHIKIN))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
    End If
    Text1(ptxAFT_S_SHIZAI_GENKA).Text = Format(wkDouble, "#0.00")

    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    If IsNumeric(KOUSEI(i, ColG_ST_URIKIN)) Then
                        wkDouble = wkDouble + CDbl(KOUSEI(i, ColG_ST_URIKIN))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
    End If
    Text1(ptxAFT_S_SHIZAI_BAIKA).Text = Format(wkDouble, "#0.00")

    'PPSC加工単価
'    If IsNumeric(StrConv(ITEMREC.S_PPSC_KAKO_KOSU, vbUnicode)) Then
'        Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text = Format(CDbl(StrConv(ITEMREC.S_PPSC_KAKO_KOSU, vbUnicode)), "#0.00")
'    Else
'        Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text = "0.00"
'    End If
    'BU加工単価
'    If IsNumeric(StrConv(ITEMREC.S_BU_KAKO_KOSU, vbUnicode)) Then
'        Text1(ptxBEF_S_BU_KAKO_KOSU).Text = Format(CDbl(StrConv(ITEMREC.S_BU_KAKO_KOSU, vbUnicode)), "#0.00")
'    Else
'        Text1(ptxBEF_S_BU_KAKO_KOSU).Text = "0.00"
'    End If



    '外装箱代
    wkDouble = 0
    If KUSATU_F Then
        If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
        Else
            For i = 1 To KOUSEI.UpperBound(1)
        
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = GAISO_KBN Then
            
            
                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
            
                        wkDouble = wkDouble + CDbl(KOUSEI(i, ColG_ST_URIKIN_KUSATU))
                    End If
            
                End If
        
        
            Next i
        End If
    End If
    Text1(ptxAFT_S_GAISO_TANKA).Text = Format(wkDouble, "#0.00")




    'PPSC原価
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(PPSC_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = PPSC_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkDouble = wkDouble + CDbl(ToRoundUp(CCur(CDbl(KOUSEI(i, ColS_KOUSU)) / 60 * CInt(Text1(ptxAFT_SEI_RATE).Text)), 2))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
        Text1(ptxAFT_S_PPSC_KAKO_KOSU).Text = Format(wkDouble, "#0.00")
    End If



    'BU原価
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(BU_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = BU_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkDouble = wkDouble + CDbl(ToRoundUp(CCur(CDbl(KOUSEI(i, ColS_KOUSU)) / 60 * CInt(Text1(ptxAFT_SEI_RATE).Text)), 2))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
    
        Text1(ptxAFT_S_BU_KAKO_KOSU).Text = Format(wkDouble, "#0.00")
    
    End If

    DoEvents

    Detail_Disp_Proc = False

End Function

Private Function MONTHLYQTY_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   月平均出荷数画面表示
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim Total       As Long

Dim S_YM        As String * 6
Dim E_YM        As String * 6
Dim GET_YM      As String * 6


Dim NOW_YM      As String * 6

Dim cVer1       As String
Dim cVer2       As String

Dim cHEX        As String

Dim i           As Integer
Dim j           As Integer
Dim k           As Integer
Dim MONTH_Cnt   As Integer
Dim MONTH_QTY   As Long


    MONTHLYQTY_Disp_Proc = True
    
    
    
    NOW_YM = Left(Format(Now, "YYYYMMDD"), 6)
    
    
    '前年度対象年月
    If Right(NOW_YM, 2) < "04" Then
        S_YM = Format(CInt(Left(NOW_YM, 4) - 2), "0000") & "04"
    Else
        S_YM = Format(CInt(Left(NOW_YM, 4) - 1), "0000") & "04"
    End If
    
    
    '月平均出荷数 (月別集計)読み込み＆集計
    Total = 0
    
    j = ptxZEN_SYUKAQTY04
    
    
    For i = 0 To 11
    
        DoEvents
    
            
            
    
        GET_YM = Left(S_YM, 4) + Format(CInt(Right(S_YM, 2)) + i, "00")
        If Right(GET_YM, 2) > "12" Then
            GET_YM = Format(CInt(Left(GET_YM, 4)) + 1, "0000") & Format(CInt(Right(GET_YM, 2)) - 12, "00")
        End If
    
    
        Call UniCode_Conv(K0_MONTHLYQTY.DT, GET_YM)
        Call UniCode_Conv(K0_MONTHLYQTY.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_MONTHLYQTY.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
        Call UniCode_Conv(K0_MONTHLYQTY.HIN_GAI, Text1(ptxHin_Gai).Text)
        
        
    
        sts = BTRV(BtOpGetEqual, MONTHLYQTY_POS, MONTHLYQTYREC, Len(MONTHLYQTYREC), K0_MONTHLYQTY, Len(K0_MONTHLYQTY), 0)
        Select Case sts
            Case BtNoErr
            
            
                cVer1 = ""
                For k = 0 To UBound(MONTHLYQTYREC.SyukaQty)
                
                    cHEX = Hex(MONTHLYQTYREC.SyukaQty(k))
                    If Len(cHEX) < 2 Then
                        cHEX = "0" & cHEX
                    End If
                            
                    cVer1 = cVer1 & cHEX
                
                Next k
                MONTH_QTY = CLng(Left(cVer1, 9))
                    
                Text1(j).Text = Format(MONTH_QTY, "#,##0")
                Total = Total + MONTH_QTY
            
            
            
            Case BtErrKeyNotFound
                Text1(j).Text = "0"
            Case Else
                Call File_Error(sts, BtOpGetEqual, "月平均出荷数 (月別集計)")
                Exit Function
    
        End Select
        
    
        j = j + 1
    
    Next i
    
    
    Total = ToRoundUp(CCur(Total / 12), 0)
    Text1(ptxZEN_AVE).Text = Format(Total, "#,##0")
    
    
    
    
    
    
    
    '今年度対象年月
    If Right(NOW_YM, 2) < "04" Then
        S_YM = Format(CInt(Left(NOW_YM, 4) - 1), "0000") & "04"
    Else
        S_YM = Left(NOW_YM, 4) & "04"
    End If
    
    E_YM = Left(Format(DateAdd("m", -1, Left(Format(Now, "YYYY/MM/DD"), 7) & "/01"), "YYYYMMDD"), 6)
    
    
    
    
    
    '月平均出荷数 (月別集計)読み込み＆集計
    Total = 0
    MONTH_Cnt = 0
    j = ptxTOU_SYUKAQTY04
    
    
    For i = 0 To 11
    
        DoEvents
    
            
            
    
        GET_YM = Left(S_YM, 4) + Format(CInt(Right(S_YM, 2)) + i, "00")
        If Right(GET_YM, 2) > "12" Then
            GET_YM = Format(CInt(Left(GET_YM, 4)) + 1, "0000") & Format(CInt(Right(GET_YM, 2)) - 12, "00")
        End If
    
        If GET_YM > E_YM Then
            Exit For
        End If
    
        Call UniCode_Conv(K0_MONTHLYQTY.DT, GET_YM)
        Call UniCode_Conv(K0_MONTHLYQTY.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_MONTHLYQTY.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
        Call UniCode_Conv(K0_MONTHLYQTY.HIN_GAI, Text1(ptxHin_Gai).Text)
        
        
    
        sts = BTRV(BtOpGetEqual, MONTHLYQTY_POS, MONTHLYQTYREC, Len(MONTHLYQTYREC), K0_MONTHLYQTY, Len(K0_MONTHLYQTY), 0)
        Select Case sts
            Case BtNoErr
            
            
                cVer1 = ""
                For k = 0 To UBound(MONTHLYQTYREC.SyukaQty)
                
                    cHEX = Hex(MONTHLYQTYREC.SyukaQty(k))
                    If Len(cHEX) < 2 Then
                        cHEX = "0" & cHEX
                    End If
                            
                    cVer1 = cVer1 & cHEX
                
                Next k
                MONTH_QTY = CLng(Left(cVer1, 9))
                    
                Text1(j).Text = Format(MONTH_QTY, "#,##0")
                Total = Total + MONTH_QTY
            
            
            
            Case BtErrKeyNotFound
                Text1(j).Text = "0"
            Case Else
                Call File_Error(sts, BtOpGetEqual, "月平均出荷数 (月別集計)")
                Exit Function
    
        End Select
        
        MONTH_Cnt = MONTH_Cnt + 1
    
        j = j + 1
    
    Next i
    
    
    If MONTH_Cnt = 0 Then
        Total = 0
    Else
        Total = ToRoundUp(CCur(Total / MONTH_Cnt), 0)
        
    End If
    Text1(ptxTOU_AVE).Text = Format(Total, "#,##0")
    
    
    
    
    
    
    
    
    MONTHLYQTY_Disp_Proc = False

End Function
Private Function TANKA_KEISAN_Proc() As Integer
'----------------------------------------------------------------------------
'                   単価計算処理
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer
Dim wkInt       As Integer
Dim wkDouble    As Double


Dim c           As String * 128
Dim wkKUSATU    As Variant
Dim INV_F       As Boolean


    TANKA_KEISAN_Proc = True
    
    '品目マスタ読み込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound

            Text1(ptxHin_Name).Text = ""
            Text1(ptxST_SOKO).Text = ""
            Text1(ptxST_RETU).Text = ""
            Text1(ptxST_REN).Text = ""
            Text1(ptxST_DAN).Text = ""
            MsgBox "入力した項目はエラーです。(品番)"
            TANKA_KEISAN_Proc = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function

    End Select


    '設定日
    Text1(ptxAFT_S_KOUSU_SET_DATE).Text = Format(Now, "YYYYMMDD")
    '担当者
    Text1(ptxAFT_SEI_TANKA_TANTO).Text = Text1(ptxTanto_Code).Text
    
    
    '-----------------------------------    前工程
    
    '�@
    If IsNumeric(StrConv(P_KANRIREC02.BEF_KOTEI(0).KOTEI, vbUnicode)) Then
        Text1(ptxBEF_KOUTEI_TANI01).Text = Format(CInt(StrConv(P_KANRIREC02.BEF_KOTEI(0).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxBEF_KOUTEI_TANI01).Text = "0"
    End If
    Text1(ptxBEF_KOUTEI_QTY01).Text = "1"
    Text1(ptxBEF_KOUTEI_KOUSU01).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI01).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY01).Text), "#0")
    '�A
    wkInt = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    wkInt = wkInt + 1
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    If IsNumeric(StrConv(P_KANRIREC02.BEF_KOTEI(1).KOTEI, vbUnicode)) Then
        Text1(ptxBEF_KOUTEI_TANI02).Text = Format(CInt(StrConv(P_KANRIREC02.BEF_KOTEI(1).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxBEF_KOUTEI_TANI02).Text = "0"
    End If
    Text1(ptxBEF_KOUTEI_QTY02).Text = Format(wkInt, "#0")
    Text1(ptxBEF_KOUTEI_KOUSU02).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI02).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY02).Text), "#0")
    '�B
    wkInt = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(DOUKON_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = DOUKON_T(j) Then
                    wkInt = wkInt + 1
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    If IsNumeric(StrConv(P_KANRIREC02.BEF_KOTEI(2).KOTEI, vbUnicode)) Then
        Text1(ptxBEF_KOUTEI_TANI03).Text = Format(CInt(StrConv(P_KANRIREC02.BEF_KOTEI(2).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxBEF_KOUTEI_TANI03).Text = "0"
    End If
    Text1(ptxBEF_KOUTEI_QTY03).Text = Format(wkInt, "#0")
    Text1(ptxBEF_KOUTEI_KOUSU03).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI03).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY03).Text), "#0")
    '�C
    If KUSATU_F Then
        '草津はINI参照
            
        wkInt = 0
        If GetIni(Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1), KUSATU_ETC, App.EXEName, c) Then
        Else
            wkKUSATU = Split(Trim(c), ",", -1)
                    
            If IsNumeric(wkKUSATU(0)) Then
                wkInt = CInt(wkKUSATU(0))
            End If
        End If
        If GetIni(Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1), Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), App.EXEName, c) Then
            If GetIni(App.EXEName, Trim(StrConv(StrConv(ITEMREC.HIN_NAME, vbUnicode), vbWide)), App.EXEName, c) Then
            Else
                wkKUSATU = Split(Trim(c), ",", -1)
                        
                If IsNumeric(wkKUSATU(0)) Then
                    wkInt = CInt(wkKUSATU(0))
                End If
            End If
        Else
            wkKUSATU = Split(Trim(c), ",", -1)
            If IsNumeric(wkKUSATU(0)) Then
                wkInt = CInt(wkKUSATU(0))
            End If
        End If
        Text1(ptxBEF_KOUTEI_TANI04).Text = Format(wkInt, "#0")
        Text1(ptxBEF_KOUTEI_QTY04).Text = "1"
        Text1(ptxBEF_KOUTEI_KOUSU04).Text = Format(CDbl(Text1(ptxBEF_KOUTEI_TANI04).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY04).Text), "#0")
    
    
    Else
        INV_F = False
        Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
            
                Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, StrConv(SOKOREC.IO_TANKA_No, vbUnicode))
                sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    
                        INV_F = True
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                        Exit Function
                End Select
            
            Case BtErrKeyNotFound
            
                INV_F = True
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                Exit Function
        
        End Select
        
        
        If INV_F Then
            
            Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, INV_IO_TANKA_No)
            sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                
                    Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_S_OUT_KOUSU, "000.00")
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                    Exit Function
            End Select
        End If
        If IsNumeric(StrConv(SE_LOC_TANKA_M_REC.SE_S_OUT_KOUSU, vbUnicode)) Then
            Text1(ptxBEF_KOUTEI_TANI04).Text = Format(CInt(StrConv(SE_LOC_TANKA_M_REC.SE_S_OUT_KOUSU, vbUnicode)), "#0")
        Else
            Text1(ptxBEF_KOUTEI_TANI04).Text = "0"
        End If
        Text1(ptxBEF_KOUTEI_QTY04).Text = "1"
        Text1(ptxBEF_KOUTEI_KOUSU04).Text = Format(CDbl(Text1(ptxBEF_KOUTEI_TANI04).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY04).Text), "#0")
    End If
    '�D
    wkInt = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    wkInt = wkInt + 1
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    If IsNumeric(StrConv(P_KANRIREC02.BEF_KOTEI(4).KOTEI, vbUnicode)) Then
        Text1(ptxBEF_KOUTEI_TANI05).Text = Format(CInt(StrConv(P_KANRIREC02.BEF_KOTEI(4).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxBEF_KOUTEI_TANI05).Text = "0"
    End If
    Text1(ptxBEF_KOUTEI_QTY05).Text = Format(wkInt, "#0")
    Text1(ptxBEF_KOUTEI_KOUSU05).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI05).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY05).Text), "#0")
    '�E
    wkInt = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(DOUKON_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = DOUKON_T(j) Then
                    wkInt = wkInt + 1
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    If IsNumeric(StrConv(P_KANRIREC02.BEF_KOTEI(5).KOTEI, vbUnicode)) Then
        Text1(ptxBEF_KOUTEI_TANI06).Text = Format(CInt(StrConv(P_KANRIREC02.BEF_KOTEI(5).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxBEF_KOUTEI_TANI06).Text = "0"
    End If
    Text1(ptxBEF_KOUTEI_QTY06).Text = Format(wkInt, "#0")
    Text1(ptxBEF_KOUTEI_KOUSU06).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI06).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY06).Text), "#0")
    
    '�F
    If IsNumeric(StrConv(P_KANRIREC02.BEF_KOTEI(6).KOTEI, vbUnicode)) Then
        Text1(ptxBEF_KOUTEI_TANI07).Text = Format(CInt(StrConv(P_KANRIREC02.BEF_KOTEI(6).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxBEF_KOUTEI_TANI07).Text = "0"
    End If
    Text1(ptxBEF_KOUTEI_QTY07).Text = "1"
    Text1(ptxBEF_KOUTEI_KOUSU07).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI07).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY07).Text), "#0")
    '�G
    If IsNumeric(StrConv(P_KANRIREC02.BEF_KOTEI(7).KOTEI, vbUnicode)) Then
        Text1(ptxBEF_KOUTEI_TANI08).Text = Format(CInt(StrConv(P_KANRIREC02.BEF_KOTEI(7).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxBEF_KOUTEI_TANI08).Text = "0"
    End If
    Text1(ptxBEF_KOUTEI_QTY08).Text = "1"
    Text1(ptxBEF_KOUTEI_KOUSU08).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI08).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY08).Text), "#0")
    '�H
    If IsNumeric(StrConv(P_KANRIREC02.BEF_KOTEI(8).KOTEI, vbUnicode)) Then
        Text1(ptxBEF_KOUTEI_TANI09).Text = Format(CInt(StrConv(P_KANRIREC02.BEF_KOTEI(8).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxBEF_KOUTEI_TANI09).Text = "0"
    End If
    Text1(ptxBEF_KOUTEI_QTY09).Text = "1"
    Text1(ptxBEF_KOUTEI_KOUSU09).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI09).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY09).Text), "#0")
    '計
    wkInt = 0
    For i = ptxBEF_KOUTEI_KOUSU01 To ptxBEF_KOUTEI_KOUSU09 Step 3
    
        wkInt = wkInt + CInt(Text1(i).Text)
    
    Next i
    Text1(ptxBEF_KOUTEI_KEI1).Text = Format(wkInt, "#0")
    
    Text1(ptxBEF_KOUTEI_R_RATE).Text = ToHalfAdjust(CCur(wkInt * CDbl(YOYU_RITU(0).Caption)), 0)
    
    
    '(秒／個)
    If IsNumeric(Text1(ptxBEF_SEI_LOT).Text) Then
        Text1(ptxBEF_KOUTEI_KEI2).Text = Format(ToHalfAdjust(CCur(CInt(Text1(ptxBEF_KOUTEI_R_RATE).Text) / CInt(Text1(ptxAFT_SEI_LOT).Text)), 0), "#0")
    Else
        Text1(ptxBEF_KOUTEI_KEI2).Text = "0"
    End If
    '(分／個)
    Text1(ptxBEF_KOUTEI_KEI3).Text = Format(ToRoundUp(CCur(CInt(Text1(ptxBEF_KOUTEI_KEI2).Text) / 60), 1), "#0.0")
    '(円／個)
    If IsNumeric(Text1(ptxBEF_SEI_RATE).Text) Then
        Text1(ptxBEF_KOUTEI_KEI4).Text = Format(ToRoundUp(CCur(CDbl(Text1(ptxBEF_KOUTEI_KEI3).Text) * CInt((Text1(ptxAFT_SEI_RATE).Text))), 2), "#0.00")
    Else
        Text1(ptxBEF_KOUTEI_KEI4).Text = "0.00"
    End If
        
    
    
    '-----------------------------------    作業工程
    '�@
    If IsNumeric(StrConv(P_KANRIREC02.MAIN_KOTEI(0).KOTEI, vbUnicode)) Then
        Text1(ptxMAIN_KOUTEI_TANI01).Text = Format(CInt(StrConv(P_KANRIREC02.MAIN_KOTEI(0).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxMAIN_KOUTEI_TANI01).Text = "0"
    End If
'    Text1(ptxMAIN_KOUTEI_QTY01).Text = "1"
    If Trim(Text1(ptxMAIN_KOUTEI_QTY01).Text) = "" Then
        If IsNumeric(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode)) Then
                Text1(ptxMAIN_KOUTEI_QTY01).Text = Format(CInt(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode)), "#0")
        Else
                Text1(ptxMAIN_KOUTEI_QTY01).Text = "1"
        End If
    End If
    Text1(ptxMAIN_KOUTEI_KOUSU01).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI01).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY01).Text), "#0")
    
    '�A
    wkInt = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkInt = wkInt + CInt(ToRoundUp(CCur(KOUSEI(i, ColS_KOUSU)) * CDbl(KOUSEI(i, ColKO_QTY)), 0))
                    End If
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI02).Text = Format(wkInt, "#0")
    Text1(ptxMAIN_KOUTEI_QTY02).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU02).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI02).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY02).Text), "#0")
    '�B
    wkInt = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(DOUKON_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = DOUKON_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColKO_QTY)) Then
                        wkInt = wkInt + CInt(ToRoundUp(CCur(KOUSEI(i, ColKO_QTY)), 0))
                    End If
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    If IsNumeric(StrConv(P_KANRIREC02.MAIN_KOTEI(3).KOTEI, vbUnicode)) Then
        Text1(ptxMAIN_KOUTEI_TANI03).Text = Format(CInt(StrConv(P_KANRIREC02.MAIN_KOTEI(3).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxMAIN_KOUTEI_TANI03).Text = "0"
    End If
    Text1(ptxMAIN_KOUTEI_QTY03).Text = Format(wkInt, "#0")
    Text1(ptxMAIN_KOUTEI_KOUSU03).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI03).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY03).Text), "#0")
    '�C
    wkInt = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(KAKOU_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = KAKOU_T(j) Then
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkInt = wkInt + CInt(KOUSEI(i, ColS_KOUSU))
                    End If
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI04).Text = Format(wkInt, "#0")
    Text1(ptxMAIN_KOUTEI_QTY04).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU04).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI04).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY04).Text), "#0")
    '�D
    wkInt = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
            
            
            For j = 0 To UBound(SHIZAI_T)
            
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    If IsNumeric(KOUSEI(i, ColSEI_SYU_KON)) Then
                        wkInt = wkInt + CInt(KOUSEI(i, ColSEI_SYU_KON))
                    End If
                End If
            
            Next j
            
            
            
            
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI05).Text = Format(wkInt, "#0")
    Text1(ptxMAIN_KOUTEI_QTY05).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU05).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI05).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY05).Text), "#0")
    '計
    wkInt = 0
    For i = ptxMAIN_KOUTEI_KOUSU01 To ptxMAIN_KOUTEI_KOUSU09 Step 3
    
        If IsNumeric(Text1(i).Text) Then
            wkInt = wkInt + CInt(Text1(i).Text)
        End If
    Next i
    Text1(ptxMAIN_KOUTEI_KEI1).Text = Format(wkInt, "#0")
    
    Text1(ptxMAIN_KOUTEI_R_RATE).Text = ToHalfAdjust(CCur(wkInt * CDbl(YOYU_RITU(0).Caption)), 0)
    '(秒／個)
    If IsNumeric(Text1(ptxBEF_SEI_LOT).Text) Then
        Text1(ptxMAIN_KOUTEI_KEI2).Text = Format(ToHalfAdjust(CCur(Text1(ptxMAIN_KOUTEI_R_RATE).Text), 0), "#0")
    Else
        Text1(ptxMAIN_KOUTEI_KEI2).Text = "0"
    End If
    '(分／個)
    Text1(ptxMAIN_KOUTEI_KEI3).Text = Format(ToRoundUp(CCur(CInt(Text1(ptxMAIN_KOUTEI_KEI2).Text) / 60), 1), "#0.0")
    '(円／個)
    If IsNumeric(Text1(ptxBEF_SEI_RATE).Text) Then
        Text1(ptxMAIN_KOUTEI_KEI4).Text = Format(ToRoundUp(CCur(CDbl(Text1(ptxMAIN_KOUTEI_KEI3).Text) * CInt((Text1(ptxAFT_SEI_RATE).Text))), 2), "#0.00")
    Else
        Text1(ptxMAIN_KOUTEI_KEI4).Text = "0.00"
    End If
    
    '-----------------------------------    後工程
    '�@
    If IsNumeric(StrConv(P_KANRIREC02.AFT_KOTEI(0).KOTEI, vbUnicode)) Then
        Text1(ptxAFT_KOUTEI_TANI01).Text = Format(CInt(StrConv(P_KANRIREC02.AFT_KOTEI(0).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxAFT_KOUTEI_TANI01).Text = "0"
    End If
    Text1(ptxAFT_KOUTEI_QTY01).Text = "1"
    Text1(ptxAFT_KOUTEI_KOUSU01).Text = Format(CInt(Text1(ptxAFT_KOUTEI_TANI01).Text) * CInt(Text1(ptxAFT_KOUTEI_QTY01).Text), "#0")
    '�A
    If KUSATU_F Then
        '草津はINIより
    
        wkInt = 0
        If GetIni(Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1), KUSATU_ETC, App.EXEName, c) Then
        Else
            wkKUSATU = Split(Trim(c), ",", -1)
                    
            If IsNumeric(wkKUSATU(1)) Then
                wkInt = CInt(wkKUSATU(1))
            End If
        End If
        If GetIni(Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1), Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), App.EXEName, c) Then
            If GetIni(App.EXEName, Trim(StrConv(StrConv(ITEMREC.HIN_NAME, vbUnicode), vbWide)), App.EXEName, c) Then
            Else
                wkKUSATU = Split(Trim(c), ",", -1)
                        
                If IsNumeric(wkKUSATU(0)) Then
                    wkInt = CInt(wkKUSATU(1))
                End If
            End If
        Else
            wkKUSATU = Split(Trim(c), ",", -1)
                    
            If IsNumeric(wkKUSATU(1)) Then
                wkInt = CInt(wkKUSATU(1))
            End If
        End If
    
        Text1(ptxAFT_KOUTEI_TANI02).Text = Format(wkInt, "#0")
    
        Text1(ptxAFT_KOUTEI_QTY02).Text = "1"
        Text1(ptxAFT_KOUTEI_KOUSU02).Text = Format(CDbl(Text1(ptxAFT_KOUTEI_TANI02).Text) * CInt(Text1(ptxAFT_KOUTEI_QTY02).Text), "#0")
    Else
        INV_F = False
        Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
            
                Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, StrConv(SOKOREC.IO_TANKA_No, vbUnicode))
                sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    
                        INV_F = True
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                        Exit Function
                End Select
            
            Case BtErrKeyNotFound
            
                INV_F = True
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                Exit Function
        
        End Select
        
        
        If INV_F Then
            
            Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, INV_IO_TANKA_No)
            sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                
                    Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_S_IN_KOUSU, "000.00")
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                    Exit Function
            End Select
        End If
        If IsNumeric(StrConv(SE_LOC_TANKA_M_REC.SE_S_IN_KOUSU, vbUnicode)) Then
            Text1(ptxAFT_KOUTEI_TANI02).Text = Format(CInt(StrConv(SE_LOC_TANKA_M_REC.SE_S_IN_KOUSU, vbUnicode)), "#0")
        Else
            Text1(ptxAFT_KOUTEI_TANI02).Text = "0"
        End If
        Text1(ptxAFT_KOUTEI_QTY02).Text = "1"
        Text1(ptxAFT_KOUTEI_KOUSU02).Text = Format(CDbl(Text1(ptxAFT_KOUTEI_TANI02).Text) * CInt(Text1(ptxAFT_KOUTEI_QTY02).Text), "#0")
    End If
    '�B
    If IsNumeric(StrConv(P_KANRIREC02.AFT_KOTEI(5).KOTEI, vbUnicode)) Then
        Text1(ptxAFT_KOUTEI_TANI03).Text = Format(CInt(StrConv(P_KANRIREC02.AFT_KOTEI(5).KOTEI, vbUnicode)), "#0")
    Else
        Text1(ptxAFT_KOUTEI_TANI03).Text = "0"
    End If
    Text1(ptxAFT_KOUTEI_QTY03).Text = "1"
    Text1(ptxAFT_KOUTEI_KOUSU03).Text = Format(CInt(Text1(ptxAFT_KOUTEI_TANI03).Text) * CInt(Text1(ptxAFT_KOUTEI_QTY03).Text), "#0")
    '計
    wkInt = 0
    For i = ptxAFT_KOUTEI_KOUSU01 To ptxAFT_KOUTEI_KOUSU09 Step 3
    
        If IsNumeric(Text1(i).Text) Then
            wkInt = wkInt + CInt(Text1(i).Text)
        End If
    Next i
    Text1(ptxAFT_KOUTEI_KEI1).Text = Format(wkInt, "#0")
    Text1(ptxAFT_KOUTEI_R_RATE).Text = ToHalfAdjust(CCur(wkInt * CDbl(YOYU_RITU(0).Caption)), 0)
    '(秒／個)
    If IsNumeric(Text1(ptxBEF_SEI_LOT).Text) Then
        Text1(ptxAFT_KOUTEI_KEI2).Text = Format(ToHalfAdjust(CCur(CInt(Text1(ptxAFT_KOUTEI_R_RATE).Text) / CInt(Text1(ptxAFT_SEI_LOT).Text)), 0), "#0")
    Else
        Text1(ptxAFT_KOUTEI_KEI2).Text = "0"
    End If
    '(分／個)
    Text1(ptxAFT_KOUTEI_KEI3).Text = Format(ToRoundUp(CCur(CInt(Text1(ptxAFT_KOUTEI_KEI2).Text) / 60), 1), "#0.0")
    '(円／個)
    If IsNumeric(Text1(ptxBEF_SEI_RATE).Text) Then
        Text1(ptxAFT_KOUTEI_KEI4).Text = Format(ToRoundUp(CCur(CDbl(Text1(ptxAFT_KOUTEI_KEI3).Text) * CInt((Text1(ptxAFT_SEI_RATE).Text))), 2), "#0.00")
    Else
        Text1(ptxAFT_KOUTEI_KEI4).Text = "0.00"
    End If
    
    
    '工程計
    Text1(ptxKOUTEI_KEI1).Text = Format(CInt(Text1(ptxBEF_KOUTEI_KEI1).Text) + CInt(Text1(ptxMAIN_KOUTEI_KEI1).Text) + CInt(Text1(ptxAFT_KOUTEI_KEI1).Text), "#0")
    
    Text1(ptxKOUTEI_R_RATE).Text = Format(CInt(Text1(ptxBEF_KOUTEI_R_RATE).Text) + CInt(Text1(ptxMAIN_KOUTEI_R_RATE).Text) + CInt(Text1(ptxAFT_KOUTEI_R_RATE).Text), "#0")
    
    Text1(ptxKOUTEI_KEI2).Text = Format(CInt(Text1(ptxBEF_KOUTEI_KEI2).Text) + CInt(Text1(ptxMAIN_KOUTEI_KEI2).Text) + CInt(Text1(ptxAFT_KOUTEI_KEI2).Text), "#0")
'''    Text1(ptxKOUTEI_KEI3).Text = Format(CDbl(Text1(ptxBEF_KOUTEI_KEI3).Text) + CDbl(Text1(ptxMAIN_KOUTEI_KEI3).Text) + CDbl(Text1(ptxAFT_KOUTEI_KEI3).Text), "#0.0")
'''    Text1(ptxKOUTEI_KEI4).Text = Format(CDbl(Text1(ptxBEF_KOUTEI_KEI4).Text) + CDbl(Text1(ptxMAIN_KOUTEI_KEI4).Text) + CDbl(Text1(ptxAFT_KOUTEI_KEI4).Text), "#0.00")
    
    
    
    '(分／個)
    Text1(ptxKOUTEI_KEI3).Text = Format(ToRoundUp(CCur(CInt(Text1(ptxKOUTEI_KEI2).Text) / 60), 1), "#0.0")
    '(円／個)
    If IsNumeric(Text1(ptxBEF_SEI_RATE).Text) Then
        Text1(ptxKOUTEI_KEI4).Text = Format(ToRoundUp(CCur(CDbl(Text1(ptxKOUTEI_KEI3).Text) * CInt((Text1(ptxBEF_SEI_RATE).Text))), 1), "#0.00")
    Else
        Text1(ptxKOUTEI_KEI4).Text = "0.00"
    End If
    
    '-----------------------------------    変更後
    
    
'    '工数
    Text1(ptxAFT_S_KOUSU).Text = Text1(ptxKOUTEI_KEI3).Text
'    '工料
    Text1(ptxAFT_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxKOUTEI_KEI4).Text), "#0.00")
'
'    '箱代
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    If IsNumeric(KOUSEI(i, ColG_ST_SHIKIN)) Then
                        wkDouble = wkDouble + CDbl(KOUSEI(i, ColG_ST_SHIKIN))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
    End If
    Text1(ptxAFT_S_SHIZAI_GENKA).Text = Format(wkDouble, "#0.00")
'
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColG_ST_URIKIN)) Then
                        wkDouble = wkDouble + CDbl(KOUSEI(i, ColG_ST_URIKIN))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
    End If
    Text1(ptxAFT_S_SHIZAI_BAIKA).Text = Format(wkDouble, "#0.00")



    '外装箱代
    wkDouble = 0
    If KUSATU_F Then
        If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
        Else
            For i = 1 To KOUSEI.UpperBound(1)
        
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = GAISO_KBN Then
            
            
                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
            
                        wkDouble = wkDouble + CDbl(KOUSEI(i, ColG_ST_URIKIN_KUSATU))
                    End If
            
                End If
        
        
            Next i
        End If
    End If
    Text1(ptxAFT_S_GAISO_TANKA).Text = Format(wkDouble, "#0.00")




    'PPSC原価
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(PPSC_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = PPSC_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkDouble = wkDouble + CDbl(ToRoundUp(CCur(CDbl(KOUSEI(i, ColS_KOUSU)) / 60 * CInt(Text1(ptxAFT_SEI_RATE).Text)), 2))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
        Text1(ptxAFT_S_PPSC_KAKO_KOSU).Text = Format(wkDouble, "#0.00")
    End If



    'BU原価
    wkDouble = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
    
            For j = 0 To UBound(BU_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = BU_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkDouble = wkDouble + CDbl(ToRoundUp(CCur(CDbl(KOUSEI(i, ColS_KOUSU)) / 60 * CInt(Text1(ptxAFT_SEI_RATE).Text)), 2))
                    End If
                    Exit For
                End If
    
            Next j
    
        Next i
    
        Text1(ptxAFT_S_BU_KAKO_KOSU).Text = Format(wkDouble, "#0.00")
    
    End If



    TANKA_KEISAN_Proc = False

End Function


Private Function Tanka_Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   単価登録処理
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim ans         As Integer

Dim wkGAISO     As Double
    
Dim i           As Integer
Dim j            As Integer
    
    
Dim wkInt       As Integer
    
    Tanka_Update_Proc = True

    '品目マスタ読み込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)


    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
            
                MsgBox "他端末でデータが、変更されています。単価登録処理を中止します。"
                Tanka_Update_Proc = False
                Exit Function
            
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Tanka_Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                Exit Function
        
        End Select
    
    Loop

    
    '2008.09.20 追加↓
    
    
    
    '前作業
    i = ptxBEF_KOUTEI_KOUSU01
    
    
    For j = 0 To 9
    
        If IsNumeric(Text1(i).Text) Then
            Call UniCode_Conv(ITEMREC.BEF_KOUTEI(j).BEF_KOUTEI, Format(CDbl(Text1(i).Text), "000.00"))
        Else
            Call UniCode_Conv(ITEMREC.BEF_KOUTEI(j).BEF_KOUTEI, "000.00")
        End If
    
        i = i + 3
    
    
    
    Next j
    
    
    
    '主作業
    i = ptxMAIN_KOUTEI_KOUSU01
    
    
    For j = 0 To 9
    
        If IsNumeric(Text1(i).Text) Then
            Call UniCode_Conv(ITEMREC.MAIN_KOUTEI(j).MAIN_KOUTEI, Format(CDbl(Text1(i).Text), "000.00"))
        Else
            Call UniCode_Conv(ITEMREC.MAIN_KOUTEI(j).MAIN_KOUTEI, "000.00")
        End If
    
    
    
        i = i + 3
    
    
    Next j
    
    
    '後作業
    i = ptxAFT_KOUTEI_KOUSU01
    
    
    For j = 0 To 9
    
        If IsNumeric(Text1(i).Text) Then
            Call UniCode_Conv(ITEMREC.AFT_KOUTEI(j).AFT_KOUTEI, Format(CDbl(Text1(i).Text), "000.00"))
        Else
            Call UniCode_Conv(ITEMREC.AFT_KOUTEI(j).AFT_KOUTEI, "000.00")
        End If
    
    
    
        i = i + 3
    
    
    Next j
    
    
    
    '倉庫区分
    Call UniCode_Conv(ITEMREC.SE_IO_TANKA_No, StrConv(SOKOREC.IO_TANKA_No, vbUnicode))
    
    
    
    
    
    '2008.09.20 追加↑
    
    
    
    Do
        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Tanka_Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "品目マスタ")
                Exit Function
        
        End Select
    
    Loop
    
    
    '単価更新履歴出力
    Do
        sts = BTRV(BtOpInsert, ITEM_HST_POS, ITEMREC, Len(ITEMREC), K0_ITEM_HST, Len(K0_ITEM_HST), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM_HST.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Tanka_Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "品目単価更新履歴")
                Exit Function
        
        End Select
    
    Loop
    

    Tanka_Update_Proc = False


End Function

Private Function Grid_Error_Check_Proc() As Integer
'----------------------------------------------------------------------------
'                   ｸﾞﾘｯﾄﾞ内容のエラーチェック処理
'----------------------------------------------------------------------------
Dim i   As Integer

Dim sts As Integer
    
    
    Grid_Error_Check_Proc = True
    
    
    
    
    Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
    
'    TDBGrid1.Refresh
    
    TDBGrid1(pGrdKOUSEI).Update
    
    If KOUSEI.Count(1) < 1 Then
        Grid_Error_Check_Proc = False
        Exit Function
    End If


    For i = 1 To KOUSEI.Count(1)
    
    
        If Trim(KOUSEI(i, ColKO_HIN_GAI)) = "" Then
            
            KOUSEI(i, ColKO_JGYOBU) = ""
            KOUSEI(i, ColKO_NAIGAI) = ""
            
            KOUSEI(i, ColKO_HIN_NAME) = ""
            KOUSEI(i, ColKO_QTY) = ""
            KOUSEI(i, ColG_ST_SHITAN) = ""
            KOUSEI(i, ColG_ST_URITAN) = ""
            KOUSEI(i, ColG_ST_SHIKIN) = ""
            KOUSEI(i, ColG_ST_URIKIN) = ""
            KOUSEI(i, ColS_KOUSU) = ""
            KOUSEI(i, ColSEI_SYU_KON) = ""
    
            KOUSEI(i, ColKO_BIKOU) = ""
    
        Else
    
    
            '品番
            If Trim(KOUSEI(i, ColKO_JGYOBU)) = "" And _
                Trim(KOUSEI(i, ColKO_NAIGAI)) = "" Then
                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
            Else
                Call UniCode_Conv(K0_ITEM.JGYOBU, KOUSEI(i, ColKO_JGYOBU))
                Call UniCode_Conv(K0_ITEM.NAIGAI, KOUSEI(i, ColKO_NAIGAI))
            End If
            
            Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(i, ColKO_HIN_GAI))
        
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    '資材品で読み替え
                                    
                    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(i, ColKO_HIN_GAI))
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            
                            If HIN_INV Then
                                '未登録品番　可　資材としておく
                                Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                                Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                                Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番")
                                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                            Else
                                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(品番)"
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                            Exit Function
                    
                    End Select
            
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                    Exit Function
            
            End Select
    
            KOUSEI(i, ColKO_JGYOBU) = StrConv(ITEMREC.JGYOBU, vbUnicode)
            KOUSEI(i, ColKO_NAIGAI) = StrConv(ITEMREC.NAIGAI, vbUnicode)
            KOUSEI(i, ColKO_HIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    
    
            '員数
            If Trim(KOUSEI(i, ColKO_QTY)) = "" Then
                KOUSEI(i, ColKO_QTY) = "1.00"
            End If
            If IsNumeric(KOUSEI(i, ColKO_QTY)) Then
                KOUSEI(i, ColKO_QTY) = Format(CDbl(KOUSEI(i, ColKO_QTY)), "#0.00")
            Else
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(員数)"
    
            End If
    
    
            '仕入＠
            If Trim(KOUSEI(i, ColG_ST_SHITAN)) = "" Then
                KOUSEI(i, ColG_ST_SHITAN) = "0.00"
            End If
            If IsNumeric(KOUSEI(i, ColG_ST_SHITAN)) Then
                KOUSEI(i, ColG_ST_SHITAN) = Format(CDbl(KOUSEI(i, ColG_ST_SHITAN)), "#0.00")
            Else
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(仕入＠)"
    
            End If
            '販売＠
            If Trim(KOUSEI(i, ColG_ST_URITAN)) = "" Then
                KOUSEI(i, ColG_ST_URITAN) = "0.00"
            End If
            If IsNumeric(KOUSEI(i, ColG_ST_URITAN)) Then
                KOUSEI(i, ColG_ST_URITAN) = Format(CDbl(KOUSEI(i, ColG_ST_URITAN)), "#0.00")
            Else
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(販売＠)"
    
            End If
            
            '仕入金額計
            If Trim(KOUSEI(i, ColG_ST_SHIKIN)) = "" Then
                KOUSEI(i, ColG_ST_SHIKIN) = "0.00"
            End If
            If IsNumeric(KOUSEI(i, ColG_ST_SHIKIN)) Then
                KOUSEI(i, ColG_ST_SHIKIN) = Format(CDbl(KOUSEI(i, ColG_ST_SHIKIN)), "#0.00")
            Else
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(仕入金額計)"
    
            End If
            '販売金額計
            If Trim(KOUSEI(i, ColG_ST_URIKIN)) = "" Then
                KOUSEI(i, ColG_ST_URIKIN) = "0.00"
            End If
            If IsNumeric(KOUSEI(i, ColG_ST_URIKIN)) Then
                KOUSEI(i, ColG_ST_URIKIN) = Format(CDbl(KOUSEI(i, ColG_ST_URIKIN)), "#0.00")
            Else
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(販売金額計)"
    
            End If
            '作業時間
            If Trim(KOUSEI(i, ColS_KOUSU)) = "" Then
                KOUSEI(i, ColS_KOUSU) = "0"
            End If
            If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                KOUSEI(i, ColS_KOUSU) = Format(CDbl(KOUSEI(i, ColS_KOUSU)), "#0")
            Else
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(作業時間)"
            End If
            '集合梱包時間
            If Trim(KOUSEI(i, ColSEI_SYU_KON)) = "" Then
                KOUSEI(i, ColSEI_SYU_KON) = "0"
            End If
            If IsNumeric(KOUSEI(i, ColSEI_SYU_KON)) Then
                KOUSEI(i, ColSEI_SYU_KON) = Format(CDbl(KOUSEI(i, ColSEI_SYU_KON)), "#0")
            Else
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(集合梱包時間)"
            End If
    
        End If
    Next i

    Grid_Error_Check_Proc = False



End Function
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   構成マスタ出力
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim ans         As Integer
Dim com         As Integer

Dim K_SEQNO     As Integer
Dim G_SEQNO     As Integer
Dim D_SEQNO     As Integer


Dim i           As Integer
Dim j           As Integer

Dim MESG        As String


    Update_Proc = True
                                        
    Call Input_Lock
                                        
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
    
    
    
    '---------------------------------------------------    '単価更新
    MESG = "単価を登録します。よろしいですか？" & vbCrLf
    MESG = MESG & "　ロット数：" & Text1(ptxAFT_SEI_LOT).Text & vbCrLf
    MESG = MESG & "　分レート：" & Text1(ptxAFT_SEI_RATE).Text & vbCrLf
    MESG = MESG & "　工数：" & Text1(ptxAFT_S_KOUSU).Text & vbCrLf
    MESG = MESG & "　（原価）工料：" & Text1(ptxAFT_S_KOUSU_GENKA).Text & vbCrLf
    MESG = MESG & "　 (売価) 工料：" & Text1(ptxAFT_S_KOUSU_BAIKA).Text & vbCrLf
    MESG = MESG & "　（原価）工料：" & Text1(ptxAFT_S_SHIZAI_GENKA).Text & vbCrLf
    MESG = MESG & "　 (売価) 工料：" & Text1(ptxAFT_S_SHIZAI_BAIKA).Text & vbCrLf
    MESG = MESG & "　 設定日：" & Text1(ptxAFT_S_KOUSU_SET_DATE).Text & vbCrLf
    MESG = MESG & "　 担当者：" & Text1(ptxAFT_SEI_TANKA_TANTO).Text & vbCrLf
    MESG = MESG & "　 メモ：" & Text1(ptxAFT_SE_TANKA_MEMO).Text & vbCrLf
            
    ans = MsgBox(MESG, vbYesNo + vbDefaultButton1 + vbExclamation, "確認入力")
    If ans = vbYes Then
        If Tanka_Update_Proc() Then
            GoTo Abort_Tran
        End If
    
    End If
    
    
        
    '---------------------------------------------------    '構成マスタ更新
    '該当データ全件削除
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHin_Gai).Text)
       
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
                        Trim(StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHin_Gai).Text) Then
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
                                                                                '仕向け先ｺｰﾄﾞ
    Call UniCode_Conv(P_COMPO_O_REC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
                                                                                '事業部
    Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
                                                                                '国内外
    Call UniCode_Conv(P_COMPO_O_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
    Call UniCode_Conv(P_COMPO_O_REC.HIN_GAI, Text1(ptxHin_Gai).Text)
    Call UniCode_Conv(P_COMPO_O_REC.DATA_KBN, P_HEAD)
    Call UniCode_Conv(P_COMPO_O_REC.SEQNO, "000")

    Call UniCode_Conv(P_COMPO_O_REC.CLASS_CODE, Text1(ptxS_CLASS_CODE).Text)    'ｸﾗｽｺｰﾄﾞ
    Call UniCode_Conv(P_COMPO_O_REC.BIKOU, RichTextBox1(prchBIKOU).Text)        '備考
    
    Call UniCode_Conv(P_COMPO_O_REC.F_CLASS_CODE, Text1(ptxF_CLASS_CODE).Text)  '付加ｺｰﾄﾞ
    
    Call UniCode_Conv(P_COMPO_O_REC.N_CLASS_CODE, Text1(ptxN_CLASS_CODE).Text)  '内職ｺｰﾄﾞ
    
    Call UniCode_Conv(P_COMPO_O_REC.FILLER, "")

    Call UniCode_Conv(P_COMPO_O_REC.UPD_TANTO, Text1(ptxTanto_Code))            '更新担当者ｺｰﾄﾞ
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



    '構成マスタ(ﾎﾞﾃﾞｨ)出力
    
    Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
    
'    TDBGrid1.Refresh
    
    TDBGrid1(pGrdKOUSEI).Update


    K_SEQNO = 0
    G_SEQNO = 0
    D_SEQNO = 0



    For i = 1 To KOUSEI.UpperBound(1)


        If Trim(KOUSEI(i, ColKO_HIN_GAI)) = "" Then
        Else
                                                                                        '仕向け先ｺｰﾄﾞ
            Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
                                                                                        '事業部
            Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
                                                                                        '国内外
            Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
            Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, Text1(ptxHin_Gai).Text)
        
        
        
            Select Case Right(KOUSEI(i, ColKO_SYUBETSU), 2)
            
                Case KOSOU_KBN          '個装
                
                    Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_KOSOU)              'データ区分
                    
                    K_SEQNO = K_SEQNO + 10
                    
                    Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(K_SEQNO, "000"))  '追番
                    Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, "")                '種別
                
                Case GAISO_KBN          '外装
            
                    Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_GAISOU)              'データ区分
                    
                    G_SEQNO = G_SEQNO + 10
                    
                    Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(G_SEQNO, "000"))  '追番
                    Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, "")                '種別
            
            
                Case Else               '同梱
            
            
                    Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_DOUKON)             'データ区分
                    
                    D_SEQNO = D_SEQNO + 10
                    
                    Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(D_SEQNO, "000"))  '追番
                                                                                    '種別
                    Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, Right(KOUSEI(i, ColKO_SYUBETSU), 2))
            
            End Select
        
        
            Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, KOUSEI(i, ColKO_JGYOBU))         '子　事業部
            Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, KOUSEI(i, ColKO_NAIGAI))         '子　国内外
            Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, KOUSEI(i, ColKO_HIN_GAI))       '子　品番
                                                                                        '員数
            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(CDbl(KOUSEI(i, ColKO_QTY)), "000.00"))
            Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, KOUSEI(i, ColKO_BIKOU))           '子　備考
        
            Call UniCode_Conv(P_COMPO_K_REC.FILLER, "")
        
            Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, Text1(ptxTanto_Code).Text)       '更新担当者ｺｰﾄﾞ
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


            Call UniCode_Conv(K0_ITEM.JGYOBU, KOUSEI(i, ColKO_JGYOBU))         '子　事業部
            Call UniCode_Conv(K0_ITEM.NAIGAI, KOUSEI(i, ColKO_NAIGAI))         '子　国内外
            Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(i, ColKO_HIN_GAI))       '子　品番


            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound
                    
                        MsgBox "他端末でデータが、変更されています。単価登録処理を中止します。"
                        Update_Proc = False
                        Exit Function
                    
                    
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

            '工数
            Call UniCode_Conv(ITEMREC.S_KOUSU, Format(KOUSEI(i, ColS_KOUSU), "00000.00"))
            '集合梱包
            Call UniCode_Conv(ITEMREC.SEI_SYU_KON, Format(KOUSEI(i, ColSEI_SYU_KON), "000.00"))


            Do
                sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound
                    
                        MsgBox "他端末でデータが、変更されています。単価登録処理を中止します。"
                        Update_Proc = False
                        Exit Function
                    
                    
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

        End If
    Next i


End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
    Call Input_UnLock
    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    
    Call Input_UnLock

End Function

Private Sub Text1_LostFocus(Index As Integer)
    
Dim i   As Integer
    
    
    If ptxHin_Gai = Index Then
        If Trim(svHin_Gai) = (Text1(ptxHin_Gai).Text) Then
        Else
            For i = 2 To 5
                Command1(i).Enabled = False
            Next i
        
            Text1(ptxMAIN_KOUTEI_QTY01).Text = ""
        
        End If
    End If
End Sub
