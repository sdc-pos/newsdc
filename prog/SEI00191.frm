VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form SEI00191 
   Caption         =   "[請求システム]見積書一括発行処理"
   ClientHeight    =   12810
   ClientLeft      =   2025
   ClientTop       =   -3510
   ClientWidth     =   15780
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
   LockControls    =   -1  'True
   ScaleHeight     =   12810
   ScaleWidth      =   15780
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame1 
      Height          =   12615
      Left            =   0
      TabIndex        =   304
      Top             =   0
      Width           =   15855
      Begin VB.ListBox List3 
         Height          =   4140
         ItemData        =   "SEI00191.frx":0000
         Left            =   8520
         List            =   "SEI00191.frx":0002
         TabIndex        =   331
         Top             =   6600
         Width           =   6225
      End
      Begin VB.TextBox Text3 
         Height          =   4215
         Left            =   2760
         MultiLine       =   -1  'True
         TabIndex        =   330
         ToolTipText     =   "品番をコピーして下さい"
         Top             =   6600
         Width           =   5412
      End
      Begin VB.TextBox txtOUT_CNT 
         Alignment       =   1  '右揃え
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   10260
         Locked          =   -1  'True
         TabIndex        =   328
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "検索"
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
         Left            =   9240
         TabIndex        =   324
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtKIN_NG_CNT 
         Alignment       =   1  '右揃え
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   13800
         Locked          =   -1  'True
         TabIndex        =   323
         Top             =   11880
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "実行"
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
         Left            =   12360
         TabIndex        =   315
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "終了"
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
         Left            =   13695
         TabIndex        =   314
         Top             =   1080
         Width           =   855
      End
      Begin VB.ListBox List2 
         Height          =   4380
         ItemData        =   "SEI00191.frx":0004
         Left            =   8550
         List            =   "SEI00191.frx":0006
         TabIndex        =   313
         Top             =   1920
         Width           =   6225
      End
      Begin VB.TextBox txtIN_CNT 
         Alignment       =   1  '右揃え
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   7260
         Locked          =   -1  'True
         TabIndex        =   312
         Top             =   11040
         Width           =   855
      End
      Begin VB.TextBox txtOK_CNT 
         Alignment       =   1  '右揃え
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   13800
         Locked          =   -1  'True
         TabIndex        =   311
         Top             =   10800
         Width           =   855
      End
      Begin VB.TextBox txtNG_CNT 
         Alignment       =   1  '右揃え
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   13800
         Locked          =   -1  'True
         TabIndex        =   310
         Top             =   11280
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ｸﾘｱ"
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
         Left            =   3060
         TabIndex        =   309
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   4335
         Left            =   2760
         MultiLine       =   -1  'True
         TabIndex        =   308
         ToolTipText     =   "品番をコピーして下さい"
         Top             =   1920
         Width           =   5412
      End
      Begin VB.TextBox txtTANTO_CODE 
         Appearance      =   0  'ﾌﾗｯﾄ
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
         Left            =   3480
         MaxLength       =   5
         TabIndex        =   307
         Top             =   240
         Width           =   750
      End
      Begin VB.TextBox txtTanto_Name 
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
         Left            =   4320
         TabIndex        =   306
         TabStop         =   0   'False
         Top             =   240
         Width           =   2325
      End
      Begin VB.ComboBox cmbSHIMUKE 
         Appearance      =   0  'ﾌﾗｯﾄ
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
         Left            =   3495
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   305
         Top             =   600
         Width           =   2100
      End
      Begin VB.Label Label1 
         Caption         =   "更新結果"
         Height          =   255
         Index           =   123
         Left            =   10920
         TabIndex        =   334
         Top             =   6360
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "親品番"
         Height          =   255
         Index           =   122
         Left            =   8640
         TabIndex        =   333
         Top             =   6360
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "親品番"
         Height          =   255
         Index           =   121
         Left            =   2880
         TabIndex        =   332
         Top             =   6360
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "出力件数"
         Height          =   255
         Index           =   120
         Left            =   10200
         TabIndex        =   329
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "子品番"
         Height          =   255
         Index           =   119
         Left            =   2880
         TabIndex        =   327
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "親品番"
         Height          =   255
         Index           =   118
         Left            =   11160
         TabIndex        =   326
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "子品番"
         Height          =   255
         Index           =   117
         Left            =   8640
         TabIndex        =   325
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "金額ｱﾝﾏｯﾁ件数"
         Height          =   255
         Index           =   116
         Left            =   12255
         TabIndex        =   322
         Top             =   12000
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "更新結果"
         Height          =   255
         Index           =   115
         Left            =   13440
         TabIndex        =   321
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "読込み件数"
         Height          =   255
         Index           =   114
         Left            =   6000
         TabIndex        =   320
         Top             =   11160
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "ＯＫ件数"
         Height          =   255
         Index           =   113
         Left            =   12855
         TabIndex        =   319
         Top             =   10920
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "ＮＧ件数"
         Height          =   255
         Index           =   112
         Left            =   12855
         TabIndex        =   318
         Top             =   11400
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "担当者"
         Height          =   240
         Index           =   111
         Left            =   2640
         TabIndex        =   317
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "仕向先"
         Height          =   240
         Index           =   110
         Left            =   2640
         TabIndex        =   316
         Top             =   660
         Width           =   720
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
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
      Index           =   177
      Left            =   4515
      TabIndex        =   100
      Top             =   8760
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
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
      Index           =   170
      Left            =   11790
      MaxLength       =   1
      TabIndex        =   177
      Top             =   9840
      Width           =   225
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
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
      Index           =   168
      Left            =   11790
      MaxLength       =   1
      TabIndex        =   175
      Top             =   9480
      Width           =   225
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
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
      Index           =   167
      Left            =   11790
      MaxLength       =   10
      TabIndex        =   174
      Top             =   9120
      Width           =   1170
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
      Index           =   166
      Left            =   7245
      TabIndex        =   37
      Top             =   2280
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
      TabIndex        =   24
      Top             =   2040
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
      TabIndex        =   39
      Top             =   2280
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
      TabIndex        =   26
      Top             =   2040
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
      TabIndex        =   38
      Top             =   2280
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
      Left            =   12825
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
      Left            =   12405
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
      TabIndex        =   171
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
      TabIndex        =   170
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
      TabIndex        =   169
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
      TabIndex        =   168
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
      TabIndex        =   102
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
      TabIndex        =   101
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
      TabIndex        =   99
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
      TabIndex        =   98
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
      TabIndex        =   134
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
      TabIndex        =   133
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
      TabIndex        =   132
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
      TabIndex        =   131
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
      TabIndex        =   166
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
      TabIndex        =   165
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
      TabIndex        =   164
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
      TabIndex        =   163
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
      Left            =   14880
      Locked          =   -1  'True
      TabIndex        =   180
      TabStop         =   0   'False
      Top             =   960
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
      Left            =   14760
      Locked          =   -1  'True
      TabIndex        =   179
      TabStop         =   0   'False
      Top             =   480
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
      Left            =   14760
      Locked          =   -1  'True
      TabIndex        =   178
      TabStop         =   0   'False
      Top             =   120
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
      TabIndex        =   162
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
      TabIndex        =   161
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
      TabIndex        =   160
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
      TabIndex        =   159
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
      TabIndex        =   158
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
      TabIndex        =   157
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
      TabIndex        =   156
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
      TabIndex        =   155
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
      TabIndex        =   154
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
      TabIndex        =   153
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
      TabIndex        =   152
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
      TabIndex        =   151
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
      TabIndex        =   150
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
      TabIndex        =   149
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
      TabIndex        =   148
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
      TabIndex        =   147
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
      TabIndex        =   146
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
      TabIndex        =   145
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
      TabIndex        =   144
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
      TabIndex        =   143
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
      TabIndex        =   142
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
      TabIndex        =   141
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
      TabIndex        =   140
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
      TabIndex        =   139
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
      TabIndex        =   138
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
      TabIndex        =   137
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
      TabIndex        =   136
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
      TabIndex        =   135
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
      TabIndex        =   130
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
      TabIndex        =   129
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
      TabIndex        =   128
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
      TabIndex        =   127
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
      TabIndex        =   126
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
      TabIndex        =   125
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
      TabIndex        =   124
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
      TabIndex        =   123
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
      TabIndex        =   122
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
      TabIndex        =   121
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
      TabIndex        =   120
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
      TabIndex        =   118
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
      TabIndex        =   119
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
      TabIndex        =   117
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
      TabIndex        =   116
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
      TabIndex        =   115
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
      TabIndex        =   114
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
      TabIndex        =   113
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
      TabIndex        =   112
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
      TabIndex        =   111
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
      TabIndex        =   110
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
      TabIndex        =   109
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
      TabIndex        =   108
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
      TabIndex        =   107
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
      TabIndex        =   106
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
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   6120
      Width           =   750
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1575
      Index           =   0
      Left            =   10740
      TabIndex        =   172
      Top             =   6120
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   2778
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"SEI00191.frx":0008
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
      TabIndex        =   97
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
      TabIndex        =   96
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
      TabIndex        =   93
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
      TabIndex        =   90
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
      TabIndex        =   87
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
      TabIndex        =   84
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
      TabIndex        =   81
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
      TabIndex        =   78
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
      TabIndex        =   75
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
      TabIndex        =   72
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
      TabIndex        =   95
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
      TabIndex        =   92
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
      TabIndex        =   89
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
      TabIndex        =   86
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
      TabIndex        =   83
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
      TabIndex        =   80
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
      TabIndex        =   77
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
      TabIndex        =   74
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
      TabIndex        =   71
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
      TabIndex        =   94
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
      TabIndex        =   91
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
      TabIndex        =   88
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
      TabIndex        =   85
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
      TabIndex        =   82
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
      TabIndex        =   79
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
      TabIndex        =   76
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
      TabIndex        =   73
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
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   6120
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
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
      Top             =   750
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
      Left            =   9675
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
      Left            =   10200
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
      Left            =   10620
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
      Left            =   11040
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   720
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
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
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   3120
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
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   3120
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
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   3120
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
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   3120
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
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   3120
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
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   3120
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
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   3120
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
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   3120
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
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   3120
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
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   3120
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
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   3120
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
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   3120
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
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   2880
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
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   2880
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
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   2880
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
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   2880
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
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   2880
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
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   2880
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
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2880
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
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   2880
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
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   2880
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
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   2880
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
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   2880
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
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   2880
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
      TabIndex        =   42
      Top             =   2280
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
      TabIndex        =   41
      Top             =   2280
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
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2280
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
      TabIndex        =   36
      Top             =   2280
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
      TabIndex        =   35
      Top             =   2280
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
      TabIndex        =   34
      Top             =   2280
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
      TabIndex        =   33
      Top             =   2280
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
      TabIndex        =   32
      Top             =   2280
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
      TabIndex        =   31
      Top             =   2280
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
      TabIndex        =   30
      Top             =   2280
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
      TabIndex        =   29
      Top             =   2040
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
      TabIndex        =   28
      Top             =   2040
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
      TabIndex        =   27
      Top             =   2040
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
      TabIndex        =   23
      Top             =   2040
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
      TabIndex        =   22
      Top             =   2040
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
      TabIndex        =   21
      Top             =   2040
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
      TabIndex        =   20
      Top             =   2040
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
      MaxLength       =   5
      TabIndex        =   19
      Top             =   2040
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
      MaxLength       =   8
      TabIndex        =   18
      Top             =   2040
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
      MaxLength       =   8
      TabIndex        =   17
      Top             =   2040
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "単価更新"
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
      TabIndex        =   187
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
      TabIndex        =   186
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
      Appearance      =   0  'ﾌﾗｯﾄ
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
      TabIndex        =   185
      ToolTipText     =   "商品化単価を計算します(F9)"
      Top             =   0
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
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
      TabIndex        =   184
      ToolTipText     =   "商品化構成を保存します"
      Top             =   0
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   11025
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   182
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton Command1 
      Caption         =   "読込"
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
      TabIndex        =   183
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
      Left            =   420
      TabIndex        =   181
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
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   3120
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
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2880
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
      TabIndex        =   104
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
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   6120
      Width           =   750
   End
   Begin TrueDBGrid80.TDBDropDown TDBDropDown1 
      Height          =   2055
      Left            =   3120
      TabIndex        =   280
      Top             =   3960
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   3625
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
      Height          =   2295
      Index           =   0
      Left            =   0
      TabIndex        =   69
      Top             =   3480
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   4048
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
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1032"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=926"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2196"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2090"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(17)=   "Column(2).Button=1"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=1905"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1799"
      Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(4).Width=4710"
      Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=4604"
      Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=8708"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(29)=   "Column(5).Width=2037"
      Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=1931"
      Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=514"
      Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(34)=   "Column(6).Width=1879"
      Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=1773"
      Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=514"
      Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(39)=   "Column(7).Width=2143"
      Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=2037"
      Splits(0)._ColumnProps(42)=   "Column(7)._ColStyle=514"
      Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(44)=   "Column(8).Width=2117"
      Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=2011"
      Splits(0)._ColumnProps(47)=   "Column(8)._ColStyle=514"
      Splits(0)._ColumnProps(48)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(49)=   "Column(9).Width=2249"
      Splits(0)._ColumnProps(50)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(9)._WidthInPix=2143"
      Splits(0)._ColumnProps(52)=   "Column(9)._ColStyle=514"
      Splits(0)._ColumnProps(53)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(54)=   "Column(10).Width=3096"
      Splits(0)._ColumnProps(55)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(10)._WidthInPix=2990"
      Splits(0)._ColumnProps(57)=   "Column(10)._ColStyle=514"
      Splits(0)._ColumnProps(58)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(59)=   "Column(11).Width=3201"
      Splits(0)._ColumnProps(60)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(11)._WidthInPix=3096"
      Splits(0)._ColumnProps(62)=   "Column(11)._ColStyle=514"
      Splits(0)._ColumnProps(63)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(64)=   "Column(12).Width=3810"
      Splits(0)._ColumnProps(65)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(66)=   "Column(12)._WidthInPix=3704"
      Splits(0)._ColumnProps(67)=   "Column(12)._ColStyle=516"
      Splits(0)._ColumnProps(68)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(69)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(70)=   "Column(13).Width=3810"
      Splits(0)._ColumnProps(71)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(72)=   "Column(13)._WidthInPix=3704"
      Splits(0)._ColumnProps(73)=   "Column(13)._ColStyle=516"
      Splits(0)._ColumnProps(74)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(75)=   "Column(13).Order=14"
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
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=0,.fontsize=975"
      _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=128"
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
      TabIndex        =   167
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
      TabIndex        =   25
      Top             =   2040
      Width           =   960
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   975
      Index           =   1
      Left            =   10740
      TabIndex        =   173
      Top             =   8040
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   1720
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"SEI00191.frx":00C6
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00E0E0E0&
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
      Index           =   171
      Left            =   4410
      TabIndex        =   11
      Top             =   1800
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00E0E0E0&
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
      Index           =   172
      Left            =   6300
      TabIndex        =   12
      Top             =   1800
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00E0E0E0&
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
      Index           =   173
      Left            =   7245
      TabIndex        =   13
      Top             =   1800
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00E0E0E0&
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
      Index           =   174
      Left            =   8190
      TabIndex        =   14
      Top             =   1800
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00E0E0E0&
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
      Index           =   175
      Left            =   9135
      TabIndex        =   15
      Top             =   1800
      Width           =   960
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
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
      Index           =   176
      Left            =   10080
      TabIndex        =   16
      Top             =   1800
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
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
      Index           =   169
      Left            =   11790
      MaxLength       =   8
      TabIndex        =   176
      Top             =   9480
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "棚番区分"
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
      Index           =   109
      Left            =   11400
      TabIndex        =   303
      Top             =   720
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "標準棚番"
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
      Index           =   108
      Left            =   8760
      TabIndex        =   302
      Top             =   720
      Width           =   945
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      Caption         =   "+"
      Height          =   255
      Left            =   4305
      TabIndex        =   301
      Top             =   8760
      Width           =   225
   End
   Begin VB.Label Label1 
      Caption         =   "旧"
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
      Index           =   107
      Left            =   420
      TabIndex        =   300
      Top             =   1800
      Width           =   450
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "切替日"
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
      Index           =   106
      Left            =   9975
      TabIndex        =   299
      Top             =   1440
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "切替区分"
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
      Index           =   105
      Left            =   10740
      TabIndex        =   298
      Top             =   9960
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "(1:新規 2:現行)"
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
      Index           =   103
      Left            =   12000
      TabIndex        =   296
      Top             =   9600
      Width           =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "見積区分"
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
      Index           =   102
      Left            =   10950
      TabIndex        =   295
      Top             =   9600
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "仕様書��"
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
      Index           =   101
      Left            =   10950
      TabIndex        =   294
      Top             =   9240
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "見積書備考"
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
      Index           =   100
      Left            =   10740
      TabIndex        =   293
      Top             =   7800
      Width           =   1065
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
      TabIndex        =   292
      Top             =   1620
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
      TabIndex        =   291
      Top             =   1620
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
      TabIndex        =   290
      Top             =   1620
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
      TabIndex        =   268
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
      TabIndex        =   279
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
      TabIndex        =   274
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
      TabIndex        =   271
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
      TabIndex        =   278
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
      TabIndex        =   273
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
      TabIndex        =   270
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
      TabIndex        =   267
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
      TabIndex        =   272
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
      TabIndex        =   269
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
      TabIndex        =   266
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
      TabIndex        =   288
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
      TabIndex        =   287
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
      TabIndex        =   286
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
      TabIndex        =   285
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
      TabIndex        =   284
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
      TabIndex        =   283
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
      TabIndex        =   277
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
      TabIndex        =   282
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
      TabIndex        =   281
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
      TabIndex        =   276
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
      TabIndex        =   275
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
      TabIndex        =   265
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
      TabIndex        =   264
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
      TabIndex        =   263
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
      TabIndex        =   262
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
      TabIndex        =   261
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
      TabIndex        =   260
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
      TabIndex        =   259
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
      TabIndex        =   258
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
      TabIndex        =   257
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
      TabIndex        =   256
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
      TabIndex        =   255
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
      TabIndex        =   254
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
      TabIndex        =   253
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
      TabIndex        =   252
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
      TabIndex        =   251
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
      TabIndex        =   250
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
      TabIndex        =   249
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
      TabIndex        =   248
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
      TabIndex        =   247
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
      TabIndex        =   246
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
      TabIndex        =   245
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
      TabIndex        =   244
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
      TabIndex        =   243
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
      TabIndex        =   242
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
      TabIndex        =   241
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
      TabIndex        =   240
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
      TabIndex        =   239
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
      TabIndex        =   204
      Top             =   1620
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
      TabIndex        =   198
      Top             =   1620
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
      TabIndex        =   203
      Top             =   1620
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
      TabIndex        =   202
      Top             =   1620
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
      TabIndex        =   201
      Top             =   1620
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
      TabIndex        =   200
      Top             =   1620
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
      TabIndex        =   199
      Top             =   1620
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
      TabIndex        =   197
      Top             =   1620
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
      TabIndex        =   196
      Top             =   1620
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
      TabIndex        =   195
      Top             =   1620
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
      Left            =   10740
      TabIndex        =   238
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
      TabIndex        =   237
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
      TabIndex        =   236
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
      TabIndex        =   235
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
      TabIndex        =   234
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
      TabIndex        =   233
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
      TabIndex        =   232
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
      TabIndex        =   231
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
      TabIndex        =   230
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
      TabIndex        =   229
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
      TabIndex        =   228
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
      TabIndex        =   227
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
      TabIndex        =   226
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
      TabIndex        =   225
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
      TabIndex        =   224
      Top             =   2640
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
      TabIndex        =   223
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
      TabIndex        =   222
      Top             =   2640
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
      TabIndex        =   221
      Top             =   2640
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
      TabIndex        =   220
      Top             =   2640
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
      TabIndex        =   219
      Top             =   2640
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
      TabIndex        =   218
      Top             =   2640
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
      TabIndex        =   217
      Top             =   2640
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
      TabIndex        =   216
      Top             =   2640
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
      TabIndex        =   215
      Top             =   2640
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
      TabIndex        =   214
      Top             =   2640
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
      TabIndex        =   213
      Top             =   2640
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
      TabIndex        =   212
      Top             =   2640
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
      TabIndex        =   211
      Top             =   2640
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
      TabIndex        =   210
      Top             =   2640
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
      TabIndex        =   209
      Top             =   3120
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
      TabIndex        =   208
      Top             =   2880
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "新"
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
      Left            =   420
      TabIndex        =   205
      Top             =   2280
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "現行"
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
      Left            =   420
      TabIndex        =   194
      Top             =   2040
      Width           =   450
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
      TabIndex        =   193
      Top             =   510
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
      Left            =   10935
      TabIndex        =   192
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
      Left            =   10515
      TabIndex        =   191
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
      Left            =   9990
      TabIndex        =   190
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
      Left            =   345
      TabIndex        =   189
      Top             =   750
      Width           =   1065
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
      Left            =   735
      TabIndex        =   188
      Top             =   1020
      Width           =   675
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
      TabIndex        =   207
      Top             =   1440
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
      TabIndex        =   206
      Top             =   1440
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
      TabIndex        =   289
      Top             =   8520
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "単価切替日"
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
      Index           =   104
      Left            =   10740
      TabIndex        =   297
      Top             =   9600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "SEI00191"
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





Private Const ptxSHIYOU_NO% = 167           '仕様書��       2009.06.02
Private Const ptxMITSUMORI_KBN% = 168       '見積り区分     2009.06.02
'Private Const ptxTANKA_KIRIKAE_DT% = 169    '単価切替日付   2009.06.02
Private Const ptxKIRIKAE_KBN% = 170         '切替区分       2009.06.02
    







'------2009.07.24
Private Const ptxOLD_S_KOUSU_BAIKA% = 171       ' 旧  (売価)商品化工料
Private Const ptxOLD_S_SHIZAI_BAIKA% = 172      ' 旧  (売価)箱代

Private Const ptxOLD_S_GAISO_TANKA% = 173       ' 旧  外装単価
Private Const ptxOLD_S_PPSC_KAKO_KOSU% = 174    ' 旧  PPSC加工単価
Private Const ptxOLD_S_BU_KAKO_KOSU% = 175      ' 旧  BU加工単価
Private Const ptxTANKA_KIRIKAE_DT% = 176        ' 旧  単価切替日付
'------2009.07.24
Private Const ptxPLUS_KOUSU% = 177              ' プラス分工数  2009.09.17


Private IN_cnt  As Integer
Private OK_cnt  As Integer
Private NG_cnt  As Integer

Private KIN_NG_CNT  As Integer


'------------------------------------   'コンボ定義
Private Const pcmbSHIMUKE% = 0          '仕向け先


'------------------------------------   'リッチテキストボックス定義
Private Const prchBIKOU% = 0            '備考

Private Const prchM_BIKOU% = 1          '見積書備考         2009.06.02



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


Dim FUTAI_KBN       As String * 2       '付帯作業 2009.09.05

'-----------------------------------    ＥＸＣＥＬ 宛名＆住所

Dim EX_NAME1        As String           '宛名１
Dim EX_NAME2        As String           '宛名２

Dim EX_SYAMEI       As String           '自社　名称
Dim EX_ADDR1        As String           '自社　住所１
Dim EX_ADDR2        As String           '自社　住所２


Dim EX_CENTER_NAME  As String           'センター   名称
Dim EX_CENTER_ADDR1 As String           'センター   住所１
Dim EX_CENTER_ADDR2 As String           'センター   住所２

Dim EX_BIKOU1       As String           '備考１
Dim EX_BIKOU2       As String           '備考２



'Dim EX_JIGYOBU      As String

'2009.06.02
Dim EX_SHIZAI_T     As Variant          '資材対象
Dim EX_SHIZAI_F     As Boolean          '資材対象

Dim EX_DOUKON_T     As Variant          '同梱対象
Dim EX_DOUKON_F     As Boolean          '同梱対象

Dim EX_FUKA_T       As Variant          '付加作業
Dim EX_FUKA_F       As Boolean          '付加作業


'2009.06.02

Dim EX_BCR_CODE     As String           'ﾊﾞｰｺｰﾄﾞﾗﾍﾞﾙｺｰﾄﾞ



Dim EXCEL_TEMPLATE  As String           'EXCELﾃﾝﾌﾟﾚｰﾄ

'--------------------------------------- EXCEL用定数    2013.09.30
Private Const xlCalculationManual% = -4135
Private Const xlLeft% = -4131
Private Const xlCenter% = -4108
Private Const xlBottom% = -4107
Private Const xlNone% = -4142
Private Const xlTop% = -4160
Private Const xlContinuous% = 1
Private Const xlThin% = 2
Private Const xlAutomatic% = -4105
Private Const xlRight% = -4152
Private Const xlDiagonalDown% = 5
Private Const xlDiagonalUp% = 6
Private Const xlEdgeLeft% = 7
Private Const xlEdgeTop% = 8
Private Const xlEdgeBottom% = 9
Private Const xlEdgeRight% = 10
Private Const xlInsideVertical% = 11
Private Const xlInsideHorizontal% = 12
Private Const xlThick% = 4
Private Const xlCalculationAutomatic% = -4105
Private Const xlPortrait% = 1
Private Const xlOpenXMLWorkbook% = 51
'--------------------------------------- EXCEL用定数

'2011.01.21
Dim Insert_Pic       As String           '捺印蘭

Dim SYONIN_Pic       As String           '承認印

Dim MAIN_HIN_GAI    As String * 20

Dim Save_Dir        As String

Dim SEI0019_LOG     As String

'Private Const LAST_UPDATE_DAY$ = "[SEI0019] 2018.03.31 15:00"
Private Const LAST_UPDATE_DAY$ = "[SEI0019] 2019.04.26 10:00"






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
Dim Errflg  As Integer


    Select Case Index
    
        Case 0      '終了
            Unload Me
    
        Case 1      '検索（表示）
        
        
            If Detail_Disp_Proc(Errflg) Then
                Unload Me
            End If
        
        
        
        
        
            Text1(ptxAFT_SEI_LOT).SetFocus
        
        
        Case 2      '保存
            
            For i = ptxTanto_Code To ptxAFT_S_BU_KAKO_KOSU
            
                If Error_Check_Proc(i) Then
                    Exit Sub
                End If
            
            
            Next i
        
        
            '2009.06.02
            For i = ptxSHIYOU_NO To ptxKIRIKAE_KBN
            
                If Error_Check_Proc(i) Then
                    Exit Sub
                End If
            
            
            Next i
        
        
        
            If Grid_Error_Check_Proc() Then
                Exit Sub
            End If
        
        
            MESG = "商品化構成データを保存します。" & vbCrLf
            MESG = MESG & "　　種別／品番／員数" & vbCrLf
            MESG = MESG & "　　指図票備考" & vbCrLf
            MESG = MESG & "よろしいですか？" & vbCrLf
        
        
            ans = MsgBox(MESG, vbYesNo + vbDefaultButton2 + vbExclamation, "商品化構成の保存確認")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            
                If Detail_Disp_Proc(Errflg) Then
                    Unload Me
                End If
            
            
            
            
            
            
            
            
            
            End If
        
                    
            Text1(ptxAFT_SEI_LOT).SetFocus
        
        
        
        
        
        
        
        Case 3      '単価計算
        
        
            For i = ptxTanto_Code To ptxAFT_S_BU_KAKO_KOSU
            
                If Error_Check_Proc(i) Then
                    Exit Sub
                End If
            
            
            Next i
        
        
            If TANKA_KEISAN_Proc() Then
                Unload Me
            End If
        
        
        Case 4      '見積書発行
            
            
'            MsgBox "現在使用不可です！！"
'            Exit Sub
            
            If Estimate_Proc() Then
                Unload Me
            End If
        
        Case 5      '単価登録
            
            For i = ptxTanto_Code To ptxAFT_S_BU_KAKO_KOSU
            
                If Error_Check_Proc(i) Then
                    Exit Sub
                End If
            
            
            Next i
            
            
            
            
            
            
            
            
            
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
                    Unload Me
                End If
            
                If Detail_Disp_Proc(Errflg) Then
                    Unload Me
                End If
            
            
            
            
            
            
            
            
            
            End If
        
                    
            Text1(ptxAFT_SEI_LOT).SetFocus
                    
    
    
    End Select






End Sub

Private Sub Command2_Click(Index As Integer)

Dim i               As Integer

Dim wkLine          As Variant
Dim wkItem          As Variant

Dim ans             As Integer
Dim sts             As Integer

Dim S_DATETIME      As String


    Select Case Index
        Case 0
            Text2.Text = ""         '2018.03.12
            Text3.Text = ""         '2018.03.12
        Case 1
        
        
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, txtTANTO_CODE.Text)
        
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    txtTanto_Name.Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    txtTanto_Name.Text = ""
            
                    MsgBox "入力した項目はエラーです。(担当者)"
                    txtTANTO_CODE.SetFocus
                    Exit Sub
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Unload Me
                    Exit Sub
            
            End Select
        
                    
        
        
        
        
        
        
        
        
        
            Beep
            ans = MsgBox("[見積書一括発行]実行しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbNo Then
                Exit Sub
            End If
Call LOG_OUT(SEI0019_LOG, "見積書一括作成　処理開始[" & Now & "]")
        
            S_DATETIME = Now
        
            For i = 0 To 2
                Command2(i).Enabled = False
            Next i
            
            Text2.Locked = True
            Text3.Locked = True
            
            
            SEI00191.MousePointer = vbHourglass
            DoEvents
        
        
            List2.Clear
            
            IN_cnt = 0
            OK_cnt = 0
            NG_cnt = 0
            
            KIN_NG_CNT = 0
            
            txtIN_CNT.Text = Format(OK_cnt, "#,##0")
            txtOK_CNT.Text = Format(OK_cnt, "#,##0")
            txtNG_CNT.Text = Format(NG_cnt, "#,##0")
                                
            txtKIN_NG_CNT.Text = Format(KIN_NG_CNT, "#,##0")
                                
            wkLine = Split(Text2.Text, vbCrLf, -1)
    
    
            Combo1(pcmbSHIMUKE).ListIndex = cmbSHIMUKE.ListIndex           '2018.03.07
    
    
            For i = 0 To UBound(wkLine)
                
                
                wkItem = Split(wkLine(i), vbTab, -1)
                
                
                If UBound(wkItem) < 0 Then
                Else
                    IN_cnt = IN_cnt + 1
                    txtIN_CNT.Text = Format(IN_cnt, "#,##0")
                
                    MAIN_HIN_GAI = wkItem(0)
                
                    If Main_Update_Proc() Then
                        Unload Me
                    End If
                
                
                
                
                    DoEvents
                
                End If
    
            Next i
                    
                    
'>>>>>>>>>>>>>>>>>  親品番分    2018.03.12
            List3.Clear
            
            
            txtIN_CNT.Text = Format(OK_cnt, "#,##0")
            txtOK_CNT.Text = Format(OK_cnt, "#,##0")
            txtNG_CNT.Text = Format(NG_cnt, "#,##0")
                                
            txtKIN_NG_CNT.Text = Format(KIN_NG_CNT, "#,##0")
                                
            wkLine = Split(Text3.Text, vbCrLf, -1)
    
    
            Combo1(pcmbSHIMUKE).ListIndex = cmbSHIMUKE.ListIndex           '2018.03.07
    
    
            For i = 0 To UBound(wkLine)
                
                
                wkItem = Split(wkLine(i), vbTab, -1)
                
                
                If UBound(wkItem) < 0 Then
                Else
                    IN_cnt = IN_cnt + 1
                    txtIN_CNT.Text = Format(IN_cnt, "#,##0")
                
                    
                    
                    MAIN_HIN_GAI = wkItem(0)
                
                
                
                    If Main_Update_OYA_Proc() Then
                        Unload Me
                    End If
                
                
                
                
                    DoEvents
                
                End If
    
            Next i
'>>>>>>>>>>>>>>>>>  親品番分    2018.03.12
                    
                    
                    
            DoEvents
        
Call LOG_OUT(SEI0019_LOG, "見積書一括作成　正常終了[" & Now & "]")
            MsgBox "正常終了しました。[" & S_DATETIME & "→" & Now & "]"
        
            For i = 0 To 2
                Command2(i).Enabled = True
            Next i
        
            Text2.Locked = False
            Text3.Locked = False
            
        
        
        
           SEI00191.MousePointer = vbDefault
           DoEvents
        
        Case 2
    
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, txtTANTO_CODE.Text)
        
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    txtTanto_Name.Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    txtTanto_Name.Text = ""
            
                    MsgBox "入力した項目はエラーです。(担当者)"
                    txtTANTO_CODE.SetFocus
                    Exit Sub
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Unload Me
                    Exit Sub
            
            End Select
    
            List2.Clear
    
            IN_cnt = 0
            OK_cnt = 0
            NG_cnt = 0
            
            KIN_NG_CNT = 0
            
            txtIN_CNT.Text = Format(OK_cnt, "#,##0")
            txtOK_CNT.Text = Format(OK_cnt, "#,##0")
            txtNG_CNT.Text = Format(NG_cnt, "#,##0")
    
    
            txtKIN_NG_CNT.Text = Format(KIN_NG_CNT, "#,##0")
    
    
    
            txtOUT_CNT = ""
            IN_cnt = 0
    
            For i = 0 To 1
                Command2(i).Enabled = False
            Next i
            
            Text2.Locked = True
            Text3.Locked = True
            
            
            
            SEI00191.MousePointer = vbHourglass
            DoEvents
    
    
    
            wkLine = Split(Text2.Text, vbCrLf, -1)
    
            Combo1(pcmbSHIMUKE).ListIndex = cmbSHIMUKE.ListIndex           '2018.03.07
            
            
            For i = 0 To UBound(wkLine)
                
                
                wkItem = Split(wkLine(i), vbTab, -1)
                
                
                If UBound(wkItem) < 0 Then
                Else
                
                    MAIN_HIN_GAI = wkItem(0)
                
                    If COUNT_Proc() Then
                        Unload Me
                    End If
                
                
                
                
                    DoEvents
                
                End If
    
            Next i
'>>>>>>>>>>>>>>>>>  親品番分    2018.03.12

            List3.Clear
    
            
            txtIN_CNT.Text = Format(OK_cnt, "#,##0")
            txtOK_CNT.Text = Format(OK_cnt, "#,##0")
            txtNG_CNT.Text = Format(NG_cnt, "#,##0")
    
    
            txtKIN_NG_CNT.Text = Format(KIN_NG_CNT, "#,##0")
    
    
    
    
            For i = 0 To 1
                Command2(i).Enabled = False
            Next i
            SEI00191.MousePointer = vbHourglass
            DoEvents
    
    
    
            wkLine = Split(Text3.Text, vbCrLf, -1)
    
            Combo1(pcmbSHIMUKE).ListIndex = cmbSHIMUKE.ListIndex           '2018.03.07
            
            
            For i = 0 To UBound(wkLine)
                
                
                wkItem = Split(wkLine(i), vbTab, -1)
                
                
                If UBound(wkItem) < 0 Then
                Else
                
                    MAIN_HIN_GAI = wkItem(0)
                    
                    List3.AddItem MAIN_HIN_GAI


                    IN_cnt = IN_cnt + 1
                    txtOUT_CNT.Text = Format(IN_cnt, "#,##0")
                
                
                
                
                    DoEvents
                
                End If
    
            Next i


'>>>>>>>>>>>>>>>>>  親品番分    2018.03.12
    
    
    
        
            For i = 0 To 1
                Command2(i).Enabled = True
            Next i
        
            Text2.Locked = False
            Text3.Locked = False
            
        
        
            SEI00191.MousePointer = vbDefault
            DoEvents
        
        
        
        Case 3
            Unload Me
    
    
    
    
    
    
    End Select
    
    
    






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
        "[請求システム]見積書一括作成処理", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)



    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                
                                
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
                                
                                '構成マスタＯＰＥＮ
    If wP_COMPO_Open(BtOpenNomal) Then
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
    
    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc
    
    '仕向け先のセット
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD, 0) Then
        Unload Me
    End If
    Combo1(pcmbSHIMUKE).ListIndex = 0

                                
                                
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
                                
                                
                                
                                '付帯作業の獲得 2009.09.05
    If GetIni(App.EXEName, "FUTAI", App.EXEName, c) Then
        FUTAI_KBN = ""
    Else
        FUTAI_KBN = Trim(c)
    End If
                                
                                
                                
                                
                                
                                
                                
                                '見積書 宛名１
    If GetIni("Estimate", "NAME1", App.EXEName, c) Then
        EX_NAME1 = ""
    Else
        EX_NAME1 = Trim(c)
    End If
                                '見積書 宛名２
    If GetIni("Estimate", "NAME2", App.EXEName, c) Then
        EX_NAME2 = ""
    Else
        EX_NAME2 = Trim(c)
    End If
                                '見積書 自社　名称
    If GetIni("Estimate", "SYAMEI", App.EXEName, c) Then
        EX_SYAMEI = ""
    Else
        EX_SYAMEI = Trim(c)
    End If
                                '見積書 自社　住所１
    If GetIni("Estimate", "ADDR1", App.EXEName, c) Then
        EX_ADDR1 = ""
    Else
        EX_ADDR1 = Trim(c)
    End If
                                '見積書 自社　住所２
    If GetIni("Estimate", "ADDR2", App.EXEName, c) Then
        EX_ADDR2 = ""
    Else
        EX_ADDR2 = Trim(c)
    End If
                                '見積書 センター   名称
    If GetIni("Estimate", "CENTER_NAME", App.EXEName, c) Then
        EX_CENTER_NAME = ""
    Else
        EX_CENTER_NAME = Trim(c)
    End If
                                '見積書 センター   住所１
    If GetIni("Estimate", "CENTER_ADDR1", App.EXEName, c) Then
        EX_CENTER_ADDR1 = ""
    Else
        EX_CENTER_ADDR1 = Trim(c)
    End If
                                '見積書 センター   住所２
    If GetIni("Estimate", "CENTER_ADDR2", App.EXEName, c) Then
        EX_CENTER_ADDR2 = ""
    Else
        EX_CENTER_ADDR2 = Trim(c)
    End If
                                
                                
                                '見積書 備考１
    If GetIni("Estimate", "BIKOU1", App.EXEName, c) Then
        EX_BIKOU1 = ""
    Else
        EX_BIKOU1 = Trim(c)
    End If
                                '見積書 備考２
    If GetIni("Estimate", "BIKOU2", App.EXEName, c) Then
        EX_BIKOU2 = ""
    Else
        EX_BIKOU2 = Trim(c)
    End If
                                
                                '見積書　表示用事業部
'    If GetIni("Estimate", "JIGYOBU", App.EXEName, c) Then
'        EX_JIGYOBU = ""
'    Else
'        EX_JIGYOBU = Trim(c)
'    End If


'2009.06.02
                                '資材対象種別
    If GetIni("Estimate", "EX_SHIZAI", App.EXEName, c) Then
        EX_SHIZAI_F = False
    Else
        EX_SHIZAI_F = True
        EX_SHIZAI_T = Split(Trim(c), ",", -1)
    End If
                                '同梱対象種別
    If GetIni("Estimate", "EX_DOUKON", App.EXEName, c) Then
        EX_DOUKON_F = False
    Else
        EX_DOUKON_F = True
        EX_DOUKON_T = Split(Trim(c), ",", -1)
    End If

                                '付加作業対象種別
    If GetIni("Estimate", "EX_FUKA", App.EXEName, c) Then
        EX_FUKA_F = False
    Else
        EX_FUKA_F = True
        EX_FUKA_T = Split(Trim(c), ",", -1)
    End If

                                'ﾊﾞｰｺｰﾄﾞﾗﾍﾞﾙｺｰﾄﾞ
    If GetIni("Estimate", "EX_BCR_CODE", App.EXEName, c) Then
        EX_BCR_CODE = ""
    Else
        EX_BCR_CODE = Trim(c)
    End If

'2009.06.02
                                
                                
    If GetIni("Estimate", "EXCEL_TEMPLATE", App.EXEName, c) Then
        EXCEL_TEMPLATE = ""
    Else
        EXCEL_TEMPLATE = Trim(c)
    End If
                                
                                
                                
'2011.01.21
    If GetIni("Estimate", "INSERT_PIC", App.EXEName, c) Then
        Insert_Pic = ""
    Else
        Insert_Pic = Trim(c)
    End If


    If GetIni("Estimate", "SYONIN_PIC", App.EXEName, c) Then
        SYONIN_Pic = ""
    Else
        SYONIN_Pic = Trim(c)
    End If

    If GetIni("Estimate", "Save_Dir", App.EXEName, c) Then
        Save_Dir = ""
    Else
        Save_Dir = Trim(c)
    End If



    If GetIni(App.EXEName, "SEI0019_LOG", App.EXEName, c) Then
        SEI0019_LOG = ""
    Else
        SEI0019_LOG = Trim(c)
    End If


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
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD, 0, 1) Then
        Unload Me
    End If
    Combo1(pcmbSHIMUKE).ListIndex = 0

    '種別セット
    If SYUBETSU_Set_Proc() Then
        Unload Me
    End If

    SEI00191.Caption = SEI00191.Caption & " " & LAST_UPDATE_DAY

    Call Init_Proc
    
    cmbSHIMUKE.ListIndex = 0
    
    
    
    
    
    txtTANTO_CODE.SetFocus
End Sub

Private Sub Form_Unload(CANCEL As Integer)
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


    SEI00191.MousePointer = vbHourglass

    Call Ctrl_Lock(SEI00191)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(SEI00191)


    SEI00191.MousePointer = vbDefault

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

Dim row         As Integer
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
    
    
    
    
    row = 0
    KOTEI_NO = 0
    For i = 1 To 10
        
        If GetIni("KOUTEI", "BEF" & Format(i, "00"), "SEI_SYS", c) Then
        Else
            If Trim(c) = "*" Then
            Else
                row = row + 1
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
                row = row + 1
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
                row = row + 1
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
                
                
'                If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                    
                    
                    If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                        
                        If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) = 0 Then      '2010.02.22
                        
                            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN) = "0.00"
                        Else
                            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN)) / CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY))), 2), "#,##0.00")
                        End If
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
        
        Select Case StrConv(ITEMREC.SEI_KBN, vbUnicode)
        
        
            Case "1"
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = "別売"
            Case "2"
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = "支給"
            Case Else
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
        
        End Select
            
        '売上金額計
        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = 0
    
        For i = 0 To UBound(SHIZAI_T)
        
            If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = SHIZAI_T(i) Then
'                If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                    If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = GAISO_KBN Then
                        
                        If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) = 0 Then
                            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = "0.00"
                        Else
                            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) / CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY))), 2), "#,##0.00")
                        End If
                    Else
                        KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN))), 2), "#,##0.00")
                    End If
                End If
            Else
            
                If KUSATU_F Then
'                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
            
                        If Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2) = GAISO_KBN Then
                        
                            If CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) = 0 Then      '2010.02.22
                            
                                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN_KUSATU) = "0.00"
                            Else
                                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URIKIN_KUSATU) = Format(ToRoundUp(CCur(CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) / CDbl(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY))), 2), "#,##0.00")
                            End If
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


Private Sub TDBGrid1_BeforeInsert(Index As Integer, CANCEL As Integer)
    
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
        
      
'            Text1(ptxMAIN_KOUTEI_QTY01).Text = ""
        
        
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
Call LOG_OUT(SEI0019_LOG, "担当者エラー= " & Text1(ptxTanto_Code).Text)
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
Call LOG_OUT(SEI0019_LOG, "品番未登録エラー= " & Text1(ptxHin_Gai).Text)
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
        
        

        Case ptxOLD_S_BU_KAKO_KOSU          ' 旧  BU加工単価
        
        
        
        
        
        
        
        
        
        
        
        Case ptxOLD_S_KOUSU_BAIKA           '旧(売価)商品化工料
        
        
            If Text1(ptxOLD_S_KOUSU_BAIKA).Text = "" Then
                Text1(ptxOLD_S_KOUSU_BAIKA).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxOLD_S_KOUSU_BAIKA).Text) Then
                MsgBox "入力した項目はエラーです。(工料売価)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxOLD_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxOLD_S_KOUSU_BAIKA).Text), "#0.00")
            End If
        
        
        
        
        
        
        
        
        
        
        
        
        
        Case ptxOLD_S_SHIZAI_BAIKA          '旧(売価)箱代

            If Text1(ptxOLD_S_SHIZAI_BAIKA).Text = "" Then
                Text1(ptxOLD_S_SHIZAI_BAIKA).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxOLD_S_SHIZAI_BAIKA).Text) Then
                MsgBox "入力した項目はエラーです。(資材売価)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxOLD_S_SHIZAI_BAIKA).Text = Format(CDbl(Text1(ptxOLD_S_SHIZAI_BAIKA).Text), "#0.00")
            End If


        Case ptxOLD_S_GAISO_TANKA           '旧外装単価
        
        
            If Text1(ptxOLD_S_GAISO_TANKA).Text = "" Then
                Text1(ptxOLD_S_GAISO_TANKA).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxOLD_S_GAISO_TANKA).Text) Then
                MsgBox "入力した項目はエラーです。(外装単価)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxOLD_S_GAISO_TANKA).Text = Format(CDbl(Text1(ptxOLD_S_GAISO_TANKA).Text), "#0.00")
            End If
        
        
        
        
        
        Case ptxOLD_S_PPSC_KAKO_KOSU        '旧PPSC加工単価
            
            If Text1(ptxOLD_S_PPSC_KAKO_KOSU).Text = "" Then
                Text1(ptxOLD_S_PPSC_KAKO_KOSU).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxOLD_S_PPSC_KAKO_KOSU).Text) Then
                MsgBox "入力した項目はエラーです。(PPSC加工単価)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxOLD_S_PPSC_KAKO_KOSU).Text = Format(CDbl(Text1(ptxOLD_S_PPSC_KAKO_KOSU).Text), "#0.00")
            End If
        
        Case ptxOLD_S_BU_KAKO_KOSU          '旧BU加工単価
    
            If Text1(ptxOLD_S_BU_KAKO_KOSU).Text = "" Then
                Text1(ptxOLD_S_BU_KAKO_KOSU).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxOLD_S_BU_KAKO_KOSU).Text) Then
                MsgBox "入力した項目はエラーです。(PPSC加工単価)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxOLD_S_BU_KAKO_KOSU).Text = Format(CDbl(Text1(ptxOLD_S_BU_KAKO_KOSU).Text), "#0.00")
            End If
        
        
        
        




        
        
        
        
        
        
        
        
        
        
        
        
        
        
        Case ptxBEF_SEI_LOT                 '変更前　   ロット数
        
            If Text1(ptxBEF_SEI_LOT).Text = "" Then
Call LOG_OUT(SEI0019_LOG, Text1(ptxHin_Gai).Text & " ロットエラー= " & Text1(ptxBEF_SEI_LOT).Text)
Exit Function

            Else
            
                If Not IsNumeric(Text1(ptxBEF_SEI_LOT).Text) Then
                    
Call LOG_OUT(SEI0019_LOG, Text1(ptxHin_Gai).Text & " ロットエラー= " & Text1(ptxBEF_SEI_LOT).Text)
Exit Function
                    
                    
                    MsgBox "入力した項目はエラーです。(ロット数)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_SEI_LOT).Text = Format(CLng(Text1(ptxBEF_SEI_LOT).Text), "#0")
                End If
            
            End If
        
        Case ptxBEF_SEI_RATE                '           分レート
        
            If Text1(ptxBEF_SEI_RATE).Text = "" Then
            
            Else
            
            
                If Not IsNumeric(Text1(ptxBEF_SEI_RATE).Text) Then
                    MsgBox "入力した項目はエラーです。(分レート)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_SEI_RATE).Text = Format(CLng(Text1(ptxBEF_SEI_RATE).Text), "#0")
                End If
            End If
        
        
        Case ptxBEF_S_KOUSU                 '           分レート
        
        
            If Text1(ptxBEF_S_KOUSU).Text = "" Then
            
            Else
                If Not IsNumeric(Text1(ptxBEF_S_KOUSU).Text) Then
                    MsgBox "入力した項目はエラーです。(工数)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_S_KOUSU).Text = Format(CDbl(Text1(ptxBEF_S_KOUSU).Text), "#0.0")
                End If
            End If
        
        
        
        
        
        
        
        
        Case ptxBEF_S_KOUSU_GENKA           '           (原価)商品化工料
        
            If Text1(ptxBEF_S_KOUSU_GENKA).Text = "" Then
            
            Else
            
            
                If Not IsNumeric(Text1(ptxBEF_S_KOUSU_GENKA).Text) Then
                    MsgBox "入力した項目はエラーです。(工料原価)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_S_KOUSU_GENKA).Text = Format(CDbl(Text1(ptxBEF_S_KOUSU_GENKA).Text), "#0.00")
                End If
            End If
        
        
        Case ptxBEF_S_KOUSU_BAIKA           '           (売価)商品化工料
        
        
            If Text1(ptxBEF_S_KOUSU_BAIKA).Text = "" Then
            
            Else
            
                If Not IsNumeric(Text1(ptxBEF_S_KOUSU_BAIKA).Text) Then
                    MsgBox "入力した項目はエラーです。(工料売価)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_S_KOUSU_BAIKA).Text = Format(CDbl(Text1(ptxBEF_S_KOUSU_BAIKA).Text), "#0.00")
                End If
            End If
        
        
        
        
        
        
        
        Case ptxBEF_S_SHIZAI_GENKA          '           (原価)箱代
        
        
            If Text1(ptxBEF_S_SHIZAI_GENKA).Text = "" Then
            
            Else
            
            
                If Not IsNumeric(Text1(ptxBEF_S_SHIZAI_GENKA).Text) Then
                    MsgBox "入力した項目はエラーです。(資材原価)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_S_SHIZAI_GENKA).Text = Format(CDbl(Text1(ptxBEF_S_SHIZAI_GENKA).Text), "#0.00")
                End If
            End If
        
        
        
        
        Case ptxBEF_S_SHIZAI_BAIKA          '           (売価)箱代

            If Text1(ptxBEF_S_SHIZAI_BAIKA).Text = "" Then
            
            Else
            
                If Not IsNumeric(Text1(ptxBEF_S_SHIZAI_BAIKA).Text) Then
                    MsgBox "入力した項目はエラーです。(資材売価)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_S_SHIZAI_BAIKA).Text = Format(CDbl(Text1(ptxBEF_S_SHIZAI_BAIKA).Text), "#0.00")
                End If
            End If

        Case ptxBEF_S_GAISO_TANKA           '           外装単価
        
        
            If Text1(ptxBEF_S_GAISO_TANKA).Text = "" Then
            
            Else
            
                If Not IsNumeric(Text1(ptxBEF_S_GAISO_TANKA).Text) Then
                    MsgBox "入力した項目はエラーです。(外装単価)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_S_GAISO_TANKA).Text = Format(CDbl(Text1(ptxBEF_S_GAISO_TANKA).Text), "#0.00")
                End If
            End If
        
        
        
        
        Case ptxBEF_S_PPSC_KAKO_KOSU        '           PPSC加工単価
            
            If Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text = "" Then
            
            Else
            
            
                If Not IsNumeric(Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text) Then
                    MsgBox "入力した項目はエラーです。(PPSC加工単価)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text = Format(CDbl(Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text), "#0.00")
                End If
        
            End If
        
        
        Case ptxBEF_S_BU_KAKO_KOSU          '           BU加工単価
    
            If Text1(ptxBEF_S_BU_KAKO_KOSU).Text = "" Then
            
            Else
            
            
                If Not IsNumeric(Text1(ptxBEF_S_BU_KAKO_KOSU).Text) Then
                    MsgBox "入力した項目はエラーです。(PPSC加工単価)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxBEF_S_BU_KAKO_KOSU).Text = Format(CDbl(Text1(ptxBEF_S_BU_KAKO_KOSU).Text), "#0.00")
                End If
            End If
        
        
        
        Case ptxBEF_S_KOUSU_SET_DATE        '           設定日
        
        
        
            If Text1(ptxBEF_S_KOUSU_SET_DATE).Text = "" Then
            
            Else
            
            
            
                If Len(Text1(ptxBEF_S_KOUSU_SET_DATE).Text) < 8 Then
                    MsgBox "入力した項目はエラーです。(設定日)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
            
                    If Not IsDate(Mid(Text1(ptxBEF_S_KOUSU_SET_DATE).Text, 1, 4) & "/" & _
                                    Mid(Text1(ptxBEF_S_KOUSU_SET_DATE).Text, 5, 2) & "/" & _
                                    Mid(Text1(ptxBEF_S_KOUSU_SET_DATE).Text, 7, 2)) Then
                        MsgBox "入力した項目はエラーです。(設定日)"
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
                End If
            End If
        
        Case ptxBEF_SEI_TANKA_TANTO         '          担当者
            If Text1(ptxBEF_SEI_TANKA_TANTO).Text = "" Then
            Else
                
                Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxBEF_SEI_TANKA_TANTO).Text)
    
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
            End If
        
        
        
        
        
        
        
        
        
        
        Case ptxBEF_SE_TANKA_MEMO           '          メモ
        
        
        
        
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
    
        Case ptxAFT_S_GAISO_TANKA       '外装単価
    
            If Text1(ptxAFT_S_GAISO_TANKA).Text = "" Then
                Text1(ptxAFT_S_GAISO_TANKA).Text = "0.00"
            End If
            
            If Not IsNumeric(Text1(ptxAFT_S_GAISO_TANKA).Text) Then
                MsgBox "入力した項目はエラーです。(外装単価)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxAFT_S_GAISO_TANKA).Text = Format(CDbl(Text1(ptxAFT_S_GAISO_TANKA).Text), "#0.00")
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
    
    
    
    
        Case ptxPLUS_KOUSU              '2009.09.17　PLUS工数
            
            If Not IsNumeric(Text1(ptxPLUS_KOUSU).Text) Then
                Text1(ptxPLUS_KOUSU).Text = "0"
            Else
                Text1(ptxPLUS_KOUSU).Text = Format(CInt(Text1(ptxPLUS_KOUSU).Text), "#0")
            End If
    
    
    
    
        Case ptxSHIYOU_NO               '仕様書��       2009.06.02
        Case ptxMITSUMORI_KBN           '見積区分       2009.06.02
        
            If Text1(ptxMITSUMORI_KBN).Text = "1" Or Text1(ptxMITSUMORI_KBN).Text = "2" Then
            Else
                MsgBox "入力した項目はエラーです。(見積区分)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        Case ptxTANKA_KIRIKAE_DT        '単価切替日付   2009.06.02
            
            If Trim(Text1(ptxTANKA_KIRIKAE_DT).Text) = "" Then
            Else
                If Len(Trim(Text1(ptxTANKA_KIRIKAE_DT).Text)) <> 8 Then
                    MsgBox "入力した項目はエラーです。(単価切替日付)"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    If IsDate(Mid(Text1(ptxTANKA_KIRIKAE_DT).Text, 1, 4) & "/" & Mid(Text1(ptxTANKA_KIRIKAE_DT).Text, 5, 2) & "/" & Mid(Text1(ptxTANKA_KIRIKAE_DT).Text, 7, 2)) Then
                    Else
                        MsgBox "入力した項目はエラーです。(単価切替日付)"
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
                End If
            End If
                
        
        Case ptxKIRIKAE_KBN             '切替区分       2009.06.02
    
    
    End Select
        
        
        
        
        
        
    Error_Check_Proc = False
    

End Function


Private Function Code_Set_Proc(Index As Integer, KBN As String, Mode As Integer, Optional flg As Integer = 0) As Integer
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
    If flg = 1 Then
        cmbSHIMUKE.Clear
    End If
    
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
        
        
        
        If flg = 1 Then
            cmbSHIMUKE.AddItem StrConv(P_CODEREC.C_RNAME, vbUnicode) & "            " & _
                                    Left(StrConv(P_CODEREC.C_Code, vbUnicode), Key_Len) & wkOption
        End If
        
        
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
    
Dim row         As Long
    
Dim FAST_FLG    As Boolean
    
    P_COMPO_Disp_Proc = True
'    Call Input_Lock             '2008.01.15
    
        
    
    
            

    

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
            
            
'            Call Input_UnLock           '2008.01.15
            P_COMPO_Disp_Proc = sts
            Exit Function
    End Select

    '--------------------------------   「子」情報
    
    Set KOUSEI = Nothing
    
    
    
    If FAST_FLG Then
    
        row = Min_Row - 1
        
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
'                    Call Input_UnLock             '2008.01.15
                    Call File_Error(sts, BtOpGetNext, "構成マスタ")
                    Exit Function
            End Select
            
            
If Trim(StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)) = "18" Then
    Debug.Print
End If
            
            If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_KOSOU Then
                Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, KOSOU_KBN)
            End If
            If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_GAISOU Then
                Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, GAISO_KBN)
            End If
            
            row = row + 1
                        
            If Grid_Set_Proc(row) Then
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















    
'    Call Input_UnLock

    
    
    P_COMPO_Disp_Proc = False

End Function
Private Function Grid_Set_Proc(row As Long) As Integer
'----------------------------------------------------------------------------
'                   構成マスタ==>Gridテーブル
'----------------------------------------------------------------------------

Dim sts As Integer
Dim i   As Integer
Dim j   As Integer
    
    Grid_Set_Proc = True

    

    KOUSEI.ReDim Min_Row, row, Min_Col, Max_Col
    
    
    '事業部
    KOUSEI(row, ColKO_JGYOBU) = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
    '国内外
    KOUSEI(row, ColKO_NAIGAI) = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
    
    '種別
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        
    Select Case sts
        Case BtNoErr
            KOUSEI(row, ColKO_SYUBETSU) = Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode)) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), 2)
        
        
        
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "コードマスタ")
            Exit Function
    
    End Select
    '品番
    KOUSEI(row, ColKO_HIN_GAI) = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
    '品名
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Select Case sts
        Case BtNoErr
            KOUSEI(row, ColKO_HIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
        Case BtErrKeyNotFound
            KOUSEI(row, ColKO_HIN_NAME) = "未登録品番"
            
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
        KOUSEI(row, ColKO_QTY) = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
    Else
        KOUSEI(row, ColKO_QTY) = "1.00"
    End If
    
    '仕入単価
    If IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
        KOUSEI(row, ColG_ST_SHITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)), "#0.00")
    Else
        KOUSEI(row, ColG_ST_SHITAN) = "0.00"
    End If
    '販売単価
'    If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
'        KOUSEI(row, ColG_ST_URITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
'    Else
'        KOUSEI(row, ColG_ST_URITAN) = "0.00"
'    End If
    
    
    Select Case StrConv(ITEMREC.SEI_KBN, vbUnicode)
    
        Case "1"
            KOUSEI(row, ColG_ST_URITAN) = "別売"
        Case "2"
            KOUSEI(row, ColG_ST_URITAN) = "支給"
        Case Else
            If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                KOUSEI(row, ColG_ST_URITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
            Else
                KOUSEI(row, ColG_ST_URITAN) = "0.00"
            End If
    End Select
    
    
    
    
    
    
    
    
    
    
    
    '仕入金額計
    KOUSEI(row, ColG_ST_SHIKIN) = 0

    For i = 0 To UBound(SHIZAI_T)
        If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = SHIZAI_T(i) Then
            
            
'            If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
            If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                
                If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                    
                    If CDbl(KOUSEI(row, ColKO_QTY)) = 0 Then '2010.02.22
                        KOUSEI(row, ColG_ST_SHIKIN) = "0.00"
                    Else
                        KOUSEI(row, ColG_ST_SHIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(row, ColG_ST_SHITAN)) / CDbl(KOUSEI(row, ColKO_QTY))), 2), "#,##0.00")
                    End If
                Else
                    KOUSEI(row, ColG_ST_SHIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(row, ColKO_QTY)) * CDbl(KOUSEI(row, ColG_ST_SHITAN))), 2), "#,##0.00")
                End If
            End If
            Exit For
        End If
    
    Next i
    If CDbl(KOUSEI(row, ColG_ST_SHIKIN)) = 0 Then
        KOUSEI(row, ColG_ST_SHIKIN) = ""
    End If
    
    '売上金額計
    KOUSEI(row, ColG_ST_URIKIN) = 0
    KOUSEI(row, ColG_ST_URIKIN_KUSATU) = 0

    For i = 0 To UBound(SHIZAI_T)
    
'        If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
        If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
    
    
            If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = SHIZAI_T(i) Then
                If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                
'                    KOUSEI(row, ColG_ST_URIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(row, ColG_ST_URITAN)) / CDbl(KOUSEI(row, ColKO_QTY))), 2), "#,##0.00")
                
                
                
                    If CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) = 0 Then  '2010.02.22
                        KOUSEI(row, ColG_ST_URIKIN) = "0.00"
                    Else
                        KOUSEI(row, ColG_ST_URIKIN) = CDbl(ToHalfAdjust(1 / CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), 3))
                    End If
                    KOUSEI(row, ColG_ST_URIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(row, ColG_ST_URITAN)) * CDbl(KOUSEI(row, ColG_ST_URIKIN))), 2), "#,##0.00")
                
                
                
                
                
                Else
                    KOUSEI(row, ColG_ST_URIKIN) = Format(ToRoundUp(CCur(CDbl(KOUSEI(row, ColKO_QTY)) * CDbl(KOUSEI(row, ColG_ST_URITAN))), 2), "#,##0.00")
                End If
    
            
            Else
            
                If KUSATU_F Then
            
                    If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                    
'                        KOUSEI(row, ColG_ST_URIKIN_KUSATU) = Format(ToRoundUp(CCur(CDbl(KOUSEI(row, ColG_ST_URITAN)) / CDbl(KOUSEI(row, ColKO_QTY))), 2), "#,##0.00")
                    
                    
                    
                    
                        If CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) = 0 Then  '2010.02.22
                            KOUSEI(row, ColG_ST_URIKIN_KUSATU) = 0
                        Else
                            KOUSEI(row, ColG_ST_URIKIN_KUSATU) = CDbl(ToHalfAdjust(1 / CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), 3))
                        End If
                        KOUSEI(row, ColG_ST_URIKIN_KUSATU) = Format(ToRoundUp(CCur(CDbl(KOUSEI(row, ColG_ST_URITAN)) * CDbl(KOUSEI(row, ColG_ST_URIKIN_KUSATU))), 2), "#,##0.00")
                    
                    
                    
                    
                    
                    Else
                        KOUSEI(row, ColG_ST_URIKIN_KUSATU) = Format(ToRoundUp(CCur(CDbl(KOUSEI(row, ColKO_QTY)) * CDbl(KOUSEI(row, ColG_ST_URITAN))), 2), "#,##0.00")
                    End If
                
                
                End If
            
            
            
            End If
        End If
    Next i
    
    
    If CDbl(KOUSEI(row, ColG_ST_URIKIN)) = 0 Then
        KOUSEI(row, ColG_ST_URIKIN) = ""
    End If
    
    
    If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
        KOUSEI(row, ColS_KOUSU) = ""
        KOUSEI(row, ColSEI_SYU_KON) = ""
    Else
        '作業時間
        If IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
            KOUSEI(row, ColS_KOUSU) = Format(CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode)), "#")
        Else
            KOUSEI(row, ColS_KOUSU) = ""
        End If
        '集合梱包
        If IsNumeric(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)) Then
            KOUSEI(row, ColSEI_SYU_KON) = Format(CInt(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)), "#")
        Else
            KOUSEI(row, ColSEI_SYU_KON) = ""
        End If
    End If
    
    
    '備考
    KOUSEI(row, ColKO_BIKOU) = StrConv(P_COMPO_K_REC.KO_BIKOU, vbUnicode)
    
    
    
    
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






Private Function Estimate_Proc() As Integer
'----------------------------------------------------------------------------
'                   ＥＸＣＥＬ（御見積書）出力
'       2009.06.02
'----------------------------------------------------------------------------


'-----------------------------------------------    2013.09.30 EXCEL Ver対応
'Dim excelApplication    As Excel.Application   2013.09.30
''Dim excelWorkBooks      As excel.Workbooks
'Dim excelWorkBook       As Excel.Workbook      2013.09.30
'Dim excelSheet          As Excel.Worksheet     2013.09.30


Dim excelApplication    As Object
Dim excelWorkBook       As Object
Dim excelSheet          As Object

'-----------------------------------------------    2013.09.30 EXCEL Ver対応

Dim i                   As Integer
Dim j                   As Integer

Dim com                 As Integer
Dim sts                 As Integer

Dim wkBikou             As Variant

Dim row                 As Integer
Dim SHIZAI_TOTAL_ROW    As Integer
Dim FUKA_TOTAL_ROW      As Integer
Dim Fsw                 As Integer      '2018.01.05
Dim TOTAL_ROW           As Integer
    
    
Dim wkint               As Integer
Dim BEF_KOTEI           As Double
Dim AFT_KOTEI           As Double
Dim MAIN_KOTEI          As Double
    
    
Dim stTime              As String
    
    
Dim wkNum1              As Currency
Dim wkNum2              As Currency
    
    
    
    
'2011.01.11
Dim S_Start             As String
Dim CREATE_EXCEL        As String
Dim HEAD_EXCEL          As String

Dim BODY1_EXCEL         As String
Dim BODY2_EXCEL         As String
Dim BODY3_EXCEL         As String

Dim INS1_EXCEL          As String
Dim INS2_EXCEL          As String
Dim INS3_EXCEL          As String


Dim TOTAL_EXCEL         As String

Dim FOOT_EXCEL          As String
Dim DSP_EXCEL           As String
Dim S_END               As String

Dim S_TITLE             As String


Dim ED_HIN_GAI          As String * 20
Dim ED_I                As Integer


'2011.01.11
    
    
'Debug.Print "in Estimate_Proc=" & Format(Now, "hh:mm:ss")
    
    
    Estimate_Proc = True
    
'stTime = Format(Now, "hh:mm:ss")
    
'    Call Input_Lock
    
    
S_TITLE = "自動計算OFF"
    
S_Start = Right(Format(Now, "hh:mm:ss"), 5)
    
    Set excelApplication = CreateObject("Excel.Application")
    

    If Trim(EXCEL_TEMPLATE) = "" Then
        Set excelWorkBook = excelApplication.Workbooks.Add
    
    Else
        DoEvents        '2018.03.16
                                                        
                                                        'ﾃﾝﾌﾟﾚｰﾄﾌﾞｯｸを開く
        Set excelWorkBook = excelApplication.Workbooks.Open(EXCEL_TEMPLATE)
    
    
        DoEvents        '2018.03.16
    
    End If

    Set excelSheet = excelWorkBook.Worksheets(1)
    
    
    
    
    
'excelApplication.Visible = True
    
excelApplication.Calculation = xlCalculationManual
excelApplication.ScreenUpdating = False

    
    
    
CREATE_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents
    
    
    excelSheet.Application.ActiveWindow.DisplayGridlines = False
    
'---    ヘッダー出力
HEAD_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents
    
    Call Estimate_Head_Proc(excelApplication, excelWorkBook, excelSheet)
    
    
    
'---    11行目
    excelSheet.Application.Rows(11).RowHeight = 13.5
    
    
'---    12行目
    Call Estimate_Line12_13_Proc(excelApplication, excelWorkBook, excelSheet)   '2011.01.11
    

'---    資材分出力
BODY1_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents

    If Estimate_SHIZAI_Proc(excelApplication, excelWorkBook, excelSheet, row) Then
'        Call Input_UnLock
        Exit Function
    End If
    SHIZAI_TOTAL_ROW = row

'---    同梱分出力
BODY2_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents

    If Estimate_DOUKON_Proc(excelApplication, excelWorkBook, excelSheet, row) Then
'        Call Input_UnLock
        Exit Function
    End If

'---    付加分出力

BODY3_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents
    If Estimate_FUKA_Proc(excelApplication, excelWorkBook, excelSheet, row, Fsw) Then
'        Call Input_UnLock
        Exit Function
    End If

    
    If Fsw Then                     '2018.01.05
        FUKA_TOTAL_ROW = row
    Else                            '2018.01.05
        FUKA_TOTAL_ROW = 0          '2018.01.05
    End If                          '2018.01.05
    
'---    42行目
    row = row + 2
    excelSheet.Application.Cells(row, 2).Font.Size = 10
    
    excelSheet.Application.Cells(row, 2).HorizontalAlignment = xlLeft
    excelSheet.Application.Cells(row, 2).Value = "【作業費】"
    
    
    
'---    43行目
    row = row + 1
    excelSheet.Application.Rows(row).RowHeight = 20.25
    
'>>>>>>>>>>>>>>>>>>>>>  2017.11.06
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 3)).HorizontalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 3)).VerticalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 3)).MergeCells = True
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row, 5)).HorizontalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row, 5)).VerticalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row, 5)).MergeCells = True
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 6), excelSheet.Application.Cells(row, 7)).HorizontalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 6), excelSheet.Application.Cells(row, 7)).VerticalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 6), excelSheet.Application.Cells(row, 7)).MergeCells = True
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 6), excelSheet.Application.Cells(row, 7)).ShrinkToFit = True
'
'    excelSheet.Application.Cells(row, 8).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row, 8).VerticalAlignment = xlCenter
'
'    excelSheet.Application.Cells(row, 9).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row, 9).VerticalAlignment = xlCenter
'>>>>>>>>>>>>>>>>>>>>>  2017.11.06



    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).MergeCells = True

'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 9)).Font.Size = 10        '2017.11.06
    
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Font.Size = 12

'>>>>>>>>>>>>>>>>>>>>>  2017.11.06
'    excelSheet.Application.Cells(row, 2).Value = "前後工程(秒)"
'    excelSheet.Application.Cells(row, 4).Value = "実作業工程(秒)"
'
'    excelSheet.Application.Cells(row, 6).Value = "作業時間計(秒/個)"
'    excelSheet.Application.Cells(row, 8).Value = "分/個"
'    excelSheet.Application.Cells(row, 9).Value = "分レート"
'>>>>>>>>>>>>>>>>>>>>>  2017.11.06
    excelSheet.Application.Cells(row, 10).Value = "�B工料単価"







'2010.05.13
INS1_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents
    
'>>>>>>>>>>>>>>>>>>>>>  2017.11.06
'    excelSheet.Application.Cells(row, 14).Font.Size = 12
'    excelSheet.Application.Cells(row, 14).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row, 14).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row, 14).Value = "単価"
'
'    excelSheet.Application.Cells(row, 15).Font.Size = 12
'    excelSheet.Application.Cells(row, 15).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row, 15).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row, 15).Value = "チェック"
'
'    excelSheet.Application.Cells(row, 17).Font.Size = 12
'    excelSheet.Application.Cells(row, 17).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row, 17).Value = "ビス・取説・保証書チェック"
'>>>>>>>>>>>>>>>>>>>>>  2017.11.06

'2010.05.13





'---    44行目
    row = row + 1
    excelSheet.Application.Rows(row).RowHeight = 20.25
    
'>>>>>>>>>>>>>>>>>>>>>  2017.11.06
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 3)).HorizontalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 3)).VerticalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 3)).MergeCells = True
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row, 5)).HorizontalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row, 5)).VerticalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row, 5)).MergeCells = True
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 6), excelSheet.Application.Cells(row, 7)).HorizontalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 6), excelSheet.Application.Cells(row, 7)).VerticalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 6), excelSheet.Application.Cells(row, 7)).MergeCells = True
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 6), excelSheet.Application.Cells(row, 7)).ShrinkToFit = True
'
'    excelSheet.Application.Cells(row, 8).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row, 8).VerticalAlignment = xlCenter
'
'    excelSheet.Application.Cells(row, 9).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row, 9).VerticalAlignment = xlCenter
'>>>>>>>>>>>>>>>>>>>>>  2017.11.06



    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).MergeCells = True

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 9)).Font.Size = 10
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Font.Size = 12




    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 10)).Font.Size = 12
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 11), excelSheet.Application.Cells(row, 12)).Font.Size = 14
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 12)).NumberFormatLocal = "#,##0_ "

    
'2009.07.01
        
    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
        
            For i = 0 To 9
            
                Call UniCode_Conv(ITEMREC.BEF_KOUTEI(i).BEF_KOUTEI, "000000000")
                Call UniCode_Conv(ITEMREC.MAIN_KOUTEI(i).MAIN_KOUTEI, "000000000")
                Call UniCode_Conv(ITEMREC.AFT_KOUTEI(i).AFT_KOUTEI, "000000000")
            
            Next i
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    
    End Select
    
    
    
    wkint = 0
    For i = 0 To 8
        If IsNumeric(StrConv(ITEMREC.BEF_KOUTEI(i).BEF_KOUTEI, vbUnicode)) Then
            wkint = wkint + CInt(StrConv(ITEMREC.BEF_KOUTEI(i).BEF_KOUTEI, vbUnicode))
        End If
    Next i
    
    
    wkint = ToHalfAdjust(CCur(wkint * CDbl(YOYU_RITU(0).Caption)), 0)
    
    
    '2009.09.17 PLUS工数加算
'    If IsNumeric(StrConv(ITEMREC.PLUS_KOUSU, vbUnicode)) Then
'        wkint = wkint + CInt(StrConv(ITEMREC.PLUS_KOUSU, vbUnicode))
'    End If
    
    '(秒／個)
    If IsNumeric(Text1(ptxBEF_SEI_LOT).Text) Then
        
        If CInt(Text1(ptxBEF_SEI_LOT).Text) = 0 Then        '2010.02.22
            BEF_KOTEI = 0
        Else
            BEF_KOTEI = ToHalfAdjust(CCur(wkint) / CInt(Text1(ptxBEF_SEI_LOT).Text), 0)
        End If
    Else
        BEF_KOTEI = 0
    End If
    
    wkint = 0
    For i = 0 To 2
        If IsNumeric(StrConv(ITEMREC.AFT_KOUTEI(i).AFT_KOUTEI, vbUnicode)) Then
            wkint = wkint + CInt(StrConv(ITEMREC.AFT_KOUTEI(i).AFT_KOUTEI, vbUnicode))
        End If
    Next i
    wkint = ToHalfAdjust(CCur(wkint * CDbl(YOYU_RITU(0).Caption)), 0)
    '(秒／個)
    If IsNumeric(Text1(ptxBEF_SEI_LOT).Text) Then
        If CInt(Text1(ptxBEF_SEI_LOT).Text) = 0 Then        '2010.02.22
            AFT_KOTEI = 0
        Else
            AFT_KOTEI = ToHalfAdjust(CCur(wkint) / CInt(Text1(ptxBEF_SEI_LOT).Text), 0)
        End If
    Else
        AFT_KOTEI = 0
    End If
    
    
    
    
'    excelSheet.Application.Cells(row, 2).Value = CDbl(Text1(ptxBEF_KOUTEI_KEI2).Text) + CDbl(Text1(ptxAFT_KOUTEI_KEI2).Text)
    
    
    '2009.10.06 IF追加
    If IsNumeric(StrConv(ITEMREC.PLUS_KOUSU, vbUnicode)) Then
        excelSheet.Application.Cells(row, 2).Value = BEF_KOTEI + AFT_KOTEI + CInt(StrConv(ITEMREC.PLUS_KOUSU, vbUnicode))
    Else
        excelSheet.Application.Cells(row, 2).Value = BEF_KOTEI + AFT_KOTEI
    End If
    
    
'    excelSheet.Application.Cells(row, 2).Value = BEF_KOTEI + AFT_KOTEI
    
    
    
    MAIN_KOTEI = 0
    For i = 0 To 4
        If IsNumeric(StrConv(ITEMREC.MAIN_KOUTEI(i).MAIN_KOUTEI, vbUnicode)) Then
            MAIN_KOTEI = MAIN_KOTEI + CInt(StrConv(ITEMREC.MAIN_KOUTEI(i).MAIN_KOUTEI, vbUnicode))
        End If
    Next i
    MAIN_KOTEI = ToHalfAdjust(CCur(MAIN_KOTEI * CDbl(YOYU_RITU(0).Caption)), 0)
    
'    excelSheet.Application.Cells(row, 4).Value = CDbl(Text1(ptxMAIN_KOUTEI_R_RATE).Text)
    excelSheet.Application.Cells(row, 4).Value = MAIN_KOTEI
    
    
    
    
''    excelSheet.Application.Cells(row, 7).Value = CDbl(Text1(ptxKOUTEI_KEI3).Text)
''    excelSheet.Application.Cells(row, 9).Value = CDbl(Text1(ptxAFT_SEI_RATE).Text)
''    excelSheet.Application.Cells(row, 10).Value = CDbl(Text1(ptxAFT_S_KOUSU_BAIKA).Text)
    excelSheet.Application.Cells(row, 6).FormulaR1C1 = "=sum(RC[-5]:RC[-1]"
    
    
    If IsNumeric(Text1(ptxBEF_S_KOUSU).Text) Then
        excelSheet.Application.Cells(row, 8).Value = Text1(ptxBEF_S_KOUSU).Text
    Else
        excelSheet.Application.Cells(row, 8).Value = ""
    End If
    
'2009.07.01    excelSheet.Application.Cells(row, 10).Value = CDbl(Text1(ptxAFT_SEI_RATE).Text)
    
    If IsNumeric(Text1(ptxBEF_SEI_RATE).Text) Then
        excelSheet.Application.Cells(row, 9).Value = CDbl(Text1(ptxBEF_SEI_RATE).Text)
    Else
        excelSheet.Application.Cells(row, 9).Value = ""
    End If
    
    
    excelSheet.Application.Cells(row, 10).FormulaR1C1 = "=round(RC[-2]*RC[-1],2)"
    
    
    If IsNumeric(excelSheet.Application.Cells(row, 10).Value) Then
        wkNum1 = CCur(excelSheet.Application.Cells(row, 10).Value)
    Else
        wkNum1 = 0
    End If
    
    
    If IsNumeric(Text1(ptxBEF_S_KOUSU_BAIKA).Text) Then
        wkNum2 = CCur(Text1(ptxBEF_S_KOUSU_BAIKA).Text)
    Else
        wkNum2 = 0
    End If
    
'>>>>>>>>>  2017.11.06
'    If wkNum1 <> wkNum2 Then
'        MsgBox "�B工料単価が計算値(分/個×分レート)と異なります。"
'        excelSheet.Application.Cells(row, 13).Value = "�B工料単価が計算値(分/個×分レート)と異なります。"
'    End If
'>>>>>>>>>  2017.11.06
    
    
    If IsNumeric(Text1(ptxBEF_S_KOUSU_BAIKA).Text) Then
        excelSheet.Application.Cells(row, 10).Value = CDbl(Text1(ptxBEF_S_KOUSU_BAIKA).Text)
        excelSheet.Application.Cells(row, 10).NumberFormatLocal = "#,##0.00_ "
    Else
        excelSheet.Application.Cells(row, 10).Value = ""
    
    End If
'>>>>>>>>>  2017.11.06
    excelSheet.Application.Cells(row, 2).Value = ""
    excelSheet.Application.Cells(row, 4).Value = ""
    excelSheet.Application.Cells(row, 6).Value = ""
    excelSheet.Application.Cells(row, 8).Value = ""
    excelSheet.Application.Cells(row, 9).Value = ""


    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic

    excelSheet.Application.Cells(row, 9).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(row, 9).Value = "作業費一式"

'>>>>>>>>>  2017.11.06



'2010.05.13
INS2_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents
    
    
'>>>>>>>>>  2017.11.06
'    excelSheet.Application.Cells(row, 14).HorizontalAlignment = xlRight
'    excelSheet.Application.Cells(row, 14).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row, 14).NumberFormatLocal = "#,##0.00_ "
'    excelSheet.Application.Cells(row, 14).FormulaR1C1 = "=roundup(roundup((RC[-12]+RC[-10])/60,1)*RC[-5],2)"


'   excelSheet.Application.Cells(row, 15).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row, 15).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row, 15).FormulaR1C1 = "=IF(RC[-5]=RC[-1],""○"",""×"")"


'    excelSheet.Application.Cells(row, 17).HorizontalAlignment = xlRight
'    excelSheet.Application.Cells(row, 17).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row, 17).NumberFormatLocal = "#,##0.00_ "
'    excelSheet.Application.Cells(row, 17).FormulaR1C1 = "=roundup(RC[-11]/60,1)"

'    excelSheet.Application.Cells(row, 18).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row, 18).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row, 18).FormulaR1C1 = "=IF(RC[-10]=RC[-1],""○"",""×"")"

'2010.05.13
'>>>>>>>>>  2017.11.06












'>>>>>>>>>>>>>>>>>> 2017.10.30
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlDiagonalUp).LineStyle = xlNone

    
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlEdgeLeft).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlEdgeBottom).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlEdgeRight).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlInsideVertical).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlInsideVertical).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlInsideVertical).ColorIndex = xlAutomatic

'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlInsideHorizontal).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlInsideHorizontal).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 2), excelSheet.Application.Cells(row, 9)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic


'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone

'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).Weight = xlThick
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).Weight = xlThick
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).Weight = xlThick
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).Weight = xlThick
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).LineStyle = xlNone
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).Weight = xlThick
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
'>>>>>>>>>>>>>>>>>> 2017.10.30








'2010.05.13
INS3_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents
    
'>>>>>>>>>  2017.11.06
    
'    excelSheet.Application.Cells(row + 1, 14).Font.Size = 12
'    excelSheet.Application.Cells(row + 1, 14).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row + 1, 14).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row + 1, 14).Value = "単価"

'    excelSheet.Application.Cells(row + 1, 15).Font.Size = 12
'    excelSheet.Application.Cells(row + 1, 15).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row + 1, 15).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row + 1, 15).Value = "チェック"
'>>>>>>>>>  2017.11.06
'2010.05.13



'---    46行目
TOTAL_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents
    
    
    row = row + 2
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row, 10)).HorizontalAlignment = xlCenter
    
    excelSheet.Application.Cells(row, 9).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(row, 9).VerticalAlignment = xlCenter
    excelSheet.Application.Cells(row, 9).Font.Size = 14
    excelSheet.Application.Cells(row, 9).Value = "商品化費用�@＋�A＋�B"

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Font.Size = 14
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Font.FontStyle = "太字"
        
'>>>>>>>>>  2017.11.06
'    If SHIZAI_TOTAL_ROW = 15 Then
'        excelSheet.Application.Cells(row, 10).FormulaR1C1 = "=R[-2]C+R[" & FUKA_TOTAL_ROW - row & "]C"
'    Else
'        excelSheet.Application.Cells(row, 10).FormulaR1C1 = "=R[-2]C+R[" & SHIZAI_TOTAL_ROW - row & "]C+R[" & FUKA_TOTAL_ROW - row & "]C"
'    End If
    
    
    If SHIZAI_TOTAL_ROW = 15 Then
        excelSheet.Application.Cells(row, 10).FormulaR1C1 = "=R[-2]C+R[" & FUKA_TOTAL_ROW - 1 - row & "]C"
    Else
'        excelSheet.Application.Cells(row, 10).FormulaR1C1 = "=R[-2]C+R[" & SHIZAI_TOTAL_ROW - row & "]C+R[" & FUKA_TOTAL_ROW - 1 - row & "]C"      '2018.01.05
        
        If FUKA_TOTAL_ROW = 0 Then                                                                                                                  '2018.01.05
            excelSheet.Application.Cells(row, 10).FormulaR1C1 = "=R[-2]C+R[" & SHIZAI_TOTAL_ROW - row & "]C"                                        '2018.01.05
        Else                                                                                                                                        '2018.01.05
            excelSheet.Application.Cells(row, 10).FormulaR1C1 = "=R[-2]C+R[" & SHIZAI_TOTAL_ROW - row & "]C+R[" & FUKA_TOTAL_ROW - row & "]C"       '2018.01.05
        End If                                                                                                                                      '2018.01.05
    End If
'>>>>>>>>>  2017.11.06
    
    
    
    excelSheet.Application.Cells(row, 10).NumberFormatLocal = "#,##0.00_ "
    
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).LineStyle = xlNone

    excelSheet.Application.Cells(row, 10).Borders(xlLeft).LineStyle = xlContinuous
    excelSheet.Application.Cells(row, 10).Borders(xlLeft).Weight = xlThick
    excelSheet.Application.Cells(row, 10).Borders(xlLeft).ColorIndex = xlAutomatic



'>>>>>>>>>  2017.11.06
'2010.05.13
'    excelSheet.Application.Cells(row, 14).HorizontalAlignment = xlRight
'    excelSheet.Application.Cells(row, 14).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row, 14).NumberFormatLocal = "#,##0.00_ "
'    If SHIZAI_TOTAL_ROW = 15 Then
'        excelSheet.Application.Cells(row, 14).FormulaR1C1 = "=R[-2]C+R[" & FUKA_TOTAL_ROW - row & "]C"
'    Else
'        excelSheet.Application.Cells(row, 14).FormulaR1C1 = "=R[-2]C+R[" & SHIZAI_TOTAL_ROW - row & "]C+R[" & FUKA_TOTAL_ROW - row & "]C"
'    End If
'>>>>>>>>>  2017.11.06


 '>>>>>>>>>  2017.11.06
'    excelSheet.Application.Cells(row, 15).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row, 15).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row, 15).FormulaR1C1 = "=IF(RC[-5]=RC[-1],""○"",""×"")"


'    excelSheet.Application.Cells(row + 1, 17).Font.Size = 11
'    excelSheet.Application.Cells(row + 1, 17).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row + 1, 17).Value = "ロット数"

'    excelSheet.Application.Cells(row + 2, 17).Font.Size = 11
'    excelSheet.Application.Cells(row + 2, 17).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row + 2, 17).Value = Text1(ptxBEF_SEI_LOT).Text
 '>>>>>>>>>  2017.11.06

'2010.05.13



'---    48行目
    row = row + 2
    excelSheet.Application.Cells(row, 2).Font.Size = 10
    
    excelSheet.Application.Cells(row, 2).HorizontalAlignment = xlLeft
    excelSheet.Application.Cells(row, 2).Value = "【備考】"


'---    49〜51行目
    
    
    row = row + 1
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).Borders(xlDiagonalUp).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).Borders(xlInsideVertical).LineStyle = xlNone
    
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).Borders(xlInsideHorizontal).LineStyle = xlNone


    '-----------------  セルを結合  2013.09.30
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).HorizontalAlignment = xlLeft
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).VerticalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).MergeCells = True
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(row + 1, 2), excelSheet.Application.Cells(row + 1, 11)).HorizontalAlignment = xlLeft
'    excelSheet.Application.Range(excelSheet.Application.Cells(row + 1, 2), excelSheet.Application.Cells(row + 1, 11)).VerticalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row + 1, 2), excelSheet.Application.Cells(row + 1, 11)).MergeCells = True

'    excelSheet.Application.Range(excelSheet.Application.Cells(row + 2, 2), excelSheet.Application.Cells(row + 2, 11)).HorizontalAlignment = xlLeft
'    excelSheet.Application.Range(excelSheet.Application.Cells(row + 2, 2), excelSheet.Application.Cells(row + 2, 11)).VerticalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row + 2, 2), excelSheet.Application.Cells(row + 2, 11)).MergeCells = True
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).HorizontalAlignment = xlLeft
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).VerticalAlignment = xlTop
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).WrapText = True
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 2, 11)).MergeCells = True
    
    
    
    '-----------------  セルを結合



    If Trim(RichTextBox1(prchM_BIKOU).Text) = "" Then
    Else
        wkBikou = Split(RTrim(RichTextBox1(prchM_BIKOU).Text), vbCrLf, -1)
        For i = row To UBound(wkBikou) + row
            excelSheet.Application.Cells(i, 2).Value = wkBikou(i - row)
        Next i
    End If
    
    
    
    



'---    53〜56行目
FOOT_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents
    
    
    row = row + 5
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 1, 3)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 1, 3)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 1, 3)).MergeCells = True
    
    Select Case Trim(Text1(ptxMITSUMORI_KBN).Text)
        Case "1"
            excelSheet.Application.Cells(row, 2).Value = "新規仕様"
        Case "2"
            excelSheet.Application.Cells(row, 2).Value = "現行仕様"
    End Select

    excelSheet.Application.Range(excelSheet.Application.Cells(row + 2, 2), excelSheet.Application.Cells(row + 3, 3)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row + 2, 2), excelSheet.Application.Cells(row + 3, 3)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row + 2, 2), excelSheet.Application.Cells(row + 3, 3)).MergeCells = True

    excelSheet.Application.Range(excelSheet.Application.Cells(row + 2, 2), excelSheet.Application.Cells(row + 3, 3)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row + 2, 2), excelSheet.Application.Cells(row + 3, 3)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row + 2, 2), excelSheet.Application.Cells(row + 3, 3)).MergeCells = True

   

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).WrapText = True
    
    excelSheet.Application.Cells(row, 4).Value = "仕様書��" & Left(Combo1(pcmbSHIMUKE).Text, Len(Combo1(pcmbSHIMUKE).Text) - 4)
    
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row + 3, 5)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row + 3, 5)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row + 3, 5)).MergeCells = True
    
    
    excelSheet.Application.Cells(row, 5).Value = Trim(Text1(ptxSHIYOU_NO).Text)







    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 3, 3)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 3, 3)).Borders(xlDiagonalUp).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 3, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 3, 3)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 3, 3)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 3, 3)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 3, 3)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 3, 3)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 3, 3)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 3, 3)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 3, 3)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 3, 3)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 3, 3)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 3, 3)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 3, 3)).Borders(xlInsideVertical).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 3, 3)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 3, 3)).Borders(xlInsideHorizontal).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row + 3, 3)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic


    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).Borders(xlDiagonalUp).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).Borders(xlInsideVertical).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 4), excelSheet.Application.Cells(row + 3, 4)).Borders(xlInsideHorizontal).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row + 3, 5)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row + 3, 5)).Borders(xlDiagonalUp).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row + 3, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row + 3, 5)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row + 3, 5)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row + 3, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row + 3, 5)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row + 3, 5)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row + 3, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row + 3, 5)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row + 3, 5)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row + 3, 5)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row + 3, 5)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row + 3, 5)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row + 3, 5)).Borders(xlInsideVertical).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row + 3, 5)).Borders(xlInsideHorizontal).LineStyle = xlNone








'''2011.01.21
    If Trim(Insert_Pic) = "" Then
        excelSheet.Application.Range(excelSheet.Application.Cells(row + 1, 9), excelSheet.Application.Cells(row + 3, 9)).HorizontalAlignment = xlCenter
        excelSheet.Application.Range(excelSheet.Application.Cells(row + 1, 9), excelSheet.Application.Cells(row + 3, 9)).VerticalAlignment = xlCenter
        excelSheet.Application.Range(excelSheet.Application.Cells(row + 1, 9), excelSheet.Application.Cells(row + 3, 9)).MergeCells = True
    
        excelSheet.Application.Range(excelSheet.Application.Cells(row + 1, 10), excelSheet.Application.Cells(row + 3, 10)).HorizontalAlignment = xlCenter
        excelSheet.Application.Range(excelSheet.Application.Cells(row + 1, 10), excelSheet.Application.Cells(row + 3, 10)).VerticalAlignment = xlCenter
        excelSheet.Application.Range(excelSheet.Application.Cells(row + 1, 10), excelSheet.Application.Cells(row + 3, 10)).MergeCells = True
    
        excelSheet.Application.Range(excelSheet.Application.Cells(row + 1, 11), excelSheet.Application.Cells(row + 3, 11)).HorizontalAlignment = xlCenter
        excelSheet.Application.Range(excelSheet.Application.Cells(row + 1, 11), excelSheet.Application.Cells(row + 3, 11)).VerticalAlignment = xlCenter
        excelSheet.Application.Range(excelSheet.Application.Cells(row + 1, 11), excelSheet.Application.Cells(row + 3, 11)).MergeCells = True
    
    
    
    
        excelSheet.Application.Cells(row, 9).HorizontalAlignment = xlCenter
        excelSheet.Application.Cells(row, 9).VerticalAlignment = xlCenter
        excelSheet.Application.Cells(row, 9).Font.Size = 10
        excelSheet.Application.Cells(row, 9).Value = "承認印"
    
        excelSheet.Application.Cells(row, 10).HorizontalAlignment = xlCenter
        excelSheet.Application.Cells(row, 10).VerticalAlignment = xlCenter
        excelSheet.Application.Cells(row, 10).Font.Size = 10
        excelSheet.Application.Cells(row, 10).Value = "検印"
    
        excelSheet.Application.Cells(row, 11).HorizontalAlignment = xlCenter
        excelSheet.Application.Cells(row, 11).VerticalAlignment = xlCenter
        excelSheet.Application.Cells(row, 11).Font.Size = 10
        excelSheet.Application.Cells(row, 11).Value = "担当印"
    
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
    
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlEdgeLeft).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlEdgeTop).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlEdgeBottom).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlEdgeRight).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlInsideVertical).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
    
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlInsideHorizontal).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row + 3, 11)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
    End If
'''2011.01.21
    If Trim(Insert_Pic) <> "" Then
        
        
        


        
        
        
        
'        excelSheet.Application.Pictures.Insert (Insert_Pic)


'        excelSheet.Pictures.Insert(Insert_Pic).Top = excelSheet.Application.Cells(row, 7).Top
'        excelSheet.Pictures.Insert(Insert_Pic).Left = excelSheet.Application.Cells(row, 7).Left
        
        
'------------------ 2013.07.02
'         With excelSheet.Pictures.Insert(Insert_Pic)
'            .Top = excelSheet.Application.Cells(row - 1, 7).Top
'            .Left = excelSheet.Application.Cells(row - 1, 7).Left
'''            .Height = 3.15 / 0.0378
'            .Width = (excelSheet.Application.Cells(row - 1, 7).Width + _
'                        excelSheet.Application.Cells(row - 1, 8).Width + _
'                        excelSheet.Application.Cells(row - 1, 9).Width + _
'                        excelSheet.Application.Cells(row - 1, 10).Width + _
'                        excelSheet.Application.Cells(row - 1, 11).Width)
''            .Width = (excelSheet.Application.Cells(row - 1, 11).Top + excelSheet.Application.Cells(row - 1, 11).Width)
'
'
'
'
''            .Height = 2.93 / 0.0378
'
'
''            .Width = 8.62 / 0.0378
'
'
'
'        End With




        excelSheet.Shapes.AddPicture(Insert_Pic, _
                                            False, _
                                            True, _
                                            excelSheet.Application.Cells(row - 1, 7).Left, _
                                            excelSheet.Application.Cells(row - 1, 7).Top, _
                                            (excelSheet.Application.Cells(row - 1, 7).Width + _
                                            excelSheet.Application.Cells(row - 1, 8).Width + _
                                            excelSheet.Application.Cells(row - 1, 9).Width + _
                                            excelSheet.Application.Cells(row - 1, 10).Width + _
                                            excelSheet.Application.Cells(row - 1, 11).Width), _
                                            100).Apply

'------------------





'        With excelSheet.Shapes(8)
'            .LockAspectRatio = True     '---(1)図形の縦横の比率を固定
'        End With


    End If



'---    大外枠
    row = row + 4

    excelSheet.Application.Rows(row).RowHeight = 45     '2011.01.24
    
    excelSheet.Application.Range(excelSheet.Application.Cells(11, 1), excelSheet.Application.Cells(11, 12)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(11, 1), excelSheet.Application.Cells(11, 12)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(11, 1), excelSheet.Application.Cells(11, 12)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(11, 1), excelSheet.Application.Cells(row, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(11, 1), excelSheet.Application.Cells(row, 1)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(11, 1), excelSheet.Application.Cells(row, 1)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 1), excelSheet.Application.Cells(row, 12)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 1), excelSheet.Application.Cells(row, 12)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 1), excelSheet.Application.Cells(row, 12)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(11, 12), excelSheet.Application.Cells(row, 12)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(11, 12), excelSheet.Application.Cells(row, 12)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(11, 12), excelSheet.Application.Cells(row, 12)).Borders(xlEdgeRight).ColorIndex = xlAutomatic


    excelSheet.Application.Cells(1, 1).Select



excelApplication.Calculation = xlCalculationAutomatic



DSP_EXCEL = Right(Format(Now, "hh:mm:ss"), 5)
DoEvents


    excelApplication.Calculation = xlCalculationAutomatic
    
    
    excelApplication.ScreenUpdating = True
'    excelApplication.Visible = True
    
    
    excelApplication.displayalerts = False
    
    
ED_HIN_GAI = Text1(ptxHin_Gai).Text
    
If Right(RTrim(ED_HIN_GAI), 1) = "." Then
'    Right(RTrim(ED_HIN_GAI), 1) = "_"

    For ED_I = 20 To 0 Step -1
        If Mid(ED_HIN_GAI, ED_I, 1) = "." Then
            Mid(ED_HIN_GAI, ED_I, 1) = "_"
            Exit For
        End If
    Next ED_I
    

End If
    
    
    
    excelWorkBook.saveas FileName:=(Save_Dir & Trim(ED_HIN_GAI)), FileFormat:=xlOpenXMLWorkbook




    
    





    Set excelSheet = Nothing
    
    Set excelWorkBook = Nothing
    
    excelApplication.quit
    
    Set excelApplication = Nothing

    
S_END = Right(Format(Now, "hh:mm:ss"), 5)
    
    
'hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
'    "S=" & S_Start & _
'    " S.CRE=" & CREATE_EXCEL & _
'    " S.BODY1=" & BODY1_EXCEL & _
'    " S.BODY2=" & BODY2_EXCEL & _
'    " S.BODY3=" & BODY3_EXCEL & _
'    " S.INS1=" & INS1_EXCEL & _
'    " S.INS2=" & INS2_EXCEL & _
'    " S.INS3=" & INS3_EXCEL & _
'    " S.TOTAL=" & TOTAL_EXCEL & _
'    " S.FOOT=" & FOOT_EXCEL & _
'    " S.VISIBLE=" & DSP_EXCEL & _
'    " E=" & S_END, Me.hwnd, 0)
    
    
    
'Call LOG_OUT(LOG_F, S_TITLE & "Hin=" & StrConv(ITEMREC.HIN_GAI, vbUnicode) & "S=" & S_Start & _
'    " S.CRE=" & CREATE_EXCEL & _
'    " S.BODY1=" & BODY1_EXCEL & _
'    " S.BODY2=" & BODY2_EXCEL & _
'    " S.BODY3=" & BODY3_EXCEL & _
'    " S.INS1=" & INS1_EXCEL & _
'    " S.INS2=" & INS2_EXCEL & _
'    " S.INS3=" & INS3_EXCEL & _
'    " S.TOTAL=" & TOTAL_EXCEL & _
'    " S.FOOT=" & FOOT_EXCEL & _
'    " S.VISIBLE=" & DSP_EXCEL & _
'    " E=" & S_END)
    
'    Call Input_UnLock
    
Debug.Print "out Estimate_Proc=" & Format(Now, "hh:mm:ss")
    
    
    
    Estimate_Proc = False
End Function

Private Function Detail_Disp_Proc(Errflg As Integer) As Integer
'----------------------------------------------------------------------------
'                   現在値画面表示
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer
Dim wkint       As Integer
Dim wkDouble    As Double

Dim wkKUSATU    As Variant
Dim c           As String * 128

Dim wkBikou     As String

Dim INV_F   As Boolean


    Detail_Disp_Proc = True
    
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
'            MsgBox "入力した項目はエラーです。(品番)"
            Errflg = True
            Detail_Disp_Proc = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function

    End Select
    
    
    
    For i = 2 To 5
        Command1(i).Enabled = True
    Next i
    
    
    '品名
    Text1(ptxHin_Name).Text = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
    '標準棚番
    Text1(ptxST_SOKO).Text = StrConv(ITEMREC.ST_SOKO, vbUnicode)
    Text1(ptxST_RETU).Text = StrConv(ITEMREC.ST_RETU, vbUnicode)
    Text1(ptxST_REN).Text = StrConv(ITEMREC.ST_REN, vbUnicode)
    Text1(ptxST_DAN).Text = StrConv(ITEMREC.ST_DAN, vbUnicode)
    
    
    
    
    
    
    
    
    '-----------------------------------------------------------    2009.06.02 ▽
    '見積書備考
    wkBikou = Replace(StrConv(ITEMREC.M_BIKOU, vbUnicode), Chr(0), " ")
    RichTextBox1(prchM_BIKOU).Text = RTrim(wkBikou)
    
    '仕様書��
    Text1(ptxSHIYOU_NO).Text = RTrim(StrConv(ITEMREC.SHIYOU_NO, vbUnicode))
    
    '見積区分
    Text1(ptxMITSUMORI_KBN).Text = RTrim(StrConv(ITEMREC.MITSUMORI_KBN, vbUnicode))
    '単価切替日
    Text1(ptxTANKA_KIRIKAE_DT).Text = RTrim(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode))
    '切替区分
    Text1(ptxKIRIKAE_KBN).Text = RTrim(StrConv(ITEMREC.KIRIKAE_KBN, vbUnicode))

    '-----------------------------------------------------------    2009.06.02 △
    
    
    
    
    
    '-----------------------------------    旧単価  2009.07.24
    
    
    '(売価)商品化工料
    If IsNumeric(StrConv(ITEMREC.BEF_S_KOUSU_BAIKA, vbUnicode)) Then
        Text1(ptxOLD_S_KOUSU_BAIKA).Text = Format(StrConv(ITEMREC.BEF_S_KOUSU_BAIKA, vbUnicode), "#0.00")
    Else
        Text1(ptxOLD_S_KOUSU_BAIKA).Text = "0.00"
    End If
    
    '(売価)商品化工料
    If IsNumeric(StrConv(ITEMREC.BEF_S_SHIZAI_BAIKA, vbUnicode)) Then
        Text1(ptxOLD_S_SHIZAI_BAIKA).Text = Format(StrConv(ITEMREC.BEF_S_SHIZAI_BAIKA, vbUnicode), "#0.00")
    Else
        Text1(ptxOLD_S_SHIZAI_BAIKA).Text = "0.00"
    End If
    
    '外装単価
    If IsNumeric(StrConv(ITEMREC.BEF_S_GAISO_TANKA, vbUnicode)) Then
        Text1(ptxOLD_S_GAISO_TANKA).Text = Format(StrConv(ITEMREC.BEF_S_GAISO_TANKA, vbUnicode), "#0.00")
    Else
        Text1(ptxOLD_S_GAISO_TANKA).Text = "0.00"
    End If
    
    'PPSC加工単価
    If IsNumeric(StrConv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, vbUnicode)) Then
        Text1(ptxOLD_S_PPSC_KAKO_KOSU).Text = Format(StrConv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, vbUnicode), "#0.00")
    Else
        Text1(ptxOLD_S_PPSC_KAKO_KOSU).Text = "0.00"
    End If
    
    'BU加工単価
    If IsNumeric(StrConv(ITEMREC.BEF_S_BU_KAKO_KOSU, vbUnicode)) Then
        Text1(ptxOLD_S_BU_KAKO_KOSU).Text = Format(StrConv(ITEMREC.BEF_S_BU_KAKO_KOSU, vbUnicode), "#0.00")
    Else
        Text1(ptxOLD_S_BU_KAKO_KOSU).Text = "0.00"
    End If
'------2009.07.24
    
    
    
    
    
    
    
    
    '-----------------------------------    旧単価  2009.07.24
    
    
    
    
    '-----------------------------------    変更前
    
    
    
    If IsNumeric(StrConv(ITEMREC.SEI_LOT, vbUnicode)) Then
        Text1(ptxBEF_SEI_LOT).Text = Format(Val(StrConv(ITEMREC.SEI_LOT, vbUnicode)), "#0")
    Else
'        Text1(ptxBEF_SEI_LOT).Text = "1"
        Text1(ptxBEF_SEI_LOT).Text = ""
    End If
    
    
    '分ﾚｰﾄ
    If IsNumeric(StrConv(ITEMREC.SEI_RATE, vbUnicode)) Then
        Text1(ptxBEF_SEI_RATE).Text = Format(Val(StrConv(ITEMREC.SEI_RATE, vbUnicode)), "#0")
    Else
        
'        If IsNumeric(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)) Then
'            Text1(ptxBEF_SEI_RATE).Text = Format(Val(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#0")
'        Else
            Text1(ptxBEF_SEI_RATE).Text = ""
'        End If
    End If
    
    
    
    
    
    '工数
    If IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
        Text1(ptxBEF_S_KOUSU).Text = Format(CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode)), "#0.0")
    Else
'        Text1(ptxBEF_S_KOUSU).Text = "0.0"
        Text1(ptxBEF_S_KOUSU).Text = ""
    End If
    '(原価)工料
    If IsNumeric(StrConv(ITEMREC.S_KOUSU_GENKA, vbUnicode)) Then
        Text1(ptxBEF_S_KOUSU_GENKA).Text = Format(CDbl(StrConv(ITEMREC.S_KOUSU_GENKA, vbUnicode)), "#0.00")
    Else
'        Text1(ptxBEF_S_KOUSU_GENKA).Text = "0.00"
        Text1(ptxBEF_S_KOUSU_GENKA).Text = ""
    End If
    '工料
    If IsNumeric(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) Then
        Text1(ptxBEF_S_KOUSU_BAIKA).Text = Format(CDbl(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)), "#0.00")
    Else
'        Text1(ptxBEF_S_KOUSU_BAIKA).Text = "0.00"
        Text1(ptxBEF_S_KOUSU_BAIKA).Text = ""
    End If
    '(原価)資材
    If IsNumeric(StrConv(ITEMREC.S_SHIZAI_GENKA, vbUnicode)) Then
        Text1(ptxBEF_S_SHIZAI_GENKA).Text = Format(CDbl(StrConv(ITEMREC.S_SHIZAI_GENKA, vbUnicode)), "#0.00")
    Else
'        Text1(ptxBEF_S_SHIZAI_GENKA).Text = "0.00"
        Text1(ptxBEF_S_SHIZAI_GENKA).Text = ""
    End If
    '資材
    If IsNumeric(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)) Then
        Text1(ptxBEF_S_SHIZAI_BAIKA).Text = Format(CDbl(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)), "#0.00")
    Else
'        Text1(ptxBEF_S_SHIZAI_BAIKA).Text = "0.00"
        Text1(ptxBEF_S_SHIZAI_BAIKA).Text = ""
    End If
    
    
    
    '外装費
    If IsNumeric(StrConv(ITEMREC.S_GAISO_TANKA, vbUnicode)) Then
        Text1(ptxBEF_S_GAISO_TANKA).Text = Format(CDbl(StrConv(ITEMREC.S_GAISO_TANKA, vbUnicode)), "#0.00")
    Else
'        Text1(ptxBEF_S_GAISO_TANKA).Text = "0.00"
        Text1(ptxBEF_S_GAISO_TANKA).Text = ""
    End If
    
    
    'PPSC加工単価
    If IsNumeric(StrConv(ITEMREC.S_PPSC_KAKO_KOSU, vbUnicode)) Then
        Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text = Format(CDbl(StrConv(ITEMREC.S_PPSC_KAKO_KOSU, vbUnicode)), "#0.00")
    Else
'        Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text = "0.00"
        Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text = ""
    End If
    'BU加工単価
    If IsNumeric(StrConv(ITEMREC.S_BU_KAKO_KOSU, vbUnicode)) Then
        Text1(ptxBEF_S_BU_KAKO_KOSU).Text = Format(CDbl(StrConv(ITEMREC.S_BU_KAKO_KOSU, vbUnicode)), "#0.00")
    Else
'        Text1(ptxBEF_S_BU_KAKO_KOSU).Text = "0.00"
        Text1(ptxBEF_S_BU_KAKO_KOSU).Text = ""
    End If
    
    
    
    
    
    
    '設定日
    Text1(ptxBEF_S_KOUSU_SET_DATE).Text = Trim(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode))
    '担当者
    Text1(ptxBEF_SEI_TANKA_TANTO).Text = Trim(StrConv(ITEMREC.SEI_TANKA_TANTO, vbUnicode))
    'メモ
    Text1(ptxBEF_SE_TANKA_MEMO).Text = Trim(StrConv(ITEMREC.SE_TANKA_MEMO, vbUnicode))


    '-----------------------------------    変更後
    
    
    If Trim(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode)) = "" Then
        'ﾛｯﾄ数
        Text1(ptxAFT_SEI_LOT).Text = "1"
    Else
        'ﾛｯﾄ数
        If IsNumeric(StrConv(ITEMREC.SEI_LOT, vbUnicode)) Then
            Text1(ptxAFT_SEI_LOT).Text = Format(Val(StrConv(ITEMREC.SEI_LOT, vbUnicode)), "#0")
        Else
            Text1(ptxAFT_SEI_LOT).Text = "1"
        End If
    End If
    
    
    
    '2011.03.29 管理ファイルの分レートを常に使用する
''    If Trim(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode)) = "" Then
''        '分ﾚｰﾄ
''        Text1(ptxAFT_SEI_RATE).Text = Format(Val(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#0")
''    Else
''        '分ﾚｰﾄ
''        If IsNumeric(StrConv(ITEMREC.SEI_RATE, vbUnicode)) Then
''            Text1(ptxAFT_SEI_RATE).Text = Format(Val(StrConv(ITEMREC.SEI_RATE, vbUnicode)), "#0")
''        Else
''
''            If IsNumeric(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)) Then
''                Text1(ptxAFT_SEI_RATE).Text = Format(Val(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#0")
''            Else
''                Text1(ptxAFT_SEI_RATE).Text = ""
''            End If
''        End If
''    End If
    
    Text1(ptxAFT_SEI_RATE).Text = Format(Val(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#0")
    '2011.03.29 管理ファイルの分レートを常に使用する
    
    
    
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
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    wkint = wkint + 1
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
    Text1(ptxBEF_KOUTEI_QTY02).Text = Format(wkint, "#0")
    Text1(ptxBEF_KOUTEI_KOUSU02).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI02).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY02).Text), "#0")
    '�B
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(DOUKON_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = DOUKON_T(j) Then
                    wkint = wkint + 1
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
    Text1(ptxBEF_KOUTEI_QTY03).Text = Format(wkint, "#0")
    Text1(ptxBEF_KOUTEI_KOUSU03).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI03).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY03).Text), "#0")
    '�C
    
    
    If KUSATU_F Then
    
        '草津はINI参照
            
        wkint = 0
        If GetIni(Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1), KUSATU_ETC, App.EXEName, c) Then
        Else
            wkKUSATU = Split(Trim(c), ",", -1)
                    
            If IsNumeric(wkKUSATU(0)) Then
                wkint = CInt(wkKUSATU(0))
            End If
        End If
        If GetIni(Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1), Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), App.EXEName, c) Then
            If GetIni(App.EXEName, Trim(StrConv(StrConv(ITEMREC.HIN_NAME, vbUnicode), vbWide)), App.EXEName, c) Then
            Else
                wkKUSATU = Split(Trim(c), ",", -1)
                        
                If IsNumeric(wkKUSATU(0)) Then
                    wkint = CInt(wkKUSATU(0))
                End If
            End If
        Else
            wkKUSATU = Split(Trim(c), ",", -1)
            If IsNumeric(wkKUSATU(0)) Then
                wkint = CInt(wkKUSATU(0))
            End If
        End If
        Text1(ptxBEF_KOUTEI_TANI04).Text = Format(wkint, "#0")
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
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    wkint = wkint + 1
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
    Text1(ptxBEF_KOUTEI_QTY05).Text = Format(wkint, "#0")
    Text1(ptxBEF_KOUTEI_KOUSU05).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI05).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY05).Text), "#0")
    '�E
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(DOUKON_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = DOUKON_T(j) Then
                    wkint = wkint + 1
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
    Text1(ptxBEF_KOUTEI_QTY06).Text = Format(wkint, "#0")
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
    
    If IsNumeric(StrConv(ITEMREC.PLUS_KOUSU, vbUnicode)) Then
        Text1(ptxPLUS_KOUSU).Text = Format(CInt(StrConv(ITEMREC.PLUS_KOUSU, vbUnicode)), "#0")
    Else
        Text1(ptxPLUS_KOUSU).Text = "0"
    End If
    
    '計
    wkint = 0
    For i = ptxBEF_KOUTEI_KOUSU01 To ptxBEF_KOUTEI_KOUSU09 Step 3
    
        wkint = wkint + CInt(Text1(i).Text)
    
    Next i
    
    Text1(ptxBEF_KOUTEI_KEI1).Text = Format(wkint, "#0")
    
    
    Text1(ptxBEF_KOUTEI_R_RATE).Text = ToHalfAdjust(CCur(wkint * CDbl(YOYU_RITU(0).Caption)), 0)
    
    
        
    
    
    
    '(秒／個)
    If IsNumeric(Text1(ptxBEF_SEI_LOT).Text) Then
        If CInt(Text1(ptxBEF_SEI_LOT).Text) <> 0 Then
            Text1(ptxBEF_KOUTEI_KEI2).Text = Format(ToHalfAdjust(CCur(CInt(Text1(ptxBEF_KOUTEI_R_RATE).Text) / CInt(Text1(ptxBEF_SEI_LOT).Text)), 0), "#0")
        Else
            Text1(ptxBEF_KOUTEI_KEI2).Text = "0"
        End If
    Else
        Text1(ptxBEF_KOUTEI_KEI2).Text = "0"
    End If
    '(分／個)
    If IsNumeric(Text1(ptxPLUS_KOUSU).Text) Then
        Text1(ptxPLUS_KOUSU).Text = Format(CInt(Text1(ptxPLUS_KOUSU).Text), "#0")
    Else
        Text1(ptxPLUS_KOUSU).Text = "0"
    End If
    
    Text1(ptxBEF_KOUTEI_KEI3).Text = Format(ToRoundUp(CCur(CInt(CInt(Text1(ptxBEF_KOUTEI_KEI2).Text) + CInt(Text1(ptxPLUS_KOUSU).Text)) / 60), 1), "#0.0")
    '(円／個)
'    If IsNumeric(Text1(ptxBEF_SEI_RATE).Text) Then
'        Text1(ptxBEF_KOUTEI_KEI4).Text = Format(ToRoundUp(CCur(CDbl(Text1(ptxBEF_KOUTEI_KEI3).Text) * CInt((Text1(ptxBEF_SEI_RATE).Text))), 2), "#0.00")
    
    
    '2009.07.09 BEF-->AFT
    If IsNumeric(Text1(ptxAFT_SEI_RATE).Text) Then
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
    
    
''    Text1(ptxMAIN_KOUTEI_QTY01).Text = "1"
    
    
    If IsNumeric(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode)) Then
        '2009.09.18
        If IsDate(Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 1, 4) & "/" & Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 5, 2) & "/" & Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 7, 4)) Then
            Text1(ptxMAIN_KOUTEI_QTY01).Text = Format(CInt(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode)), "#0")
        Else
            Text1(ptxMAIN_KOUTEI_QTY01).Text = "1"
        End If
    Else
        Text1(ptxMAIN_KOUTEI_QTY01).Text = "1"
    End If
    Text1(ptxMAIN_KOUTEI_KOUSU01).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI01).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY01).Text), "#0")
    
    
    
    
    
    '�A
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                    
                        wkint = wkint + CInt(ToRoundUp(CCur(KOUSEI(i, ColS_KOUSU)) * CDbl(KOUSEI(i, ColKO_QTY)), 0))
                    
                    End If
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI02).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_QTY02).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU02).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI02).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY02).Text), "#0")
    '�B
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(DOUKON_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = DOUKON_T(j) Then
                    
                    
                    If IsNumeric(KOUSEI(i, ColKO_QTY)) Then
                        wkint = wkint + CInt(ToRoundUp(CCur(KOUSEI(i, ColKO_QTY)), 0))
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
    Text1(ptxMAIN_KOUTEI_QTY03).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_KOUSU03).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI03).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY03).Text), "#0")
    '�C
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(KAKOU_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = KAKOU_T(j) Then
                    
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkint = wkint + CInt(KOUSEI(i, ColS_KOUSU))
                    End If
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI04).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_QTY04).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU04).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI04).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY04).Text), "#0")
    '�D
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
            
            
            For j = 0 To UBound(SHIZAI_T)
            
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    If IsNumeric(KOUSEI(i, ColSEI_SYU_KON)) Then
                        wkint = wkint + CInt(KOUSEI(i, ColSEI_SYU_KON))
                    End If
                End If
            
            Next j
            
            
            
            
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI05).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_QTY05).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU05).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI05).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY05).Text), "#0")
    '計
    wkint = 0
    For i = ptxMAIN_KOUTEI_KOUSU01 To ptxMAIN_KOUTEI_KOUSU09 Step 3
    
        If IsNumeric(Text1(i).Text) Then
            wkint = wkint + CInt(Text1(i).Text)
        End If
    Next i
    Text1(ptxMAIN_KOUTEI_KEI1).Text = Format(wkint, "#0")
    
    
    Text1(ptxMAIN_KOUTEI_R_RATE).Text = ToHalfAdjust(CCur(wkint * CDbl(YOYU_RITU(0).Caption)), 0)
    
    
    '(秒／個)
    If IsNumeric(Text1(ptxBEF_SEI_LOT).Text) Then
        Text1(ptxMAIN_KOUTEI_KEI2).Text = Format(ToHalfAdjust(CCur(Text1(ptxMAIN_KOUTEI_R_RATE).Text), 0), "#0")
    Else
        Text1(ptxMAIN_KOUTEI_KEI2).Text = "0"
    End If
    '(分／個)
    Text1(ptxMAIN_KOUTEI_KEI3).Text = Format(ToRoundUp(CCur(CInt(Text1(ptxMAIN_KOUTEI_KEI2).Text) / 60), 1), "#0.0")
    '(円／個)
'    If IsNumeric(Text1(ptxBEF_SEI_RATE).Text) Then
'        Text1(ptxMAIN_KOUTEI_KEI4).Text = Format(ToRoundUp(CCur(CDbl(Text1(ptxMAIN_KOUTEI_KEI3).Text) * CInt((Text1(ptxBEF_SEI_RATE).Text))), 2), "#0.00")
    If IsNumeric(Text1(ptxAFT_SEI_RATE).Text) Then
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
    
        wkint = 0
        If GetIni(Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1), KUSATU_ETC, App.EXEName, c) Then
        Else
            wkKUSATU = Split(Trim(c), ",", -1)
                    
            If IsNumeric(wkKUSATU(1)) Then
                wkint = CInt(wkKUSATU(1))
            End If
        End If
        If GetIni(Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1), Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), App.EXEName, c) Then
            If GetIni(App.EXEName, Trim(StrConv(StrConv(ITEMREC.HIN_NAME, vbUnicode), vbWide)), App.EXEName, c) Then
            Else
                wkKUSATU = Split(Trim(c), ",", -1)
                        
                If IsNumeric(wkKUSATU(0)) Then
                    wkint = CInt(wkKUSATU(1))
                End If
            End If
        Else
            wkKUSATU = Split(Trim(c), ",", -1)
                    
            If IsNumeric(wkKUSATU(1)) Then
                wkint = CInt(wkKUSATU(1))
            End If
        End If
    
        Text1(ptxAFT_KOUTEI_TANI02).Text = Format(wkint, "#0")
    
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
    wkint = 0
    For i = ptxAFT_KOUTEI_KOUSU01 To ptxAFT_KOUTEI_KOUSU09 Step 3
    
        If IsNumeric(Text1(i).Text) Then
            wkint = wkint + CInt(Text1(i).Text)
        End If
    Next i
    Text1(ptxAFT_KOUTEI_KEI1).Text = Format(wkint, "#0")
    
    Text1(ptxAFT_KOUTEI_R_RATE).Text = ToHalfAdjust(CCur(wkint * CDbl(YOYU_RITU(0).Caption)), 0)
    
    '(秒／個)
    If IsNumeric(Text1(ptxBEF_SEI_LOT).Text) Then
        
        If CInt(Text1(ptxBEF_SEI_LOT).Text) = 0 Then    '2010.02.22
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
'    If IsNumeric(Text1(ptxBEF_SEI_RATE).Text) Then
'        Text1(ptxAFT_KOUTEI_KEI4).Text = Format(ToRoundUp(CCur(CDbl(Text1(ptxAFT_KOUTEI_KEI3).Text) * CInt((Text1(ptxBEF_SEI_RATE).Text))), 1), "#0.00")
    If IsNumeric(Text1(ptxAFT_SEI_RATE).Text) Then
        Text1(ptxAFT_KOUTEI_KEI4).Text = Format(ToRoundUp(CCur(CDbl(Text1(ptxAFT_KOUTEI_KEI3).Text) * CInt((Text1(ptxAFT_SEI_RATE).Text))), 2), "#0.00")
    Else
        Text1(ptxAFT_KOUTEI_KEI4).Text = "0.00"
    End If
    
    
    '工程計
    Text1(ptxKOUTEI_KEI1).Text = Format(CInt(Text1(ptxBEF_KOUTEI_KEI1).Text) + CInt(Text1(ptxMAIN_KOUTEI_KEI1).Text) + CInt(Text1(ptxAFT_KOUTEI_KEI1).Text), "#0")
    
    Text1(ptxKOUTEI_R_RATE).Text = Format(CInt(Text1(ptxBEF_KOUTEI_R_RATE).Text) + CInt(Text1(ptxMAIN_KOUTEI_R_RATE).Text) + CInt(Text1(ptxAFT_KOUTEI_R_RATE).Text), "#0")
    
    Text1(ptxKOUTEI_KEI2).Text = Format(CInt(Text1(ptxBEF_KOUTEI_KEI2).Text) + CInt(Text1(ptxPLUS_KOUSU).Text) + CInt(Text1(ptxMAIN_KOUTEI_KEI2).Text) + CInt(Text1(ptxAFT_KOUTEI_KEI2).Text), "#0")
    
'''    Text1(ptxKOUTEI_KEI3).Text = Format(CDbl(Text1(ptxBEF_KOUTEI_KEI3).Text) + CDbl(Text1(ptxMAIN_KOUTEI_KEI3).Text) + CDbl(Text1(ptxAFT_KOUTEI_KEI3).Text), "#0.0")
'''    Text1(ptxKOUTEI_KEI4).Text = Format(CDbl(Text1(ptxBEF_KOUTEI_KEI4).Text) + CDbl(Text1(ptxMAIN_KOUTEI_KEI4).Text) + CDbl(Text1(ptxAFT_KOUTEI_KEI4).Text), "#0.00")
    
    
    
    '(分／個)
    Text1(ptxKOUTEI_KEI3).Text = Format(ToRoundUp(CCur(CInt(Text1(ptxKOUTEI_KEI2).Text) / 60), 1), "#0.0")
    '(円／個)
'    If IsNumeric(Text1(ptxBEF_SEI_RATE).Text) Then
'        Text1(ptxKOUTEI_KEI4).Text = Format(ToRoundUp(CCur(CDbl(Text1(ptxKOUTEI_KEI3).Text) * CInt((Text1(ptxBEF_SEI_RATE).Text))), 1), "#0.00")
    
    If IsNumeric(Text1(ptxAFT_SEI_RATE).Text) Then
        Text1(ptxKOUTEI_KEI4).Text = Format(ToRoundUp(CCur(CDbl(Text1(ptxKOUTEI_KEI3).Text) * CInt((Text1(ptxAFT_SEI_RATE).Text))), 1), "#0.00")
    
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
Dim wkint       As Integer
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
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    wkint = wkint + 1
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
    Text1(ptxBEF_KOUTEI_QTY02).Text = Format(wkint, "#0")
    Text1(ptxBEF_KOUTEI_KOUSU02).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI02).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY02).Text), "#0")
    '�B
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(DOUKON_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = DOUKON_T(j) Then
                    wkint = wkint + 1
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
    Text1(ptxBEF_KOUTEI_QTY03).Text = Format(wkint, "#0")
    Text1(ptxBEF_KOUTEI_KOUSU03).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI03).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY03).Text), "#0")
    '�C
    If KUSATU_F Then
        '草津はINI参照
            
        wkint = 0
        If GetIni(Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1), KUSATU_ETC, App.EXEName, c) Then
        Else
            wkKUSATU = Split(Trim(c), ",", -1)
                    
            If IsNumeric(wkKUSATU(0)) Then
                wkint = CInt(wkKUSATU(0))
            End If
        End If
        If GetIni(Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1), Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), App.EXEName, c) Then
            If GetIni(App.EXEName, Trim(StrConv(StrConv(ITEMREC.HIN_NAME, vbUnicode), vbWide)), App.EXEName, c) Then
            Else
                wkKUSATU = Split(Trim(c), ",", -1)
                        
                If IsNumeric(wkKUSATU(0)) Then
                    wkint = CInt(wkKUSATU(0))
                End If
            End If
        Else
            wkKUSATU = Split(Trim(c), ",", -1)
            If IsNumeric(wkKUSATU(0)) Then
                wkint = CInt(wkKUSATU(0))
            End If
        End If
        Text1(ptxBEF_KOUTEI_TANI04).Text = Format(wkint, "#0")
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
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    wkint = wkint + 1
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
    Text1(ptxBEF_KOUTEI_QTY05).Text = Format(wkint, "#0")
    Text1(ptxBEF_KOUTEI_KOUSU05).Text = Format(CInt(Text1(ptxBEF_KOUTEI_TANI05).Text) * CInt(Text1(ptxBEF_KOUTEI_QTY05).Text), "#0")
    '�E
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(DOUKON_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = DOUKON_T(j) Then
                    wkint = wkint + 1
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
    Text1(ptxBEF_KOUTEI_QTY06).Text = Format(wkint, "#0")
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
    wkint = 0
    For i = ptxBEF_KOUTEI_KOUSU01 To ptxBEF_KOUTEI_KOUSU09 Step 3
    
        wkint = wkint + CInt(Text1(i).Text)
    
    Next i
    
    
    If IsNumeric(Text1(ptxPLUS_KOUSU).Text) Then
        Text1(ptxPLUS_KOUSU).Text = Format(CInt(Text1(ptxPLUS_KOUSU).Text), "#0")
    Else
        Text1(ptxPLUS_KOUSU).Text = "0"
    End If
    
    Text1(ptxBEF_KOUTEI_KEI1).Text = Format(wkint, "#0")
    
    Text1(ptxBEF_KOUTEI_R_RATE).Text = ToHalfAdjust(CCur(wkint * CDbl(YOYU_RITU(0).Caption)), 0)
    
    
    '(秒／個)
    'ptxBEF_SEI_LOT-->ptxAFT_SEI_LOT 2009.11.05
    If IsNumeric(Text1(ptxAFT_SEI_LOT).Text) Then
        
        If CInt(Text1(ptxAFT_SEI_LOT).Text) <> 0 Then
        
            Text1(ptxBEF_KOUTEI_KEI2).Text = Format(ToHalfAdjust(CCur(CInt(Text1(ptxBEF_KOUTEI_R_RATE).Text) / CInt(Text1(ptxAFT_SEI_LOT).Text)), 0), "#0")
        Else
            Text1(ptxBEF_KOUTEI_KEI2).Text = "0"
        End If
   
    
    Else
        Text1(ptxBEF_KOUTEI_KEI2).Text = "0"
    End If
    '(分／個)
    Text1(ptxBEF_KOUTEI_KEI3).Text = Format(ToRoundUp(CCur(CInt(CInt(Text1(ptxBEF_KOUTEI_KEI2).Text) + CInt(Text1(ptxPLUS_KOUSU).Text)) / 60), 1), "#0.0")
    '(円／個)
    'ptxBEF_SEI_RATE-->ptxAFT_SEI_RATE 2009.11.05
    If IsNumeric(Text1(ptxAFT_SEI_RATE).Text) Then
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
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkint = wkint + CInt(ToRoundUp(CCur(KOUSEI(i, ColS_KOUSU)) * CDbl(KOUSEI(i, ColKO_QTY)), 0))
                    End If
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI02).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_QTY02).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU02).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI02).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY02).Text), "#0")
    '�B
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(DOUKON_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = DOUKON_T(j) Then
                    
                    If IsNumeric(KOUSEI(i, ColKO_QTY)) Then
                        wkint = wkint + CInt(ToRoundUp(CCur(KOUSEI(i, ColKO_QTY)), 0))
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
    Text1(ptxMAIN_KOUTEI_QTY03).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_KOUSU03).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI03).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY03).Text), "#0")
    '�C
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(KAKOU_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = KAKOU_T(j) Then
                    If IsNumeric(KOUSEI(i, ColS_KOUSU)) Then
                        wkint = wkint + CInt(KOUSEI(i, ColS_KOUSU))
                    End If
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI04).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_QTY04).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU04).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI04).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY04).Text), "#0")
    '�D
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
            
            
            For j = 0 To UBound(SHIZAI_T)
            
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    If IsNumeric(KOUSEI(i, ColSEI_SYU_KON)) Then
                        wkint = wkint + CInt(KOUSEI(i, ColSEI_SYU_KON))
                    End If
                End If
            
            Next j
            
            
            
            
        Next i
    End If
    Text1(ptxMAIN_KOUTEI_TANI05).Text = Format(wkint, "#0")
    Text1(ptxMAIN_KOUTEI_QTY05).Text = "1"
    Text1(ptxMAIN_KOUTEI_KOUSU05).Text = Format(CInt(Text1(ptxMAIN_KOUTEI_TANI05).Text) * CInt(Text1(ptxMAIN_KOUTEI_QTY05).Text), "#0")
    '計
    wkint = 0
    For i = ptxMAIN_KOUTEI_KOUSU01 To ptxMAIN_KOUTEI_KOUSU09 Step 3
    
        If IsNumeric(Text1(i).Text) Then
            wkint = wkint + CInt(Text1(i).Text)
        End If
    Next i
    Text1(ptxMAIN_KOUTEI_KEI1).Text = Format(wkint, "#0")
    
    Text1(ptxMAIN_KOUTEI_R_RATE).Text = ToHalfAdjust(CCur(wkint * CDbl(YOYU_RITU(0).Caption)), 0)
    '(秒／個)
'2009.11.10    If IsNumeric(Text1(ptxAFT_SEI_LOT).Text) Then
        Text1(ptxMAIN_KOUTEI_KEI2).Text = Format(ToHalfAdjust(CCur(Text1(ptxMAIN_KOUTEI_R_RATE).Text), 0), "#0")
'    Else
'        Text1(ptxMAIN_KOUTEI_KEI2).Text = "0"
'    End If
    '(分／個)
    Text1(ptxMAIN_KOUTEI_KEI3).Text = Format(ToRoundUp(CCur(CInt(Text1(ptxMAIN_KOUTEI_KEI2).Text) / 60), 1), "#0.0")
    '(円／個)
    If IsNumeric(Text1(ptxAFT_SEI_RATE).Text) Then
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
    
        wkint = 0
        If GetIni(Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1), KUSATU_ETC, App.EXEName, c) Then
        Else
            wkKUSATU = Split(Trim(c), ",", -1)
                    
            If IsNumeric(wkKUSATU(1)) Then
                wkint = CInt(wkKUSATU(1))
            End If
        End If
        If GetIni(Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1), Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), App.EXEName, c) Then
            If GetIni(App.EXEName, Trim(StrConv(StrConv(ITEMREC.HIN_NAME, vbUnicode), vbWide)), App.EXEName, c) Then
            Else
                wkKUSATU = Split(Trim(c), ",", -1)
                        
                If IsNumeric(wkKUSATU(0)) Then
                    wkint = CInt(wkKUSATU(1))
                End If
            End If
        Else
            wkKUSATU = Split(Trim(c), ",", -1)
                    
            If IsNumeric(wkKUSATU(1)) Then
                wkint = CInt(wkKUSATU(1))
            End If
        End If
    
        Text1(ptxAFT_KOUTEI_TANI02).Text = Format(wkint, "#0")
    
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
    wkint = 0
    For i = ptxAFT_KOUTEI_KOUSU01 To ptxAFT_KOUTEI_KOUSU09 Step 3
    
        If IsNumeric(Text1(i).Text) Then
            wkint = wkint + CInt(Text1(i).Text)
        End If
    Next i
    Text1(ptxAFT_KOUTEI_KEI1).Text = Format(wkint, "#0")
    Text1(ptxAFT_KOUTEI_R_RATE).Text = ToHalfAdjust(CCur(wkint * CDbl(YOYU_RITU(0).Caption)), 0)
    '(秒／個)
    If IsNumeric(Text1(ptxAFT_SEI_LOT).Text) Then
        If Val(Text1(ptxAFT_SEI_LOT).Text) <> 0 Then
            Text1(ptxAFT_KOUTEI_KEI2).Text = Format(ToHalfAdjust(CCur(CInt(Text1(ptxAFT_KOUTEI_R_RATE).Text) / CInt(Text1(ptxAFT_SEI_LOT).Text)), 0), "#0")
        Else
            Text1(ptxAFT_KOUTEI_KEI2).Text = "0"
        End If
    Else
        Text1(ptxAFT_KOUTEI_KEI2).Text = "0"
    End If
    '(分／個)
    Text1(ptxAFT_KOUTEI_KEI3).Text = Format(ToRoundUp(CCur(CInt(Text1(ptxAFT_KOUTEI_KEI2).Text) / 60), 1), "#0.0")
    '(円／個)
    If IsNumeric(Text1(ptxAFT_SEI_RATE).Text) Then
        Text1(ptxAFT_KOUTEI_KEI4).Text = Format(ToRoundUp(CCur(CDbl(Text1(ptxAFT_KOUTEI_KEI3).Text) * CInt((Text1(ptxAFT_SEI_RATE).Text))), 2), "#0.00")
    Else
        Text1(ptxAFT_KOUTEI_KEI4).Text = "0.00"
    End If
    
    
    If IsNumeric(Text1(ptxPLUS_KOUSU).Text) Then
        Text1(ptxPLUS_KOUSU).Text = Format(CInt(Text1(ptxPLUS_KOUSU).Text), "#0")
    Else
        Text1(ptxPLUS_KOUSU).Text = "0"
    End If
    
    
    
    '工程計
    Text1(ptxKOUTEI_KEI1).Text = Format(CInt(Text1(ptxBEF_KOUTEI_KEI1).Text) + CInt(Text1(ptxMAIN_KOUTEI_KEI1).Text) + CInt(Text1(ptxAFT_KOUTEI_KEI1).Text), "#0")
    
    Text1(ptxKOUTEI_R_RATE).Text = Format(CInt(Text1(ptxBEF_KOUTEI_R_RATE).Text) + CInt(Text1(ptxMAIN_KOUTEI_R_RATE).Text) + CInt(Text1(ptxAFT_KOUTEI_R_RATE).Text), "#0")
    
    Text1(ptxKOUTEI_KEI2).Text = Format(CInt(Text1(ptxBEF_KOUTEI_KEI2).Text) + CInt(Text1(ptxPLUS_KOUSU).Text) + CInt(Text1(ptxMAIN_KOUTEI_KEI2).Text) + CInt(Text1(ptxAFT_KOUTEI_KEI2).Text), "#0")
    Text1(ptxKOUTEI_KEI3).Text = Format(CDbl(Text1(ptxBEF_KOUTEI_KEI3).Text) + CDbl(Text1(ptxMAIN_KOUTEI_KEI3).Text) + CDbl(Text1(ptxAFT_KOUTEI_KEI3).Text), "#0.0")
    Text1(ptxKOUTEI_KEI4).Text = Format(CDbl(Text1(ptxBEF_KOUTEI_KEI4).Text) + CDbl(Text1(ptxMAIN_KOUTEI_KEI4).Text) + CDbl(Text1(ptxAFT_KOUTEI_KEI4).Text), "#0.00")
    
    
    
    '(分／個)
    
    
    Text1(ptxKOUTEI_KEI3).Text = Format(ToRoundUp(CCur(CInt(Text1(ptxKOUTEI_KEI2).Text) / 60), 1), "#0.0")
    '(円／個)
    If IsNumeric(Text1(ptxAFT_SEI_RATE).Text) Then
        Text1(ptxKOUTEI_KEI4).Text = Format(ToRoundUp(CCur(CDbl(Text1(ptxKOUTEI_KEI3).Text) * CInt((Text1(ptxAFT_SEI_RATE).Text))), 1), "#0.00")
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
            
            
'                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                    If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
            
                        wkDouble = wkDouble + CDbl(KOUSEI(i, ColG_ST_URIKIN_KUSATU))
                    End If
            
                End If
        
        
            Next i
        End If
    End If
    Text1(ptxAFT_S_GAISO_TANKA).Text = Format(wkDouble, "#0.00")




    'PPSC原価   2011.06.23
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
    
    
Dim wkint       As Integer
    
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


    '新単価−−＞旧単価 2009.06.02
    Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode))
    Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode))
    Call UniCode_Conv(ITEMREC.BEF_S_GAISO_TANKA, StrConv(ITEMREC.S_GAISO_TANKA, vbUnicode))
    Call UniCode_Conv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, StrConv(ITEMREC.S_PPSC_KAKO_KOSU, vbUnicode))
    Call UniCode_Conv(ITEMREC.BEF_S_BU_KAKO_KOSU, StrConv(ITEMREC.S_BU_KAKO_KOSU, vbUnicode))



    'ロット数
    Call UniCode_Conv(ITEMREC.SEI_LOT, Format(CLng(Text1(ptxAFT_SEI_LOT).Text), "00000000"))
    '分レート
    Call UniCode_Conv(ITEMREC.SEI_RATE, Format(CDbl(Text1(ptxAFT_SEI_RATE).Text), "0000.00"))
    '工数
    Call UniCode_Conv(ITEMREC.S_KOUSU, Format(CDbl(Text1(ptxAFT_S_KOUSU).Text), "00000.00"))
    '工数原価
    Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, Format(CDbl(Text1(ptxAFT_S_KOUSU_GENKA).Text), "00000000.00"))
    '工数売価
    Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, Format(CDbl(Text1(ptxAFT_S_KOUSU_BAIKA).Text), "00000000.00"))
    '設定日
    Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, Format(Now, "YYYYMMDD"))
    
    
    '箱代原価
    Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, Format(CDbl(Text1(ptxAFT_S_SHIZAI_GENKA).Text), "00000000.00"))
    '箱代売価
    Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, Format(CDbl(Text1(ptxAFT_S_SHIZAI_BAIKA).Text), "00000000.00"))
    
    
    
    '外装箱代
    If IsNumeric(Text1(ptxAFT_S_GAISO_TANKA).Text) Then
        Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, Format(CDbl(Text1(ptxAFT_S_GAISO_TANKA).Text), "00000.00"))
    Else
        Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, "00000.00")
    End If
    
    
    'PPSC単価
    
    If IsNumeric(Text1(ptxAFT_S_PPSC_KAKO_KOSU).Text) Then
        Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, Format(CDbl(Text1(ptxAFT_S_PPSC_KAKO_KOSU).Text), "00000.00"))
    Else
        Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, "00000.00")
    End If
    'BU単価
    If IsNumeric(Text1(ptxAFT_S_BU_KAKO_KOSU).Text) Then
       Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, Format(CDbl(Text1(ptxAFT_S_BU_KAKO_KOSU).Text), "00000.00"))
    Else
       Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, "00000.00")
    End If
    
    
    
    '設定日
    Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, Format(Now, "YYYYMMDD"))
    '担当者
    Call UniCode_Conv(ITEMREC.SEI_TANKA_TANTO, Text1(ptxTanto_Code).Text)
    'メモ
    Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, Text1(ptxAFT_SE_TANKA_MEMO).Text)
    
    'ラベル貼り付け枚数
    Call UniCode_Conv(ITEMREC.SEI_LABEL_QTY, Format(CInt(Text1(ptxMAIN_KOUTEI_QTY01).Text), "00"))
    
    '更新担当者
    Call UniCode_Conv(ITEMREC.UPD_TANTO, Text1(ptxTanto_Code).Text)
    '更新 日時
    Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
    
    
    '2008.09.03 追加↓
    
    '仕向け先
    Call UniCode_Conv(ITEMREC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    
        
    '資材件数
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(SHIZAI_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = SHIZAI_T(j) Then
                    wkint = wkint + 1
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    Call UniCode_Conv(ITEMREC.SEI_SZI_CNT, Format(wkint, "00"))
        
    '同梱件数
    wkint = 0
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then
    Else
        For i = 1 To KOUSEI.UpperBound(1)
        
            For j = 0 To UBound(DOUKON_T)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = DOUKON_T(j) Then
                    wkint = wkint + 1
                    Exit For
                End If
        
            Next j
        
        Next i
    End If
    Call UniCode_Conv(ITEMREC.SEI_DKN_CNT, Format(wkint, "00"))
        
        
        
    

    
    
    '2008.09.03 追加↑
    
    
    
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
    
    '2009.09.17 PLUS工数加算
    If IsNumeric(Text1(ptxPLUS_KOUSU).Text) Then
        Call UniCode_Conv(ITEMREC.PLUS_KOUSU, Format(CDbl(Text1(ptxPLUS_KOUSU).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.PLUS_KOUSU, "000.00")
    End If
    
    
    
    
    
    
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
    
    
    
    '2009.06.5 追加↓
    'メモ
    Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, Text1(ptxAFT_SE_TANKA_MEMO).Text)
    '見積書備考
    Call UniCode_Conv(ITEMREC.M_BIKOU, RichTextBox1(prchM_BIKOU).Text)
    '仕様書��
    Call UniCode_Conv(ITEMREC.SHIYOU_NO, Text1(ptxSHIYOU_NO).Text)
    '見積区分
    Call UniCode_Conv(ITEMREC.MITSUMORI_KBN, Text1(ptxMITSUMORI_KBN).Text)
    '単価切替日
    Call UniCode_Conv(ITEMREC.TANKA_KIRIKAE_DT, Text1(ptxTANKA_KIRIKAE_DT).Text)
    '切替区分
    Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, Text1(ptxKIRIKAE_KBN).Text)
    '2009.06.5 追加↑
    
    
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
    
    
Dim K_SEQNO As Integer
Dim G_SEQNO As Integer
Dim D_SEQNO As Integer
    
    
    
    
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
    
    
    
            Select Case Right(KOUSEI(i, ColKO_SYUBETSU), 2)
            
                Case KOSOU_KBN          '個装
                    K_SEQNO = K_SEQNO + 10
                
                    If K_SEQNO > 50 Then
                        MsgBox "個装資材登録件数がオーバーしています。"
                        Exit Function
                    End If
                
                Case GAISO_KBN          '外装
                    G_SEQNO = G_SEQNO + 10
                    If G_SEQNO > 30 Then
                        MsgBox "外装資材登録件数がオーバーしています。"
                        Exit Function
                    End If
                Case Else               '同梱
                    D_SEQNO = D_SEQNO + 10
                    If D_SEQNO > 250 Then
                        MsgBox "同梱登録件数がオーバーしています。"
                        Exit Function
                    End If
            End Select
    
    
    
    
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
            
            Select Case StrConv(ITEMREC.SEI_KBN, vbUnicode)
            
            
                Case "1"
            
                    KOUSEI(i, ColG_ST_URITAN) = "別売"
            
                Case "2"
            
                    KOUSEI(i, ColG_ST_URITAN) = "支給"
            
            
                Case Else
                    If Trim(KOUSEI(i, ColG_ST_URITAN)) = "" Then
                        KOUSEI(i, ColG_ST_URITAN) = "0.00"
                    End If
                    
                    If IsNumeric(KOUSEI(i, ColG_ST_URITAN)) Then
                        KOUSEI(i, ColG_ST_URITAN) = Format(CDbl(KOUSEI(i, ColG_ST_URITAN)), "#0.00")
                    Else
                        MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(販売＠)"
            
                    End If
            
            End Select
            
            
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
            
            If StrConv(ITEMREC.SEI_KBN, vbUnicode) <> "1" And StrConv(ITEMREC.SEI_KBN, vbUnicode) <> "2" Then
            
                If Trim(KOUSEI(i, ColG_ST_URIKIN)) = "" Then
                    KOUSEI(i, ColG_ST_URIKIN) = "0.00"
                End If
                If IsNumeric(KOUSEI(i, ColG_ST_URIKIN)) Then
                    KOUSEI(i, ColG_ST_URIKIN) = Format(CDbl(KOUSEI(i, ColG_ST_URIKIN)), "#0.00")
                Else
                    MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(販売金額計)"
        
                End If
            
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
                                        
'    Call Input_Lock
                                        
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
    
    
    
    '---------------------------------------------------    '単価更新
'    MESG = "単価を登録します。よろしいですか？" & vbCrLf
'    MESG = MESG & "　ロット数：" & Text1(ptxAFT_SEI_LOT).Text & vbCrLf
'    MESG = MESG & "　分レート：" & Text1(ptxAFT_SEI_RATE).Text & vbCrLf
'    MESG = MESG & "　工数：" & Text1(ptxAFT_S_KOUSU).Text & vbCrLf
'    MESG = MESG & "　（原価）工料：" & Text1(ptxAFT_S_KOUSU_GENKA).Text & vbCrLf
'    MESG = MESG & "　 (売価) 工料：" & Text1(ptxAFT_S_KOUSU_BAIKA).Text & vbCrLf
'    MESG = MESG & "　（原価）工料：" & Text1(ptxAFT_S_SHIZAI_GENKA).Text & vbCrLf
'    MESG = MESG & "　 (売価) 工料：" & Text1(ptxAFT_S_SHIZAI_BAIKA).Text & vbCrLf
'    MESG = MESG & "　 設定日：" & Text1(ptxAFT_S_KOUSU_SET_DATE).Text & vbCrLf
'    MESG = MESG & "　 担当者：" & Text1(ptxAFT_SEI_TANKA_TANTO).Text & vbCrLf
'    MESG = MESG & "　 メモ：" & Text1(ptxAFT_SE_TANKA_MEMO).Text & vbCrLf
            
'    ans = MsgBox(MESG, vbYesNo + vbDefaultButton1 + vbExclamation, "確認入力")
'    If ans = vbYes Then
'        If Tanka_Update_Proc() Then
'            GoTo Abort_Tran
'        End If
'
'    End If
    
    
        
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


    '2009.03.24
    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then

    Else


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
                        
                            MsgBox "他端末でデータが、変更されています。構成−保存処理を中止します。"
                            Update_Proc = False
                            GoTo Abort_Tran
                        
                        
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Update_Proc = False
                                GoTo Abort_Tran
                            End If
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                            GoTo Abort_Tran
                    
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
                        
                            MsgBox "他端末でデータが、変更されています。構成−保存処理を中止します。"
                            Update_Proc = False
                            GoTo Abort_Tran
                        
                        
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Update_Proc = False
                                GoTo Abort_Tran
                            End If
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                            GoTo Abort_Tran
                    
                    End Select
                
                Loop
    
            End If
        Next i
    End If


    '---------------------------------------------------    '品目ﾏｽﾀ　親品番更新    2009.06.02

    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)

    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
            
                MsgBox "他端末でデータが、変更されています。構成−保存処理を中止します。"
                Update_Proc = False
                GoTo Abort_Tran
            
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Update_Proc = False
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                GoTo Abort_Tran
        
        End Select
    Loop

    'メモ
'    Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, Text1(ptxBEF_SE_TANKA_MEMO).Text)
    '見積書備考
    Call UniCode_Conv(ITEMREC.M_BIKOU, RichTextBox1(prchM_BIKOU).Text)
    '仕様書��
    Call UniCode_Conv(ITEMREC.SHIYOU_NO, Text1(ptxSHIYOU_NO).Text)
    '見積区分
    Call UniCode_Conv(ITEMREC.MITSUMORI_KBN, Text1(ptxMITSUMORI_KBN).Text)
    '単価切替日
    Call UniCode_Conv(ITEMREC.TANKA_KIRIKAE_DT, Text1(ptxTANKA_KIRIKAE_DT).Text)
    '切替区分
    Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, Text1(ptxKIRIKAE_KBN).Text)




    '-----  単価欄 2009.07.24
    'ロット数
    
    If IsNumeric(Text1(ptxBEF_SEI_LOT).Text) Then
        Call UniCode_Conv(ITEMREC.SEI_LOT, Format(CLng(Text1(ptxBEF_SEI_LOT).Text), "00000000"))
    Else
        Call UniCode_Conv(ITEMREC.SEI_LOT, "")
    End If
      '分レート
    If IsNumeric(Text1(ptxBEF_SEI_RATE).Text) Then
        Call UniCode_Conv(ITEMREC.SEI_RATE, Format(CDbl(Text1(ptxBEF_SEI_RATE).Text), "0000.00"))
    Else
        Call UniCode_Conv(ITEMREC.SEI_RATE, "")
    End If
    '工数
    If IsNumeric(Text1(ptxBEF_S_KOUSU).Text) Then
        Call UniCode_Conv(ITEMREC.S_KOUSU, Format(CDbl(Text1(ptxBEF_S_KOUSU).Text), "00000.00"))
    Else
        Call UniCode_Conv(ITEMREC.S_KOUSU, "")
    End If
    '工数原価
    If IsNumeric(Text1(ptxBEF_S_KOUSU_GENKA).Text) Then
        Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, Format(CDbl(Text1(ptxBEF_S_KOUSU_GENKA).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, "")
    End If
    '工数売価
    If IsNumeric(Text1(ptxBEF_S_KOUSU_BAIKA).Text) Then
        Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, Format(CDbl(Text1(ptxBEF_S_KOUSU_BAIKA).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "")
    End If
    '設定日
    If Trim(Text1(ptxBEF_S_KOUSU_SET_DATE).Text) = "" Then
        Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, "")
    Else
        Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, Format(Now, "YYYYMMDD"))
    End If
    '箱代原価
    If IsNumeric(Text1(ptxBEF_S_SHIZAI_GENKA).Text) Then
        Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, Format(CDbl(Text1(ptxBEF_S_SHIZAI_GENKA).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, "")
    End If
    '箱代売価
    If IsNumeric(Text1(ptxBEF_S_SHIZAI_BAIKA).Text) Then
        Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, Format(CDbl(Text1(ptxBEF_S_SHIZAI_BAIKA).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "")
    End If
    '外装箱代
    If IsNumeric(Text1(ptxBEF_S_GAISO_TANKA).Text) Then
        Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, Format(CDbl(Text1(ptxBEF_S_GAISO_TANKA).Text), "00000.00"))
    Else
        Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, "")
    End If
    'PPSC単価
    If IsNumeric(Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text) Then
        Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, Format(CDbl(Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text), "00000.00"))
    Else
        Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, "")
    End If
    'BU単価
    If IsNumeric(Text1(ptxBEF_S_BU_KAKO_KOSU).Text) Then
       Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, Format(CDbl(Text1(ptxBEF_S_BU_KAKO_KOSU).Text), "00000.00"))
    Else
       Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, "")
    End If
    '設定日
    If Trim(Text1(ptxBEF_S_KOUSU_SET_DATE).Text) = "" Then
        Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, "")
    Else
        Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, Format(Now, "YYYYMMDD"))
    End If
    '担当者
    Call UniCode_Conv(ITEMREC.SEI_TANKA_TANTO, Text1(ptxTanto_Code).Text)
    'メモ
    Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, Text1(ptxBEF_SE_TANKA_MEMO).Text)
    'ラベル貼り付け枚数
    Call UniCode_Conv(ITEMREC.SEI_LABEL_QTY, Format(CInt(Text1(ptxMAIN_KOUTEI_QTY01).Text), "00"))
    
    
    
    '工数売価
    Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, Format(CDbl(Text1(ptxOLD_S_KOUSU_BAIKA).Text), "00000000.00"))
    '箱代売価
    Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, Format(CDbl(Text1(ptxOLD_S_SHIZAI_BAIKA).Text), "00000000.00"))
    '外装箱代
    If IsNumeric(Text1(ptxOLD_S_GAISO_TANKA).Text) Then
        Call UniCode_Conv(ITEMREC.BEF_S_GAISO_TANKA, Format(CDbl(Text1(ptxOLD_S_GAISO_TANKA).Text), "00000.00"))
    Else
        Call UniCode_Conv(ITEMREC.BEF_S_GAISO_TANKA, "00000.00")
    End If
    'PPSC単価
    If IsNumeric(Text1(ptxOLD_S_PPSC_KAKO_KOSU).Text) Then
        Call UniCode_Conv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, Format(CDbl(Text1(ptxOLD_S_PPSC_KAKO_KOSU).Text), "00000.00"))
    Else
        Call UniCode_Conv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, "00000.00")
    End If
    'BU単価
    If IsNumeric(Text1(ptxOLD_S_BU_KAKO_KOSU).Text) Then
       Call UniCode_Conv(ITEMREC.BEF_S_BU_KAKO_KOSU, Format(CDbl(Text1(ptxOLD_S_BU_KAKO_KOSU).Text), "00000.00"))
    Else
       Call UniCode_Conv(ITEMREC.BEF_S_BU_KAKO_KOSU, "00000.00")
    End If
    
    '2009.09.17 PLUS工数加算
    If IsNumeric(Text1(ptxPLUS_KOUSU).Text) Then
        Call UniCode_Conv(ITEMREC.PLUS_KOUSU, Format(CDbl(Text1(ptxPLUS_KOUSU).Text), "000.00"))
    Else
        Call UniCode_Conv(ITEMREC.PLUS_KOUSU, "000.00")
    End If
    
    
    
    
    
    '-----  単価欄 2009.07.24

    '更新担当者
    Call UniCode_Conv(ITEMREC.UPD_TANTO, Text1(ptxTanto_Code).Text)
    '更新 日時
    Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))


    Do
        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
            
                MsgBox "他端末でデータが、変更されています。構成−保存処理を中止します。"
                Update_Proc = False
                GoTo Abort_Tran
            
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Update_Proc = False
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                GoTo Abort_Tran
        
        End Select
    
    Loop

End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
'    Call Input_UnLock
    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    
'    Call Input_UnLock

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


Private Sub Estimate_Head_Proc(excelApplication As Object, excelWorkBook As Object, excelSheet As Object)                               '2013.09.30 EXCEL Ver対応
'Private Sub Estimate_Head_Proc(excelApplication As Excel.Application, excelWorkBook As Excel.Workbook, excelSheet As Excel.Worksheet)
'----------------------------------------------------------------------------
'                   ＥＸＣＥＬ（御見積書ヘッダー）出力
'       2009.06.02
'----------------------------------------------------------------------------
Dim i   As Integer
Debug.Print "in Estimate_head_Proc=" & Now
    
    excelApplication.StandardFontSize = 11
    
    excelApplication.StandardFont = "ＭＳ　Ｐゴシック"
    
    'ページ設定
    
    If Trim(EXCEL_TEMPLATE) = "" Then
    
        With excelSheet.Application.ActiveSheet.PageSetup
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .Orientation = xlPortrait
        End With
    
    Else
    
        With excelSheet.Application.ActiveSheet.PageSetup
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
        End With
    
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 2011.02.22
    '列の幅
'    excelSheet.Application.Columns("A").ColumnWidth = 2.75
'    excelSheet.Application.Columns("B:D").ColumnWidth = 8.5
'    excelSheet.Application.Columns("E").ColumnWidth = 12
'    excelSheet.Application.Columns("F").ColumnWidth = 6.25
'    excelSheet.Application.Columns("G").ColumnWidth = 6
'    excelSheet.Application.Columns("H").ColumnWidth = 6.25
'    excelSheet.Application.Columns("I:J").ColumnWidth = 8.5
'    excelSheet.Application.Columns("K").ColumnWidth = 12
'    excelSheet.Application.Columns("L").ColumnWidth = 3
    
    
'2010.05.13
'    excelSheet.Application.Columns("M").ColumnWidth = 4.88
'    excelSheet.Application.Columns("N").ColumnWidth = 8.38
'    excelSheet.Application.Columns("O").ColumnWidth = 8.38
'    excelSheet.Application.Columns("P").ColumnWidth = 4.5
'    excelSheet.Application.Columns("Q").ColumnWidth = 7
'    excelSheet.Application.Columns("R").ColumnWidth = 8.38
'    excelSheet.Application.Columns("S").ColumnWidth = 8.38
'2010.05.13

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 2011.02.22
    
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 2011.03.16
    '行の幅
''    excelSheet.Application.Rows(1).RowHeight = 28.5
''    excelSheet.Application.Rows(2).RowHeight = 13.5
''    excelSheet.Application.Rows("3:4").RowHeight = 17.25
''    excelSheet.Application.Rows("5:8").RowHeight = 13.5
''    excelSheet.Application.Rows("9:10").RowHeight = 28.5
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 2011.03.16




'---    １行目
    'セルの結合
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 5), excelSheet.Application.Cells(1, 8)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 5), excelSheet.Application.Cells(1, 8)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 5), excelSheet.Application.Cells(1, 8)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 5), excelSheet.Application.Cells(1, 8)).Font.FontStyle = "太字"
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 5), excelSheet.Application.Cells(1, 8)).Font.Size = 24
    excelSheet.Application.Cells(1, 5).Value = "　御　見　積　書　"
'---    ２行目
    'セルの結合
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 10), excelSheet.Application.Cells(2, 12)).HorizontalAlignment = xlRight
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 10), excelSheet.Application.Cells(2, 12)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 10), excelSheet.Application.Cells(2, 12)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 10), excelSheet.Application.Cells(2, 12)).Font.Size = 11
    excelSheet.Application.Cells(2, 10).Value = Format(Now, "yyyy年m月d日")
'---    ３行目
    excelSheet.Application.Cells(3, 1).Font.Size = 13
    excelSheet.Application.Cells(3, 1).Value = Trim(EX_NAME1)
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(3, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(3, 5)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(3, 5)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'---    ４行目
    
    If Trim(EX_NAME2) <> "" Then
    
        excelSheet.Application.Cells(4, 1).Font.Size = 13
        excelSheet.Application.Cells(4, 1).Value = Trim(EX_NAME2)
        excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(4, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(4, 5)).Borders(xlEdgeBottom).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(4, 5)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    End If
'---    ５行目
    excelSheet.Application.Cells(5, 1).Font.Size = 9
    excelSheet.Application.Cells(5, 1).Value = Trim(EX_BIKOU1)
    
    
    excelSheet.Application.Cells(5, 12).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(5, 12).Value = Trim(EX_SYAMEI)
'---    ６行目
    excelSheet.Application.Cells(6, 1).Font.Size = 9
    excelSheet.Application.Cells(6, 1).Value = Trim(EX_BIKOU2)
        
    
    excelSheet.Application.Range(excelSheet.Application.Cells(6, 9), excelSheet.Application.Cells(6, 12)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(6, 9), excelSheet.Application.Cells(6, 12)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(6, 9), excelSheet.Application.Cells(6, 12)).MergeCells = True
    excelSheet.Application.Cells(6, 9).Font.Size = 9
    excelSheet.Application.Cells(6, 9).Value = Trim(EX_ADDR1)
'---    ７行目
    excelSheet.Application.Range(excelSheet.Application.Cells(7, 9), excelSheet.Application.Cells(7, 12)).HorizontalAlignment = xlRight
    excelSheet.Application.Range(excelSheet.Application.Cells(7, 9), excelSheet.Application.Cells(7, 12)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(7, 9), excelSheet.Application.Cells(7, 12)).MergeCells = True
    excelSheet.Application.Cells(7, 9).Font.Size = 9
    excelSheet.Application.Cells(7, 9).Value = Trim(EX_ADDR2)


'---    ８行目
    excelSheet.Application.Cells(8, 10).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(8, 10).Value = Trim(EX_CENTER_NAME)
'---    ９行目
    excelSheet.Application.Range(excelSheet.Application.Cells(9, 8), excelSheet.Application.Cells(9, 10)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(9, 8), excelSheet.Application.Cells(9, 10)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(9, 8), excelSheet.Application.Cells(9, 10)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(9, 8), excelSheet.Application.Cells(9, 10)).Font.Size = 9
    excelSheet.Application.Cells(9, 8).Value = Trim(EX_CENTER_ADDR1)
    excelSheet.Application.Cells(9, 8).ShrinkToFit = True
        
'---    10行目
    excelSheet.Application.Range(excelSheet.Application.Cells(10, 8), excelSheet.Application.Cells(10, 10)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(10, 8), excelSheet.Application.Cells(10, 10)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(10, 8), excelSheet.Application.Cells(10, 10)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(10, 8), excelSheet.Application.Cells(10, 10)).Font.Size = 9
    excelSheet.Application.Cells(10, 8).Value = Trim(EX_CENTER_ADDR2)
    excelSheet.Application.Cells(10, 8).ShrinkToFit = True
        
''    With excelApplication.ActiveSheet.Shapes.AddTextbox(1, 491.25, 118.5, 49.5, 19.5).TextFrame
'''''    With excelApplication.ActiveSheet.Shapes.AddTextbox(1, 491.25, 118.5, 48.75, 19.5)
'''''
'''''        .TextFrame.Characters.Text = "承認印"
'''''        .TextFrame.Characters.Font.Size = 8              ' フォントサイズを10ポイントに
'''''        .TextFrame.Characters.Font.NAME = "ＭＳ Ｐゴシック"              ' フォントサイズを10ポイントに
'''''        .TextFrame.HorizontalAlignment = xlHAlignCenter   ' 中央揃え
'''''        .TextFrame.VerticalAlignment = xlVAlignCenter     ' 中央揃え
'''''
'''''        .Line.ForeColor.SchemeColor = 8
'''''
'''''
'''''    End With
    
''    excelApplication.ActiveSheet.Shapes("Text Box 1").Line.ForeColor.SchemeColor = 8
''
''
''    With excelApplication.ActiveSheet.Shapes.AddTextbox(1, 491.25, 118.5, 48.75, 19.5).TextFrame
''
''        .Characters.Text = "承認印"
''        .Characters.Font.Size = 8              ' フォントサイズを10ポイントに
''        .Characters.Font.NAME = "ＭＳ Ｐゴシック"              ' フォントサイズを10ポイントに
''        .HorizontalAlignment = xlHAlignCenter   ' 中央揃え
''        .VerticalAlignment = xlVAlignCenter     ' 中央揃え
''    End With
    
    
    
    
    
    
    
    
''    With excelApplication.ActiveSheet.Shapes.AddTextbox(1, 540.75, 118.5, 49.5, 19.5).TextFrame
''    With excelApplication.ActiveSheet.Shapes.AddTextbox(1, 540, 118.5, 48.75, 19.5).TextFrame
''        .Characters.Text = "担当印"
''        .Characters.Font.Size = 8              ' フォントサイズを10ポイントに
''        .Characters.Font.NAME = "ＭＳ Ｐゴシック"              ' フォントサイズを10ポイントに
''        .HorizontalAlignment = xlHAlignCenter   ' 中央揃え
''        .VerticalAlignment = xlVAlignCenter     ' 中央揃え
''    End With
''    excelApplication.ActiveSheet.Shapes("Text Box 2").Line.ForeColor.SchemeColor = 8
   
   
'''''    With excelApplication.ActiveSheet.Shapes.AddTextbox(1, 540, 118.5, 48.75, 19.5)
'''''        .TextFrame.Characters.Text = "担当印"
'''''        .TextFrame.Characters.Font.Size = 8              ' フォントサイズを10ポイントに
'''''        .TextFrame.Characters.Font.NAME = "ＭＳ Ｐゴシック"              ' フォントサイズを10ポイントに
'''''        .TextFrame.HorizontalAlignment = xlHAlignCenter   ' 中央揃え
'''''        .TextFrame.VerticalAlignment = xlVAlignCenter     ' 中央揃え
'''''
'''''        .Line.ForeColor.SchemeColor = 8
'''''
'''''    End With
   





''    excelApplication.ActiveSheet.Shapes.AddShape 1, 548.25, 117.75, 97.5, 60.75
    
    
''    excelApplication.ActiveSheet.Shapes.AddShape 1, 491.25, 118.5, 97.5, 60#
    
    
    
    
    
    
    
'''''    excelApplication.ActiveSheet.Shapes.AddLine(491.25, 118.5, 588.75, 118.5).Line.ForeColor.SchemeColor = 8


'''''    excelApplication.ActiveSheet.Shapes.AddLine(491.25, 118.5, 491.25, 178.5).Line.ForeColor.SchemeColor = 8
'''''    excelApplication.ActiveSheet.Shapes.AddLine(588.75, 118.5, 588.75, 178.5).Line.ForeColor.SchemeColor = 8
'''''    excelApplication.ActiveSheet.Shapes.AddLine(491.25, 178.5, 588.75, 178.5).Line.ForeColor.SchemeColor = 8







'    excelApplication.ActiveSheet.Shapes.AddLine 597#, 118.5, 597#, 177.75
'''''    excelApplication.ActiveSheet.Shapes.AddLine(540, 118.5, 540, 178.5).Line.ForeColor.SchemeColor = 8

    
''    With objAppXL.ActiveSheet.Shapes.AddShape(msoShapeFlowchartProcess, _
''            20, 20, 100, 20).TextFrame
''        .Characters.Text = "Allegy"             ' オートシェイプの中に文字
''        .Characters.Font.Size = 10              ' フォントサイズを10ポイントに
''        .HorizontalAlignment = xlHAlignCenter   ' 中央揃え
''        .VerticalAlignment = xlVAlignCenter     ' 中央揃え
''    End With
    
    
    
''    With Selection.Characters.Font
''        .NAME = "ＭＳ Ｐゴシック"
''        .FontStyle = "標準"
''        .Size = 8
''        .Strikethrough = False
''        .Superscript = False
''        .Subscript = False
''        .OutlineFont = False
''        .Shadow = False
''        .Underline = xlUnderlineStyleNone
''        .ColorIndex = xlAutomatic
''    End With
''    With Selection
''        .HorizontalAlignment = xlCenter
''        .VerticalAlignment = xlCenter
''        .ReadingOrder = xlContext
''        .Orientation = xlHorizontal
''        .AutoSize = False
''        .AddIndent = False
''    End With


''    ActiveSheet.Shapes.AddTextbox 1, 597#, 117.75, 48.75, 19.5
''    Selection.Characters.Text = "担当印"
''    With Selection.Characters.Font
''        .NAME = "ＭＳ Ｐゴシック"
''        .FontStyle = "標準"
''        .Size = 8
''        .Strikethrough = False
''        .Superscript = False
''        .Subscript = False
''        .OutlineFont = False
''        .Shadow = False
''        .Underline = xlUnderlineStyleNone
''        .ColorIndex = xlAutomatic
''    End With
''    With Selection
''        .HorizontalAlignment = xlCenter
''        .VerticalAlignment = xlCenter
''        .ReadingOrder = xlContext
''        .Orientation = xlHorizontal
''        .AutoSize = False
''        .AddIndent = False
''    End With


'''''''''''''''''''''''''''''   2011.03.16
'    If Trim(SYONIN_Pic) <> "" Then
'
'         With excelSheet.Pictures.Insert(SYONIN_Pic)
'            .Top = excelSheet.Application.Cells(8, 11).Top
'            .Left = excelSheet.Application.Cells(8, 11).Left
'
'
'
'            .Height = (excelSheet.Application.Cells(8, 11).Height + _
'                        excelSheet.Application.Cells(9, 11).Height + _
'                        Round(excelSheet.Application.Cells(10, 11).Height * 0.75, 0))
'
'
'
'            .Width = (excelSheet.Application.Cells(8, 11).Width + _
'                        excelSheet.Application.Cells(8, 12).Width)
'
'
'''            .NAME = "SYONIN"
'        End With
'''''''''''''''''''''''''''''    2011.03.16



''       For i = 1 To ActiveSheet.Shapes.Count
''
''           If ActiveSheet.Shapes(i).NAME = "SYONIN" Then
''
''                With ActiveSheet.Shapes(i)
''                    .LockAspectRatio = True   '---(1)図形の縦横の比率を固定
''
''                    .Height = (excelSheet.Application.Cells(8, 11).Height + _
''                                excelSheet.Application.Cells(9, 11).Height + _
''                                Round(excelSheet.Application.Cells(10, 11).Height * 0.75, 0))
''
''
''
''                    .Width = (excelSheet.Application.Cells(8, 11).Width + _
''                                excelSheet.Application.Cells(8, 12).Width)
''
''
''                End With
''
''                Exit For
''            End If
''        Next i
'''''''''''''''''''''''''''''    2011.03.16
'    End If
'''''''''''''''''''''''''''''       2011.03.16

Debug.Print "out Estimate_head_Proc=" & Now

End Sub

Private Function Estimate_SHIZAI_Proc(excelApplication As Object, excelWorkBook As Object, excelSheet As Object, row As Integer) As Integer                 '2013.09.30 EXCEL Ver対応
'Private Function Estimate_SHIZAI_Proc(excelApplication As Excel.Application, excelWorkBook As Excel.Workbook, excelSheet As Excel.Worksheet, row As Integer) As Integer
'----------------------------------------------------------------------------
'                   ＥＸＣＥＬ（御見積書 資材）出力
'       2009.06.02
'----------------------------------------------------------------------------
Dim i       As Integer
Dim j       As Integer

Dim com     As Integer
Dim sts     As Integer


Dim wkNum1  As Currency
Dim wkNum2  As Currency

Debug.Print "in Estimate_shizai_Proc=" & Now

    Estimate_SHIZAI_Proc = True
'---    14行目
    excelSheet.Application.Rows(14).RowHeight = 13.5
    excelSheet.Application.Cells(14, 2).Font.Size = 10
    excelSheet.Application.Cells(14, 2).Value = "【副資材費】"
    
    
'---    15行目
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Font.Size = 10
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).HorizontalAlignment = xlCenter
    excelSheet.Application.Cells(15, 2).Value = "資材品番"
    excelSheet.Application.Cells(15, 4).Value = "種別"
    excelSheet.Application.Cells(15, 5).Value = "形式・サイズ等"
    excelSheet.Application.Cells(15, 8).Value = "数量"
    excelSheet.Application.Cells(15, 9).Value = "単価"
    excelSheet.Application.Cells(15, 10).Value = "金 額"
    
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 3)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 3)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 3)).MergeCells = True
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 5), excelSheet.Application.Cells(15, 7)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 5), excelSheet.Application.Cells(15, 7)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 5), excelSheet.Application.Cells(15, 7)).MergeCells = True

    excelSheet.Application.Range(excelSheet.Application.Cells(15, 10), excelSheet.Application.Cells(15, 11)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 10), excelSheet.Application.Cells(15, 11)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 10), excelSheet.Application.Cells(15, 11)).MergeCells = True
    
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
    
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlInsideVertical).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(15, 2), excelSheet.Application.Cells(15, 11)).Borders(xlInsideHorizontal).LineStyle = xlNone
    



'2010.05.13
    excelSheet.Application.Cells(15, 14).Font.Size = 12
    excelSheet.Application.Cells(15, 14).HorizontalAlignment = xlCenter
    excelSheet.Application.Cells(15, 14).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(15, 14).Value = "単価"

    excelSheet.Application.Cells(15, 15).Font.Size = 12
    excelSheet.Application.Cells(15, 15).HorizontalAlignment = xlCenter
    excelSheet.Application.Cells(15, 15).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(15, 15).Value = "チェック"


    excelSheet.Application.Cells(15, 17).Font.Size = 12
    excelSheet.Application.Cells(15, 17).HorizontalAlignment = xlCenter
    excelSheet.Application.Cells(15, 17).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(15, 17).Value = "チェック"

    excelSheet.Application.Cells(16, 17).VerticalAlignment = xlBottom
    excelSheet.Application.Cells(16, 17).FormulaR1C1 = Text1(ptxPLUS_KOUSU).Text
'2010.05.13

    
    
'---    16〜20行目
    If EX_SHIZAI_F Then
        
            
        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHin_Gai).Text)
           
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_KOSOU)
        Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
        
        com = BtOpGetGreaterEqual
            
        row = 15
        Do
            DoEvents
            sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
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
                    Call File_Error(sts, com, "構成マスタ")
                    Exit Function
            End Select
            
'ﾁｪｯｸやめる 2009.07.02
'
'            If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_KOSOU And _
'                StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_GAISOU Then
'                Exit Do
'            End If
        
            If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_KOSOU Then
                Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, KOSOU_KBN)
            End If
            If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_GAISOU Then
                Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, GAISO_KBN)
            End If
        
        
            For j = 0 To UBound(EX_SHIZAI_T)
                If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = EX_SHIZAI_T(j) Then
                    
                    
                    If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = FUTAI_KBN Then   '2009.09.05
                    Else
                    
                    
                    
                        row = row + 1
                        excelSheet.Application.Cells(row, 2).Value = Trim(StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                        
                        If IsNumeric(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) Then
    '                        excelSheet.Application.Cells(row, 9).NumberFormatLocal = "#,##0_ "
                            
                            
                            If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN And CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) <> 0 Then
                                excelSheet.Application.Cells(row, 8).Value = CDbl(ToHalfAdjust(1 / CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), 3))
                            Else
                                excelSheet.Application.Cells(row, 8).Value = CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))
                            End If
                        End If
                    
                        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
                        Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
                        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                        Select Case sts
                            Case BtNoErr
                                excelSheet.Application.Cells(row, 4).Value = Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode))
                                
                            Case BtErrKeyNotFound
                            Case Else
                                Call File_Error(sts, com, "コードマスタ")
                                Exit Function
                        End Select
                    
                    
                    
                        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                        
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                                
                                
                                excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row, 7)).HorizontalAlignment = xlLeft
                                excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row, 7)).VerticalAlignment = xlBottom
                                excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row, 7)).MergeCells = True
                                
                                
                                excelSheet.Application.Cells(row, 5).Value = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                                '2009.07.06
                                excelSheet.Application.Cells(row, 5).ShrinkToFit = True
                                
                                excelSheet.Application.Range(excelSheet.Application.Cells(row, 8), excelSheet.Application.Cells(row, 10)).HorizontalAlignment = xlCenter
 
 
 
 '                               If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
 '   '                                excelSheet.Application.Cells(row, 10).NumberFormatLocal = "#,##0_ "
 '                                   excelSheet.Application.Cells(row, 9).Value = CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode))
 '                               Else
 '   '                                excelSheet.Application.Cells(row, 10).Value = "未定"   2009.07.02
 '                                   excelSheet.Application.Cells(row, 9).Value = "別売"
 '                               End If
                                
                                
                                Select Case StrConv(ITEMREC.SEI_KBN, vbUnicode)
                                                    
                                    Case "1"
                                        excelSheet.Application.Cells(row, 9).Value = "別売"
                                    Case "2"
                                        excelSheet.Application.Cells(row, 9).Value = "支給"
                                    Case Else
                                
                                        If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                                            excelSheet.Application.Cells(row, 9).Value = CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode))
                                        Else
                                            excelSheet.Application.Cells(row, 9).Value = "別売"
                                        End If
                                        
                                End Select
                                
                                
                                
                                
                                
                                
                                
                            Case BtErrKeyNotFound
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                Exit Function
                        End Select
                    
                        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).HorizontalAlignment = xlCenter
                        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).VerticalAlignment = xlBottom
                        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).MergeCells = True
                    
                        If IsNumeric(excelSheet.Application.Cells(row, 8).Value) And IsNumeric(excelSheet.Application.Cells(row, 9).Value) Then
                                
    '                        If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
    '
    '
    '                            excelSheet.Application.Cells(row, 11).FormulaR1C1 = "=ROUNDUP(RC[-1]/RC[-2],2)"
    '
    '                        Else
    '                            excelSheet.Application.Cells(row, 11).FormulaR1C1 = "=ROUNDUP(RC[-2]*RC[-1],2)"
    '                        End If
                        
                        
                        
                        
                        
                            If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" Then
                        
                        
    '                            If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
    '
    '
    '                                If Not IsNumeric(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) Then
    '                                Else
    '                                    If Val(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) <> 0 Then
    '                                        excelSheet.Application.Cells(row, 10).Value = CDbl(ToRoundUp(CCur(CDbl(excelSheet.Application.Cells(row, 9).Value) / CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))), 3))
    ''                                    End If
    '                                End If
    '                            Else
                                    excelSheet.Application.Cells(row, 10).FormulaR1C1 = "=ROUNDUP(RC[-2]*RC[-1],2)"
                            
    '                            End If
                                
                        
                                 excelSheet.Application.Cells(row, 10).NumberFormatLocal = "#,##0.00_ "
                                
                            Else
                                
                                If KUSATU_F Then
                            
     '                               If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
     '                                   If Not IsNumeric(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) Then
     '                                   Else
     '                                       If Val(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) <> 0 Then
     '                                           excelSheet.Application.Cells(row, 10).Value = CDbl(ToRoundUp(CCur(CDbl(excelSheet.Application.Cells(row, 9).Value) / CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))), 3))
     '                                       End If
     '                                   End If
     '                               Else
                                        excelSheet.Application.Cells(row, 10).FormulaR1C1 = "=ROUNDUP(RC[-2]*RC[-1],2)"
     '                               End If
                                
                                    excelSheet.Application.Cells(row, 10).NumberFormatLocal = "#,##0.00_ "
                                
                                
                                End If
                                
                                
                                
                            End If
                    
                        
                        
                        
                        
                        
                        End If
                    
                        '2010.05.13
                        excelSheet.Application.Cells(row, 14).HorizontalAlignment = xlRight
                        excelSheet.Application.Cells(row, 14).VerticalAlignment = xlBottom
                        excelSheet.Application.Cells(row, 14).NumberFormatLocal = "#,##0.00_ "
                        excelSheet.Application.Cells(row, 14).FormulaR1C1 = "=RC[-4]"


                        excelSheet.Application.Cells(row, 15).HorizontalAlignment = xlCenter
                        excelSheet.Application.Cells(row, 15).VerticalAlignment = xlBottom
                        excelSheet.Application.Cells(row, 15).FormulaR1C1 = "=IF(RC[-5]=RC[-1],""○"",""×"")"
'
                        '2010.05.13
                    
                    
                    
                    
                    
                    
                    End If  '2009.09.05
                
                
                
                
                
                
                
                
                
                
                End If
            
            
            Next j
        
            com = BtOpGetNext
        
        Loop
        'ﾊﾞｰｺｰﾄﾞﾗﾍﾞﾙ
        If Trim(EX_BCR_CODE) <> "" Then
        
            If IsNumeric(Trim(Text1(ptxMAIN_KOUTEI_QTY01).Text)) Then
                If CDbl(Trim(Text1(ptxMAIN_KOUTEI_QTY01).Text)) > 0 Then
                    row = row + 1
                
                    excelSheet.Application.Cells(row, 2).Value = Trim(EX_BCR_CODE)

                
                    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, EX_BCR_CODE)
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                            excelSheet.Application.Cells(row, 5).Value = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                        Case BtErrKeyNotFound
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Exit Function
                    End Select
                
                                    
                    excelSheet.Application.Range(excelSheet.Application.Cells(row, 8), excelSheet.Application.Cells(row, 10)).HorizontalAlignment = xlCenter
        
        '            excelSheet.Application.Cells(row, 9).NumberFormatLocal = "#,##0_ "
                    excelSheet.Application.Cells(row, 8).Value = CDbl(Trim(Text1(ptxMAIN_KOUTEI_QTY01).Text))
                    excelSheet.Application.Cells(row, 9).Value = "別売"
                
                    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).HorizontalAlignment = xlCenter
                    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).VerticalAlignment = xlBottom
                    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).MergeCells = True
                
                
                     '2010.05.13
                    excelSheet.Application.Cells(row, 14).HorizontalAlignment = xlRight
                    excelSheet.Application.Cells(row, 14).VerticalAlignment = xlBottom
                    excelSheet.Application.Cells(row, 14).NumberFormatLocal = "#,##0.00_ "
                    excelSheet.Application.Cells(row, 14).FormulaR1C1 = "=RC[-4]"


                    excelSheet.Application.Cells(row, 15).HorizontalAlignment = xlCenter
                    excelSheet.Application.Cells(row, 15).VerticalAlignment = xlBottom
                    excelSheet.Application.Cells(row, 15).FormulaR1C1 = "=IF(RC[-5]=RC[-1],""○"",""×"")"
                    
                    '2010.05.13
               
                
                
                End If
            End If
        End If
    
    
'---    明細罫線
        
        If row <> 15 Then
        
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
        
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).LineStyle = xlNone
        
            If row > 16 Then
                excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).Weight = xlThin
                excelSheet.Application.Range(excelSheet.Application.Cells(16, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
            End If
    
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 4), excelSheet.Application.Cells(row, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 4), excelSheet.Application.Cells(row, 4)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 4), excelSheet.Application.Cells(row, 4)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 6), excelSheet.Application.Cells(row, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 6), excelSheet.Application.Cells(row, 5)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 6), excelSheet.Application.Cells(row, 5)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 9), excelSheet.Application.Cells(row, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 9), excelSheet.Application.Cells(row, 8)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 9), excelSheet.Application.Cells(row, 8)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 10), excelSheet.Application.Cells(row, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 10), excelSheet.Application.Cells(row, 9)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 10), excelSheet.Application.Cells(row, 9)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 11), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 11), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(16, 11), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        End If

'        If row <> 15 Or (IsNumeric(Text1(ptxBEF_S_SHIZAI_BAIKA).Text) And Val(Text1(ptxBEF_S_SHIZAI_BAIKA).Text) > 0) Then
'---    27行目
            row = row + 1
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row, 9)).HorizontalAlignment = xlRight
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 9), excelSheet.Application.Cells(row, 9)).VerticalAlignment = xlCenter
            excelSheet.Application.Cells(row, 9).Value = "�@副資材合計金額"
        
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).HorizontalAlignment = xlCenter
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).VerticalAlignment = xlCenter
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).MergeCells = True
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Font.Size = 14
                
''2009.07.01            excelSheet.Application.Cells(row, 11).FormulaR1C1 = "=SUM(R[-1]C:R[" & -row + 15 & "]C)"
            
            
            '合計金額エラーチェック 2009.09.05
            excelSheet.Application.Cells(row, 10).FormulaR1C1 = "=SUM(R[-1]C:R[" & -row + 15 & "]C)"
            
            If IsNumeric(excelSheet.Application.Cells(row, 10).Value) Then
                wkNum1 = CCur(excelSheet.Application.Cells(row, 10).Value)
            Else
                wkNum1 = 0
            End If
            
            If IsNumeric(Text1(ptxBEF_S_SHIZAI_BAIKA).Text) Then
                wkNum2 = CCur(Text1(ptxBEF_S_SHIZAI_BAIKA).Text)
            Else
                wkNum2 = 0
            End If
                        
            
            If IsNumeric(Text1(ptxBEF_S_GAISO_TANKA).Text) Then
                wkNum2 = CCur(wkNum2 + CCur(Text1(ptxBEF_S_GAISO_TANKA).Text))
            End If
            
            
            
'Debug.Print wkNum1 - wkNum2
            
'            If CDbl(excelSheet.Application.Cells(row, 10).Value) <> (CDbl(Text1(ptxBEF_S_SHIZAI_BAIKA).Text) + CDbl(Text1(ptxBEF_S_GAISO_TANKA).Text)) Then
            If wkNum1 <> wkNum2 Then
'                MsgBox "�@副資材合計金額が副資材明細の合計金額と異なります。"
                
Call LOG_OUT(SEI0019_LOG, "�@副資材合計金額が副資材明細の合計金額と異なります。" & Text1(ptxHin_Gai).Text & " " & wkNum1 & " " & wkNum2)
                
                KIN_NG_CNT = KIN_NG_CNT + 1
                txtKIN_NG_CNT.Text = Format(KIN_NG_CNT, "#,##0")
    
                
                excelSheet.Application.Cells(row, 13).Value = "�@副資材合計金額が副資材明細の合計金額と異なります。"
            End If
            
            

            If IsNumeric(Text1(ptxBEF_S_SHIZAI_BAIKA).Text) Then
                excelSheet.Application.Cells(row, 10).Value = Val(Text1(ptxBEF_S_SHIZAI_BAIKA).Text)
            Else
                excelSheet.Application.Cells(row, 10).Value = 0
            End If
'2009.07.06
            If IsNumeric(Text1(ptxAFT_S_GAISO_TANKA).Text) Then
                excelSheet.Application.Cells(row, 10).Value = Val(excelSheet.Application.Cells(row, 10).Value) + Val(Text1(ptxBEF_S_GAISO_TANKA).Text)
            End If
            excelSheet.Application.Cells(row, 10).NumberFormatLocal = "#,##0.00_ "
            
            
            
            
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
        
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).Weight = xlThick
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).Weight = xlThick
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).Weight = xlThick
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).Weight = xlThick
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).LineStyle = xlNone
        
            excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).LineStyle = xlNone
        
        
            excelSheet.Application.Cells(row, 10).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Cells(row, 10).Borders(xlEdgeLeft).Weight = xlThick
            excelSheet.Application.Cells(row, 10).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        
'        End If
    
    
            '2010.05.13
            excelSheet.Application.Cells(row, 14).HorizontalAlignment = xlRight
            excelSheet.Application.Cells(row, 14).VerticalAlignment = xlBottom
            excelSheet.Application.Cells(row, 14).NumberFormatLocal = "#,##0.00_ "

            If (-row + 16) = 0 Then
                excelSheet.Application.Cells(row, 14).Value = 0
            Else
                excelSheet.Application.Cells(row, 14).FormulaR1C1 = "=SUM(R[-1]C:R[" & -row + 16 & "]C)"
            End If

            excelSheet.Application.Cells(row, 15).HorizontalAlignment = xlCenter
            excelSheet.Application.Cells(row, 15).VerticalAlignment = xlBottom
            excelSheet.Application.Cells(row, 15).FormulaR1C1 = "=IF(RC[-5]=RC[-1],""○"",""×"")"
            
            '2010.05.13
    
    
    End If

    Estimate_SHIZAI_Proc = False

Debug.Print "out Estimate_shizai_Proc=" & Now

End Function

Private Function Estimate_DOUKON_Proc(excelApplication As Object, excelWorkBook As Object, excelSheet As Object, row As Integer) As Integer                 '2013.09.30 EXCEL Ver対応
'Private Function Estimate_DOUKON_Proc(excelApplication As Excel.Application, excelWorkBook As Excel.Workbook, excelSheet As Excel.Worksheet, row As Integer) As Integer
'----------------------------------------------------------------------------
'                   ＥＸＣＥＬ（御見積書 同梱）出力
'       2009.06.02
'----------------------------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim k           As Integer
Dim l           As Integer


Dim com         As Integer
Dim sts         As Integer

Dim start_row   As Integer


    Estimate_DOUKON_Proc = True
    row = row + 2

'---    29行目
    excelSheet.Application.Cells(row, 2).Font.Size = 10
    excelSheet.Application.Cells(row, 2).Value = "【同梱部品明細】"
    
'---    同梱部品欄
    row = row + 1
        
        
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 3)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 3)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 3)).MergeCells = True
    

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row, 7)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row, 7)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row, 7)).MergeCells = True

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).MergeCells = True
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
    
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).LineStyle = xlNone
    
    
    
    
    
    
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 8)).HorizontalAlignment = xlCenter
    excelSheet.Application.Cells(row, 2).Value = "同梱品番"
    excelSheet.Application.Cells(row, 4).Value = "種別"
    excelSheet.Application.Cells(row, 5).Value = "品名"
    excelSheet.Application.Cells(row, 8).Value = "数量"
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 8)).Font.Size = 10
    
    start_row = row
'---    31〜40行目
    If EX_DOUKON_F Then
        
            
        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHin_Gai).Text)
           
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_DOUKON)
        Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
        
        com = BtOpGetGreaterEqual
            
        Do
           
            DoEvents
           
            sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
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
                    Call File_Error(sts, com, "構成マスタ")
                    Exit Function
            End Select
            
        
            If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_DOUKON Then
                Exit Do
            End If
        
        
            For j = 0 To UBound(EX_DOUKON_T)
                If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = EX_DOUKON_T(j) Then
                    
                    
                    
                    
                    row = row + 1

                    
                    
                    
                    excelSheet.Application.Cells(row, 2).Value = Trim(StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    
                    
                    
                    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
                    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
                    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                    Select Case sts
                        Case BtNoErr
                            excelSheet.Application.Cells(row, 4).Value = Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode))
                            
                        Case BtErrKeyNotFound
                        Case Else
                            Call File_Error(sts, com, "コードマスタ")
                            Exit Function
                    End Select
                    
                    
                    
                    
                    If IsNumeric(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) Then
                        excelSheet.Application.Cells(row, 8).NumberFormatLocal = "#,##0_ "
                        excelSheet.Application.Cells(row, 8).HorizontalAlignment = xlCenter
                        excelSheet.Application.Cells(row, 8).Value = CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))
                    End If
                
                
                
                
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                            
                            
                            
                            excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row, 7)).HorizontalAlignment = xlLeft
                            excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row, 7)).VerticalAlignment = xlBottom
                            excelSheet.Application.Range(excelSheet.Application.Cells(row, 5), excelSheet.Application.Cells(row, 7)).MergeCells = True
                            
                            
                            excelSheet.Application.Cells(row, 5).Value = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                            '2009.07.06
                            excelSheet.Application.Cells(row, 5).ShrinkToFit = True
                            
                            
                        Case BtErrKeyNotFound
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Exit Function
                    End Select
                
                
                
                End If
            
                com = BtOpGetNext
            
            
            Next j
        
        
        
        
        
        
        
        
        
        Loop
    
    
    
    
    End If
    
    If row <> start_row Then
    
    
    
    
            start_row = start_row + 1
    
    
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
        
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
        
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).LineStyle = xlNone
        
            If start_row <> row Then
                excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).Weight = xlThin
                excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
            End If

            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 4), excelSheet.Application.Cells(row, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 4), excelSheet.Application.Cells(row, 4)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 4), excelSheet.Application.Cells(row, 4)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 5), excelSheet.Application.Cells(row, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 5), excelSheet.Application.Cells(row, 5)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 5), excelSheet.Application.Cells(row, 5)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 8), excelSheet.Application.Cells(row, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 8), excelSheet.Application.Cells(row, 8)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 8), excelSheet.Application.Cells(row, 8)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 9), excelSheet.Application.Cells(row, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 9), excelSheet.Application.Cells(row, 9)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 9), excelSheet.Application.Cells(row, 9)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 10), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 10), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeLeft).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 10), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
    
    
    
    
    
    
    End If
    
    
    
    
    Estimate_DOUKON_Proc = False
End Function



Private Function Estimate_FUKA_Proc(excelApplication As Object, excelWorkBook As Object, excelSheet As Object, row As Integer, Fsw As Integer) As Integer  '2013.09.30 EXCEL Ver対応
'Private Function Estimate_FUKA_Proc(excelApplication As Excel.Application, excelWorkBook As Excel.Workbook, excelSheet As Excel.Worksheet, row As Integer) As Integer
'----------------------------------------------------------------------------
'                   ＥＸＣＥＬ（御見積書 付加作業）出力
'       2009.06.02
'----------------------------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim k           As Integer
Dim l           As Integer


Dim com         As Integer
Dim sts         As Integer

Dim start_row   As Integer

Dim wkNum1      As Currency
Dim wkNum2      As Currency


Debug.Print "in Estimate_FUKA_Proc=" & Now

    Estimate_FUKA_Proc = True
    row = row + 2

'---    29行目
    excelSheet.Application.Cells(row, 2).Font.Size = 10
    excelSheet.Application.Cells(row, 2).Value = "【付加作業費】"
    
'---    付加作業欄
    row = row + 1
        
'>>>>>>>>>>>    2017.10.17
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 9)).HorizontalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 9)).VerticalAlignment = xlBottom
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 9)).MergeCells = True
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).HorizontalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).VerticalAlignment = xlBottom
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).MergeCells = True
'
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).LineStyle = xlNone
'
'
'
'
'
'
'
'
'    excelSheet.Application.Cells(row, 2).Value = "作業内容"
'    excelSheet.Application.Cells(row, 10).Value = "工数(秒)"
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 10)).Font.Size = 10
'
'    start_row = row
'
'
''---    31〜40行目
'    If EX_FUKA_F Then
'
'
'        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
'        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
'        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
'        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHin_Gai).Text)
'
'        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_DOUKON)
'        Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
'
'        com = BtOpGetGreaterEqual
'
'        Do
'
'            DoEvents
'
'            sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
'            Select Case sts
'                Case BtNoErr
'
'
'                    If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
'                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
'                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
'                        Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHin_Gai).Text) Then
'
'                        Exit Do
'
'                    End If
'
'                Case BtErrEOF
'                    Exit Do
'                Case Else
'                    Call File_Error(sts, com, "構成マスタ")
'                    Exit Function
'            End Select
'
'
'            If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_DOUKON Then
'                Exit Do
'            End If
'
'
'            For j = 0 To UBound(EX_FUKA_T)
'                If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = EX_FUKA_T(j) Then
'
'
'
'
'                    row = row + 1
'
'
'                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
'                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
'                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
'
'                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                    Select Case sts
'                        Case BtNoErr
'
'                            If Not IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
'                                Call UniCode_Conv(ITEMREC.S_KOUSU, "00000000")
'                            End If
'
'
'
'                        Case BtErrKeyNotFound
'
'
'                            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
'                            Call UniCode_Conv(ITEMREC.S_KOUSU, "00000000")
'
'
'                        Case Else
'                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
'                            Exit Function
'                    End Select
'
'
'
'
'                    excelSheet.Application.Cells(row, 2).Value = Trim(StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)) & " " & _
'                                                                    Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)) & " " & _
'                                                                    Trim(StrConv(P_COMPO_K_REC.KO_BIKOU, vbUnicode))
'
'
'                    If IsNumeric(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) Then
''                        excelSheet.Application.Cells(row, 11).NumberFormatLocal = "#,##0_ "
'                        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).HorizontalAlignment = xlCenter
'                        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).VerticalAlignment = xlBottom
'                        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).MergeCells = True
'                        excelSheet.Application.Cells(row, 10).Value = CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode))
'                    End If
'
'
'
'                    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 10)).Font.Size = 11
'
'
'
'
'                End If
'
'                com = BtOpGetNext
'
'
'            Next j
'
'
'
'
'
'
'
'
'
'        Loop
'
'
'
'
'    End If
'
'    If row <> start_row Then
'
'
'
'
'
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).Weight = xlThin
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).Weight = xlThin
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).Weight = xlThin
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).Weight = xlThin
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
'
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).LineStyle = xlNone
'
'        If row = start_row + 1 Then
'            excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).LineStyle = xlNone
'        Else
'            excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'            excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).Weight = xlThin
'            excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 2), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
'        End If
'
'
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 10), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 10), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeLeft).Weight = xlThin
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row + 1, 10), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'
'
'
'
'
'
'
'    End If
'
''---    付加作業欄（見出し）
'    row = row + 1
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 6), excelSheet.Application.Cells(row, 8)).HorizontalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 6), excelSheet.Application.Cells(row, 8)).VerticalAlignment = xlBottom
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 6), excelSheet.Application.Cells(row, 8)).Font.Size = 10
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 6), excelSheet.Application.Cells(row, 8)).MergeCells = True
'    excelSheet.Application.Cells(row, 6).Value = "作業時間計(秒/個)"
'    excelSheet.Application.Cells(row, 6).ShrinkToFit = True
'
'
''    excelSheet.Application.Cells(row, 8).HorizontalAlignment = xlCenter
''    excelSheet.Application.Cells(row, 8).VerticalAlignment = xlBottom
''    excelSheet.Application.Cells(row, 8).Font.Size = 10
''    excelSheet.Application.Cells(row, 8).Value = "分/個"
'
'    excelSheet.Application.Cells(row, 9).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row, 9).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row, 9).Font.Size = 10
'    excelSheet.Application.Cells(row, 9).Value = "分レート"
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).HorizontalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).VerticalAlignment = xlBottom
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Font.Size = 12
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).MergeCells = True
'    excelSheet.Application.Cells(row, 10).Value = "�A付加作業費"
'
''---    付加作業欄（内容）
'    row = row + 1
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 6), excelSheet.Application.Cells(row, 8)).HorizontalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 6), excelSheet.Application.Cells(row, 8)).VerticalAlignment = xlBottom
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 6), excelSheet.Application.Cells(row, 8)).Font.Size = 12
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 6), excelSheet.Application.Cells(row, 8)).MergeCells = True
'    If (row - 2) = start_row Then
'        excelSheet.Application.Cells(row, 6).Value = 0
'    Else
'        excelSheet.Application.Cells(row, 6).FormulaR1C1 = "=SUM(R[-2]C[4]:R[" & start_row - row + 1 & "]C[4]"
'    End If
'
''    excelSheet.Application.Cells(row, 8).HorizontalAlignment = xlCenter
''    excelSheet.Application.Cells(row, 8).VerticalAlignment = xlBottom
''    excelSheet.Application.Cells(row, 8).Font.Size = 12
''    excelSheet.Application.Cells(row, 8).FormulaR1C1 = "=round(RC[-2]/60,2)"
'
'
'    excelSheet.Application.Cells(row, 9).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row, 9).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row, 9).Font.Size = 12
'    excelSheet.Application.Cells(row, 9).Value = Text1(ptxBEF_SEI_RATE).Text
'
'
'    excelSheet.Application.Cells(row, 10).FormulaR1C1 = "=round(RC[-4]/60*RC[-1],2)"
'
'
'
'    If IsNumeric(excelSheet.Application.Cells(row, 10).Value) Then
'        wkNum1 = CCur(excelSheet.Application.Cells(row, 10).Value)
'    Else
'        wkNum1 = 0
'    End If
'
'
'    If IsNumeric(Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text) Then
'        wkNum2 = CCur(Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text)
'    Else
'        wkNum2 = 0
'    End If
'
'    If wkNum1 <> wkNum2 Then
'        MsgBox "�A付加作業費が計算値(分/個×分レート)と異なります。"
'        excelSheet.Application.Cells(row, 13).Value = "�A付加作業費が計算値(分/個×分レート)と異なります。"
'    End If
'
'
'
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).HorizontalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).VerticalAlignment = xlBottom
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Font.Size = 14
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).MergeCells = True
'    excelSheet.Application.Cells(row, 10).Value = Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text
'    excelSheet.Application.Cells(row, 10).NumberFormatLocal = "#,##0.00_ "
'
'
'
''2010.05.13
'    excelSheet.Application.Cells(row - 1, 14).Font.Size = 12
'    excelSheet.Application.Cells(row - 1, 14).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row - 1, 14).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row - 1, 14).Value = "単価"
'
'    excelSheet.Application.Cells(row, 14).FormulaR1C1 = "=round(RC[-8]/60*RC[-5],2)"
'    excelSheet.Application.Cells(row, 14).NumberFormatLocal = "#,##0.00_ "
'    excelSheet.Application.Cells(row, 14).HorizontalAlignment = xlRight
'    excelSheet.Application.Cells(row, 14).VerticalAlignment = xlBottom
'
'
'    excelSheet.Application.Cells(row - 1, 15).Font.Size = 12
'    excelSheet.Application.Cells(row - 1, 15).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row - 1, 15).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row - 1, 15).Value = "チェック"
'
'    excelSheet.Application.Cells(row, 15).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row, 15).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row, 15).FormulaR1C1 = "=IF(RC[-5]=RC[-1],""○"",""×"")"
'
'
''2010.05.13
'
'
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeLeft).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeBottom).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeRight).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlInsideVertical).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlInsideHorizontal).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
'
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).Weight = xlThick
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).Weight = xlThick
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).Weight = xlThick
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).Weight = xlThick
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).Weight = xlThick
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).Weight = xlThick
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
'>>>>>>>>>>>    2017.10.17
    
    
'>>>>>>>>>> 2017.10.30
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).MergeCells = True


    excelSheet.Application.Range(excelSheet.Application.Cells(row + 1, 2), excelSheet.Application.Cells(row + 1, 7)).HorizontalAlignment = xlLeft
    excelSheet.Application.Range(excelSheet.Application.Cells(row + 1, 2), excelSheet.Application.Cells(row + 1, 7)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(row + 1, 2), excelSheet.Application.Cells(row + 1, 7)).MergeCells = True


    
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlDiagonalUp).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeLeft).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeBottom).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeRight).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlInsideVertical).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlInsideVertical).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlInsideHorizontal).LineStyle = xlNone
    
    
    
    
    
    
    
    excelSheet.Application.Cells(row, 2).Value = "作業内容"
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 10)).Font.Size = 10
    
'---    付加作業欄（見出し）
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).VerticalAlignment = xlBottom
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Font.Size = 12
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).MergeCells = True
    excelSheet.Application.Cells(row, 10).Value = "�A付加作業費"
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlDiagonalUp).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
       
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlInsideVertical).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlInsideVertical).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlInsideVertical).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlInsideHorizontal).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic


    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
    
    
    
    start_row = row

    
'---    31〜40行目
    If EX_FUKA_F Then
        
            
        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHin_Gai).Text)
           
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_DOUKON)
        Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
        
        com = BtOpGetGreaterEqual
            
        Do
           
            DoEvents
           
            sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
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
                    Call File_Error(sts, com, "構成マスタ")
                    Exit Function
            End Select
            
        
            If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_DOUKON Then
                Exit Do
            End If
        
        
            For j = 0 To UBound(EX_FUKA_T)
                If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = EX_FUKA_T(j) Then
                    
                    
                    
                    

                    
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                            
                            If Not IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
                                Call UniCode_Conv(ITEMREC.S_KOUSU, "00000000")
                            End If
                            
                            
                            
                        Case BtErrKeyNotFound
                        
                        
                            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                            Call UniCode_Conv(ITEMREC.S_KOUSU, "00000000")
                        
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Exit Function
                    End Select
                    
                    row = row + 1
                                        
                    
                    excelSheet.Application.Cells(row, 2).Value = Trim(StrConv(P_COMPO_K_REC.KO_BIKOU, vbUnicode))
    '2018.01.05                excelSheet.Application.Cells(row, 10).Value = CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode))
                    
                
                

                
                
                
                End If
            
                com = BtOpGetNext
            
            
            Next j
        
        
        
        
        
        
        
        
        
        Loop
    
'>>>>>>>>>>>>>>> 2018.01.05
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlDiagonalDown).LineStyle = xlNone
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeLeft).Weight = xlThin
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeTop).LineStyle = xlContinuous
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeTop).Weight = xlThin
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeBottom).Weight = xlThin
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeRight).Weight = xlThin
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
'
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlInsideVertical).LineStyle = xlContinuous
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlInsideVertical).Weight = xlThin
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
'
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlInsideHorizontal).Weight = xlThin
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
'
'
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).Weight = xlThin
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).Weight = xlThin
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).Weight = xlThin
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).Weight = xlThin
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
'
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).Weight = xlThin
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
'
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).Weight = xlThin
'        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
'>>>>>>>>>>>>>>> 2018.01.05
    
    End If
'>>>>>>>>>>>>>>> 2018.01.05
''---    付加作業欄（内容）
'    row = row + 1
'    If (row - 3) = start_row Then
'    If (row - 1 = start_row) Then
'        excelSheet.Application.Cells(row - 1, 6).Value = 0
'        row = row + 1       '2017.11.07
'    Else
'        excelSheet.Application.Cells(row - 1, 6).FormulaR1C1 = "=SUM(R[-2]C[4]:R[" & start_row - row + 1 & "]C[4]"
''    End If '2017.11.07
'
'        excelSheet.Application.Cells(row - 1, 8).HorizontalAlignment = xlCenter
'        excelSheet.Application.Cells(row - 1, 8).VerticalAlignment = xlBottom
'        excelSheet.Application.Cells(row - 1, 8).Font.Size = 12
'        excelSheet.Application.Cells(row - 1, 8).FormulaR1C1 = "=round(RC[-2]/60,2)"
'
'
'        excelSheet.Application.Cells(row - 1, 9).HorizontalAlignment = xlCenter
'        excelSheet.Application.Cells(row - 1, 9).VerticalAlignment = xlBottom
'        excelSheet.Application.Cells(row - 1, 9).Font.Size = 12
'        excelSheet.Application.Cells(row - 1, 9).Value = Text1(ptxBEF_SEI_RATE).Text
'
'
'        excelSheet.Application.Cells(row - 1, 10).FormulaR1C1 = "=round(RC[-4]/60*RC[-1],2)"
'>>>>>>>>>>>>>>> 2018.01.05
    
    
    
'>>>>>>>>>> 2017.10.30
'    If IsNumeric(excelSheet.Application.Cells(row, 10).Value) Then
'        wkNum1 = CCur(excelSheet.Application.Cells(row, 10).Value)
'    Else
'        wkNum1 = 0
'    End If
    
    
'    If IsNumeric(Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text) Then
'        wkNum2 = CCur(Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text)
'    Else
'        wkNum2 = 0
'    End If
    
'    If wkNum1 <> wkNum2 Then
'        MsgBox "�A付加作業費が計算値(分/個×分レート)と異なります。"
'        excelSheet.Application.Cells(row, 13).Value = "�A付加作業費が計算値(分/個×分レート)と異なります。"
'    End If
'>>>>>>>>>> 2017.10.30
    
    
    
    '>>>>>>>>>>>> 2018.01.05
        If start_row = row Then
            Fsw = False
            Estimate_FUKA_Proc = False
    
            Exit Function
        End If
        Fsw = True
    '>>>>>>>>>>>> 2018.01.05
        
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).HorizontalAlignment = xlCenter
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).VerticalAlignment = xlBottom
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Font.Size = 14
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).MergeCells = True
        excelSheet.Application.Cells(row, 10).Value = Text1(ptxBEF_S_PPSC_KAKO_KOSU).Text
        excelSheet.Application.Cells(row, 10).NumberFormatLocal = "#,##0.00_ "
        excelSheet.Application.Cells(row, 8).Value = ""
        excelSheet.Application.Cells(row, 9).HorizontalAlignment = xlRight
        excelSheet.Application.Cells(row, 9).Value = "付加作業費一式"
    
    
    '>>>>>>>>>>>> 2018.01.05
        
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 2)).Borders(xlEdgeLeft).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 2), excelSheet.Application.Cells(row, 2)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 10), excelSheet.Application.Cells(start_row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 10), excelSheet.Application.Cells(start_row, 11)).Borders(xlEdgeTop).Weight = xlThin
'        excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 10), excelSheet.Application.Cells(start_row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeBottom).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 2), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 7), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 7), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeRight).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(start_row, 7), excelSheet.Application.Cells(row, 7)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
        
        
        For l = start_row + 1 To row
            excelSheet.Application.Range(excelSheet.Application.Cells(l, 2), excelSheet.Application.Cells(l, 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            excelSheet.Application.Range(excelSheet.Application.Cells(l, 2), excelSheet.Application.Cells(l, 7)).Borders(xlEdgeBottom).Weight = xlThin
            excelSheet.Application.Range(excelSheet.Application.Cells(l, 2), excelSheet.Application.Cells(l, 7)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        Next l
        
        
        
        
        
        
        
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeLeft).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
        
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 11), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 11), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).Weight = xlThin
        excelSheet.Application.Range(excelSheet.Application.Cells(row, 11), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
    '>>>>>>>>>>>> 2018.01.05
    
    
    
'>>>>>>>>>>>> 2018.01.05
'    End If  '2017.11.07
'>>>>>>>>>>>> 2018.01.05
    
    
'>>>>>>>>>> 2017.10.30
''2010.05.13
'    excelSheet.Application.Cells(row - 1, 14).Font.Size = 12
'    excelSheet.Application.Cells(row - 1, 14).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row - 1, 14).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row - 1, 14).Value = "単価"
'
'    excelSheet.Application.Cells(row, 14).FormulaR1C1 = "=round(RC[-8]/60*RC[-5],2)"
'    excelSheet.Application.Cells(row, 14).NumberFormatLocal = "#,##0.00_ "
'    excelSheet.Application.Cells(row, 14).HorizontalAlignment = xlRight
'    excelSheet.Application.Cells(row, 14).VerticalAlignment = xlBottom
'
'    excelSheet.Application.Cells(row - 1, 15).Font.Size = 12
'    excelSheet.Application.Cells(row - 1, 15).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row - 1, 15).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row - 1, 15).Value = "チェック"
'
'    excelSheet.Application.Cells(row, 15).HorizontalAlignment = xlCenter
'    excelSheet.Application.Cells(row, 15).VerticalAlignment = xlBottom
'    excelSheet.Application.Cells(row, 15).FormulaR1C1 = "=IF(RC[-5]=RC[-1],""○"",""×"")"
'
''2010.05.13
'>>>>>>>>>> 2017.10.30
    
    
    
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlDiagonalUp).LineStyle = xlNone
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeLeft).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeBottom).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeRight).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlInsideVertical).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlInsideHorizontal).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 6), excelSheet.Application.Cells(row, 10)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
    
    
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).Weight = xlThick
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).Weight = xlThick
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).Weight = xlThick
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).Weight = xlThick
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).Weight = xlThick
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).Weight = xlThick
'    excelSheet.Application.Range(excelSheet.Application.Cells(row - 1, 10), excelSheet.Application.Cells(row, 11)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
    
    
    
    
    
    Estimate_FUKA_Proc = False
Debug.Print row
Debug.Print "out Estimate_FUKA_Proc=" & Now

End Function


Private Sub Estimate_Line12_13_Proc(excelApplication As Object, excelWorkBook As Object, excelSheet As Object)                  '2013.09.30 EXCEL Ver対応
'Private Sub Estimate_Line12_13_Proc(excelApplication As Excel.Application, excelWorkBook As Excel.Workbook, excelSheet As Excel.Worksheet)
'----------------------------------------------------------------------------
'                   ＥＸＣＥＬ（御見積書 12-13行目）出力
'       2011.01.11
'----------------------------------------------------------------------------
    
    
    excelSheet.Application.Rows(12).RowHeight = 23.25
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 2)).HorizontalAlignment = xlLeft
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 2)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 2)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 2)).Font.Size = 14
    excelSheet.Application.Cells(12, 1).Value = "部品品番"

        
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 3)).HorizontalAlignment = xlLeft
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 3)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 3)).Font.Size = 16
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 3)).Font.NAME = "ＭＳ　ゴシック"
    excelSheet.Application.Cells(12, 3).Value = Trim(Text1(ptxHin_Gai).Text)
    
    
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlDiagonalUp).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeLeft).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeTop).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeBottom).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeRight).Weight = xlThick
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlInsideVertical).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 3), excelSheet.Application.Cells(12, 5)).Borders(xlInsideHorizontal).LineStyle = xlNone


    excelSheet.Application.Cells(12, 6).Font.Size = 10
    excelSheet.Application.Cells(12, 6).VerticalAlignment = xlCenter
    excelSheet.Application.Cells(12, 6).HorizontalAlignment = xlLeft
    excelSheet.Application.Cells(12, 6).Value = Trim(Text1(ptxHin_Name).Text)


'---    13行目
    excelSheet.Application.Rows(11).RowHeight = 13.5

End Sub

Private Sub txtTANTO_CODE_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sts     As Integer
        
        
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
        
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, txtTANTO_CODE.Text)

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            txtTanto_Name.Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            txtTanto_Name.Text = ""
    
            MsgBox "入力した項目はエラーです。(担当者)"
            txtTANTO_CODE.SetFocus
            Exit Sub
        Case Else
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Unload Me
            Exit Sub
    
    End Select

End Sub

Private Function Main_Update_Proc() As Integer

Dim sts     As Integer
Dim com     As Integer
Dim Fsw     As Integer

Dim i       As Integer
Dim Errflg  As Integer


    Main_Update_Proc = True
    
    
    
    
    Call UniCode_Conv(wK2_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(wK2_P_COMPO.KO_JGYOBU, SHIZAI)
    Call UniCode_Conv(wK2_P_COMPO.KO_NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(wK2_P_COMPO.KO_HIN_GAI, MAIN_HIN_GAI)
       
    Call UniCode_Conv(wK2_P_COMPO.DATA_KBN, P_HEAD)
    Call UniCode_Conv(wK2_P_COMPO.SEQNO, "000")
       
    Fsw = 0
    com = BtOpGetGreater
    
    Do
        DoEvents
        
        sts = BTRV(com, wP_COMPO_POS, wP_COMPO_K_REC, Len(wP_COMPO_K_REC), wK2_P_COMPO, Len(wK2_P_COMPO), 2)
        Select Case sts
            Case BtNoErr
            
                If StrConv(wP_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                    StrConv(wP_COMPO_K_REC.KO_JGYOBU, vbUnicode) <> SHIZAI Or _
                    StrConv(wP_COMPO_K_REC.KO_NAIGAI, vbUnicode) <> NAIGAI_NAI Or _
                    StrConv(wP_COMPO_K_REC.KO_HIN_GAI, vbUnicode) <> MAIN_HIN_GAI Then
                    
                    If Fsw = 0 Then
                        
                        List2.AddItem MAIN_HIN_GAI & " " & Space(20) & "NG"
Call LOG_OUT(SEI0019_LOG, "[NG]" & "KO_HIN_GAI_= " & MAIN_HIN_GAI & "HIN_GAI_= " & "該当なし" & " " & Now)
                        
                        NG_cnt = NG_cnt + 1
                        txtNG_CNT.Text = Format(NG_cnt, "#,##0")
                        DoEvents
                    
                    End If

                    Exit Do
                
                End If

                If StrConv(wP_COMPO_K_REC.DATA_KBN, vbUnicode) = "0" Or _
                    StrConv(wP_COMPO_K_REC.SEQNO, vbUnicode) = "000" Then
                Else

                    Fsw = 1
                    
                    Text1(ptxTanto_Code).Text = txtTANTO_CODE.Text
                    Text1(ptxHin_Gai).Text = StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode)
                    
                    
                    Combo1(pcmbSHIMUKE).ListIndex = cmbSHIMUKE.ListIndex

                    'ステータスウィンドウを作成する
                    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                        "[請求システム]見積書一括作成処理　子品番= " & MAIN_HIN_GAI & " 親品番= " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode), Me.hwnd, 0)



                    Errflg = False
                    If Detail_Disp_Proc(Errflg) Then
Call LOG_OUT(SEI0019_LOG, "[NG]" & "KO_HIN_GAI_= " & MAIN_HIN_GAI & "HIN_GAI_= " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode) & " " & Now)
                        Exit Function
                    End If
                        
                    
                    Errflg = False
                    For i = ptxTanto_Code To ptxAFT_S_BU_KAKO_KOSU
                    
                        If Error_Check_Proc(i) Then
                            Errflg = True
                            Exit For
                        End If
                    
                    
                    Next i
                    
                        
                        
                    If Not Errflg Then
                    
                        If TANKA_KEISAN_Proc() Then
Call LOG_OUT(SEI0019_LOG, "[NG]" & "KO_HIN_GAI_= " & MAIN_HIN_GAI & "HIN_GAI_= " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode) & " " & Now)
Call LOG_OUT(SEI0019_LOG, "見積書一括作成　異常終了[" & Now & "]")
                                
                            Exit Function
                        End If
                
                        If Tanka_Update_Proc() Then
Call LOG_OUT(SEI0019_LOG, "[NG]" & "KO_HIN_GAI_= " & MAIN_HIN_GAI & "HIN_GAI_= " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode) & " " & Now)
Call LOG_OUT(SEI0019_LOG, "見積書一括作成　異常終了[" & Now & "]")
                                    
                            Exit Function
                        End If
                    
                        If Detail_Disp_Proc(Errflg) Then
Call LOG_OUT(SEI0019_LOG, "見積書一括作成　異常終了[" & Now & "]")
                            Unload Me
                        End If
                        
                        If Estimate_Proc() Then
Call LOG_OUT(SEI0019_LOG, "[NG]" & "KO_HIN_GAI_= " & MAIN_HIN_GAI & "HIN_GAI_= " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode) & " " & Now)
Call LOG_OUT(SEI0019_LOG, "見積書一括作成　異常終了[" & Now & "]")
                            Exit Function
                        End If
                
                        OK_cnt = OK_cnt + 1
                        txtOK_CNT.Text = Format(OK_cnt, "#,##0")
                        DoEvents
                    
                        List2.AddItem MAIN_HIN_GAI & " " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode) & "OK"
Call LOG_OUT(SEI0019_LOG, "[OK]" & "KO_HIN_GAI_= " & MAIN_HIN_GAI & "HIN_GAI_= " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode) & " " & Now)
                    
                    Else
Call LOG_OUT(SEI0019_LOG, "[NG]" & "KO_HIN_GAI_= " & MAIN_HIN_GAI & "HIN_GAI_= " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode) & " " & Now)
                        
                        List2.AddItem MAIN_HIN_GAI & " " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode) & "NG"
                        NG_cnt = NG_cnt + 1
                        txtNG_CNT.Text = Format(NG_cnt, "#,##0")
                    End If
                
                End If
            
            Case BtErrEOF
                If Fsw = 0 Then
                    
Call LOG_OUT(SEI0019_LOG, "[NG]" & "KO_HIN_GAI_= " & MAIN_HIN_GAI & "HIN_GAI_= " & "該当なし" & " " & Now)
                    
                    NG_cnt = NG_cnt + 1
                    List2.AddItem MAIN_HIN_GAI & " " & Space(20) & "NG"
                    txtNG_CNT.Text = Format(NG_cnt, "#,##0")
                    
                    
                    DoEvents
               End If
            
            
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "構成マスタ")
Call LOG_OUT(SEI0019_LOG, "見積書一括作成　異常終了[" & Now & "]")
                Exit Function
                
    
        End Select
    
    
        com = BtOpGetNext
    
    Loop
    
    
    
    
    Main_Update_Proc = False
    


End Function




Private Function Main_Update_OYA_Proc() As Integer

Dim sts     As Integer
Dim com     As Integer
Dim Fsw     As Integer

Dim i       As Integer
Dim Errflg  As Integer


    Main_Update_OYA_Proc = True
    
                    
    Text1(ptxTanto_Code).Text = txtTANTO_CODE.Text
    Text1(ptxHin_Gai).Text = MAIN_HIN_GAI
    
    Combo1(pcmbSHIMUKE).ListIndex = cmbSHIMUKE.ListIndex
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]見積書一括作成処理　親品番= " & MAIN_HIN_GAI, Me.hwnd, 0)



    If Detail_Disp_Proc(Errflg) Then
        Call LOG_OUT(SEI0019_LOG, "[NG]" & "HIN_GAI_= " & MAIN_HIN_GAI & " " & Now)
        Exit Function
    End If
                        
                    
    Errflg = False
    For i = ptxTanto_Code To ptxAFT_S_BU_KAKO_KOSU
            
        If Error_Check_Proc(i) Then
            Errflg = True
            Exit For
        End If
            
            
    Next i
                    
                        
                        
    If Not Errflg Then
                    
        If TANKA_KEISAN_Proc() Then
            Call LOG_OUT(SEI0019_LOG, "[NG]" & "HIN_GAI_= " & MAIN_HIN_GAI & " " & Now)
            Exit Function
        End If
                
        If Tanka_Update_Proc() Then
            Call LOG_OUT(SEI0019_LOG, "[NG]" & "HIN_GAI_= " & MAIN_HIN_GAI & " " & Now)
            Exit Function
        End If
                    
        If Detail_Disp_Proc(Errflg) Then
            Unload Me
        End If
                        
        If Estimate_Proc() Then
            Call LOG_OUT(SEI0019_LOG, "[NG]" & "HIN_GAI_= " & MAIN_HIN_GAI & " " & Now)
            Exit Function
        End If
                
        OK_cnt = OK_cnt + 1
        txtOK_CNT.Text = Format(OK_cnt, "#,##0")
        DoEvents
                    
        List3.AddItem MAIN_HIN_GAI & "OK"
        Call LOG_OUT(SEI0019_LOG, "[OK]" & "HIN_GAI_= " & MAIN_HIN_GAI & " " & Now)
                    
    Else
        Call LOG_OUT(SEI0019_LOG, "[NG]" & "HIN_GAI_= " & MAIN_HIN_GAI & " " & Now)
                        
        List3.AddItem MAIN_HIN_GAI & "NG"
        NG_cnt = NG_cnt + 1
        txtNG_CNT.Text = Format(NG_cnt, "#,##0")
    End If
                
    
    
    
    
    
    
    Main_Update_OYA_Proc = False
    


End Function


Private Function COUNT_Proc() As Integer

Dim sts     As Integer
Dim com     As Integer
Dim Fsw     As Integer

Dim i       As Integer
Dim Errflg  As Integer


    COUNT_Proc = True
    
    Call UniCode_Conv(wK2_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(wK2_P_COMPO.KO_JGYOBU, SHIZAI)
    Call UniCode_Conv(wK2_P_COMPO.KO_NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(wK2_P_COMPO.KO_HIN_GAI, MAIN_HIN_GAI)
       
    Call UniCode_Conv(wK2_P_COMPO.DATA_KBN, P_HEAD)
    Call UniCode_Conv(wK2_P_COMPO.SEQNO, "000")
       
    Fsw = 0
    com = BtOpGetGreater
       
    
    Do
        DoEvents
        
        sts = BTRV(com, wP_COMPO_POS, wP_COMPO_K_REC, Len(wP_COMPO_K_REC), wK2_P_COMPO, Len(wK2_P_COMPO), 2)
        Select Case sts
            Case BtNoErr
            
                If StrConv(wP_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                    StrConv(wP_COMPO_K_REC.KO_JGYOBU, vbUnicode) <> SHIZAI Or _
                    StrConv(wP_COMPO_K_REC.KO_NAIGAI, vbUnicode) <> NAIGAI_NAI Or _
                    StrConv(wP_COMPO_K_REC.KO_HIN_GAI, vbUnicode) <> MAIN_HIN_GAI Then
                    

                    Exit Do
                
                End If

                If StrConv(wP_COMPO_K_REC.DATA_KBN, vbUnicode) = "0" Or _
                    StrConv(wP_COMPO_K_REC.SEQNO, vbUnicode) = "000" Then
                Else


                    List2.AddItem MAIN_HIN_GAI & " " & StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode)


                    IN_cnt = IN_cnt + 1
                    txtOUT_CNT.Text = Format(IN_cnt, "#,##0")
                
                                    
                End If
            
            Case BtErrEOF
            
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "構成マスタ")
                Exit Function
                
    
        End Select
    
    
        com = BtOpGetNext
    
    Loop
    
    
    
    
    COUNT_Proc = False

End Function



