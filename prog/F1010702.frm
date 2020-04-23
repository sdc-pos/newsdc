VERSION 5.00
Begin VB.Form F1010702 
   BackColor       =   &H00FFFFFF&
   Caption         =   "[作業管理マスタ]メニュー管理メンテナンス"
   ClientHeight    =   9315
   ClientLeft      =   2130
   ClientTop       =   2430
   ClientWidth     =   15180
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
   MaxButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   15180
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   2
      Top             =   840
      Width           =   375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      Height          =   7215
      Left            =   12120
      TabIndex        =   158
      Top             =   1320
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton cmdUpd03 
         Caption         =   "登"
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
         Index           =   8
         Left            =   2040
         TabIndex        =   197
         Top             =   6120
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd03 
         Caption         =   "登"
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
         Index           =   7
         Left            =   2040
         TabIndex        =   196
         Top             =   5400
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd03 
         Caption         =   "登"
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
         Index           =   6
         Left            =   2040
         TabIndex        =   195
         Top             =   4680
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd03 
         Caption         =   "登"
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
         Left            =   2040
         TabIndex        =   194
         Top             =   3960
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd03 
         Caption         =   "登"
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
         Left            =   2040
         TabIndex        =   193
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd03 
         Caption         =   "登"
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
         Left            =   2040
         TabIndex        =   192
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd03 
         Caption         =   "登"
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
         Left            =   2040
         TabIndex        =   191
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd03 
         Caption         =   "登"
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
         Left            =   2040
         TabIndex        =   190
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd03 
         Caption         =   "登"
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
         Left            =   2040
         TabIndex        =   189
         Top             =   480
         Width           =   375
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   6255
         LargeChange     =   9
         Left            =   2520
         Max             =   36
         Min             =   1
         SmallChange     =   9
         TabIndex        =   186
         Top             =   480
         Value           =   1
         Width           =   255
      End
      Begin VB.TextBox txtMTS_NAME 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   8
         Left            =   120
         TabIndex        =   185
         Top             =   6360
         Width           =   1095
      End
      Begin VB.TextBox txtSS_CODE 
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
         IMEMode         =   2  'ｵﾌ
         Index           =   8
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   184
         Top             =   6120
         Width           =   975
      End
      Begin VB.TextBox txtMTS_CODE 
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
         IMEMode         =   2  'ｵﾌ
         Index           =   8
         Left            =   120
         MaxLength       =   8
         TabIndex        =   183
         Top             =   6120
         Width           =   975
      End
      Begin VB.TextBox txtMTS_NAME 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   7
         Left            =   120
         TabIndex        =   182
         Top             =   5640
         Width           =   1095
      End
      Begin VB.TextBox txtSS_CODE 
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
         IMEMode         =   2  'ｵﾌ
         Index           =   7
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   181
         Top             =   5400
         Width           =   975
      End
      Begin VB.TextBox txtMTS_CODE 
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
         IMEMode         =   2  'ｵﾌ
         Index           =   7
         Left            =   120
         MaxLength       =   8
         TabIndex        =   180
         Top             =   5400
         Width           =   975
      End
      Begin VB.TextBox txtMTS_NAME 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   6
         Left            =   120
         TabIndex        =   179
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox txtSS_CODE 
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
         IMEMode         =   2  'ｵﾌ
         Index           =   6
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   178
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtMTS_CODE 
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
         IMEMode         =   2  'ｵﾌ
         Index           =   6
         Left            =   120
         MaxLength       =   8
         TabIndex        =   177
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtMTS_NAME 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   120
         TabIndex        =   176
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox txtSS_CODE 
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
         IMEMode         =   2  'ｵﾌ
         Index           =   5
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   175
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox txtMTS_CODE 
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
         IMEMode         =   2  'ｵﾌ
         Index           =   5
         Left            =   120
         MaxLength       =   8
         TabIndex        =   174
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox txtMTS_NAME 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   120
         TabIndex        =   173
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox txtSS_CODE 
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
         IMEMode         =   2  'ｵﾌ
         Index           =   4
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   172
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtMTS_CODE 
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
         IMEMode         =   2  'ｵﾌ
         Index           =   4
         Left            =   120
         MaxLength       =   8
         TabIndex        =   171
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtMTS_NAME 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   120
         TabIndex        =   170
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtSS_CODE 
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
         IMEMode         =   2  'ｵﾌ
         Index           =   3
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   169
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtMTS_CODE 
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
         IMEMode         =   2  'ｵﾌ
         Index           =   3
         Left            =   120
         MaxLength       =   8
         TabIndex        =   168
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtMTS_NAME 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   120
         TabIndex        =   167
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtSS_CODE 
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
         IMEMode         =   2  'ｵﾌ
         Index           =   2
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   166
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtMTS_CODE 
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
         IMEMode         =   2  'ｵﾌ
         Index           =   2
         Left            =   120
         MaxLength       =   8
         TabIndex        =   165
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtMTS_NAME 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   164
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtSS_CODE 
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
         IMEMode         =   2  'ｵﾌ
         Index           =   1
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   163
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtMTS_CODE 
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
         IMEMode         =   2  'ｵﾌ
         Index           =   1
         Left            =   120
         MaxLength       =   8
         TabIndex        =   162
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtSS_CODE 
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
         IMEMode         =   2  'ｵﾌ
         Index           =   0
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   161
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtMTS_NAME 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   160
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtMTS_CODE 
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
         IMEMode         =   2  'ｵﾌ
         Index           =   0
         Left            =   120
         MaxLength       =   8
         TabIndex        =   159
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "SS"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   188
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "MTS"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   187
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   6375
      LargeChange     =   9
      Left            =   11760
      Max             =   36
      Min             =   1
      SmallChange     =   9
      TabIndex        =   117
      Top             =   1680
      Value           =   1
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame FrmMENU02 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   8
      Left            =   6840
      TabIndex        =   113
      Top             =   7320
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdSel02 
         Caption         =   "選"
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
         Index           =   8
         Left            =   4440
         TabIndex        =   135
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd02 
         Caption         =   "登"
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
         Index           =   8
         Left            =   4080
         TabIndex        =   153
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtMenu02 
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
         Index           =   8
         Left            =   720
         MaxLength       =   10
         TabIndex        =   115
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkMENU02_D 
         BackColor       =   &H80000005&
         Caption         =   "削"
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
         Index           =   8
         Left            =   120
         TabIndex        =   114
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cboMENU02 
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
         Index           =   8
         ItemData        =   "F1010702.frx":0000
         Left            =   2760
         List            =   "F1010702.frx":0002
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   116
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FrmMENU02 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   7
      Left            =   6840
      TabIndex        =   109
      Top             =   6600
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdSel02 
         Caption         =   "選"
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
         Index           =   7
         Left            =   4440
         TabIndex        =   134
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd02 
         Caption         =   "登"
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
         Index           =   7
         Left            =   4080
         TabIndex        =   152
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtMenu02 
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
         Index           =   7
         Left            =   720
         MaxLength       =   10
         TabIndex        =   111
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkMENU02_D 
         BackColor       =   &H80000005&
         Caption         =   "削"
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
         Index           =   7
         Left            =   120
         TabIndex        =   110
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cboMENU02 
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
         Index           =   7
         ItemData        =   "F1010702.frx":0004
         Left            =   2760
         List            =   "F1010702.frx":0006
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   112
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FrmMENU02 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   6
      Left            =   6840
      TabIndex        =   105
      Top             =   5880
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdSel02 
         Caption         =   "選"
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
         Index           =   6
         Left            =   4440
         TabIndex        =   133
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd02 
         Caption         =   "登"
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
         Index           =   6
         Left            =   4080
         TabIndex        =   151
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtMenu02 
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
         Index           =   6
         Left            =   720
         MaxLength       =   10
         TabIndex        =   107
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkMENU02_D 
         BackColor       =   &H80000005&
         Caption         =   "削"
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
         Index           =   6
         Left            =   120
         TabIndex        =   106
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cboMENU02 
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
         Index           =   6
         ItemData        =   "F1010702.frx":0008
         Left            =   2760
         List            =   "F1010702.frx":000A
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   108
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FrmMENU02 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   5
      Left            =   6840
      TabIndex        =   101
      Top             =   5160
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdSel02 
         Caption         =   "選"
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
         Left            =   4440
         TabIndex        =   132
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd02 
         Caption         =   "登"
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
         Left            =   4080
         TabIndex        =   150
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtMenu02 
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
         Left            =   720
         MaxLength       =   10
         TabIndex        =   103
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkMENU02_D 
         BackColor       =   &H80000005&
         Caption         =   "削"
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
         Left            =   120
         TabIndex        =   102
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cboMENU02 
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
         Index           =   5
         ItemData        =   "F1010702.frx":000C
         Left            =   2760
         List            =   "F1010702.frx":000E
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   104
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FrmMENU02 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   4
      Left            =   6840
      TabIndex        =   97
      Top             =   4440
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdSel02 
         Caption         =   "選"
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
         Left            =   4440
         TabIndex        =   131
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd02 
         Caption         =   "登"
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
         Left            =   4080
         TabIndex        =   149
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtMenu02 
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
         Left            =   720
         MaxLength       =   10
         TabIndex        =   99
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkMENU02_D 
         BackColor       =   &H80000005&
         Caption         =   "削"
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
         Left            =   120
         TabIndex        =   98
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cboMENU02 
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
         Index           =   4
         ItemData        =   "F1010702.frx":0010
         Left            =   2760
         List            =   "F1010702.frx":0012
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   100
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FrmMENU02 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   3
      Left            =   6840
      TabIndex        =   93
      Top             =   3720
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdSel02 
         Caption         =   "選"
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
         Left            =   4440
         TabIndex        =   130
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd02 
         Caption         =   "登"
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
         Left            =   4080
         TabIndex        =   148
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtMenu02 
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
         Left            =   720
         MaxLength       =   10
         TabIndex        =   95
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkMENU02_D 
         BackColor       =   &H80000005&
         Caption         =   "削"
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
         Left            =   120
         TabIndex        =   94
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cboMENU02 
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
         Index           =   3
         ItemData        =   "F1010702.frx":0014
         Left            =   2760
         List            =   "F1010702.frx":0016
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   96
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FrmMENU02 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   2
      Left            =   6840
      TabIndex        =   89
      Top             =   3000
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdSel02 
         Caption         =   "選"
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
         Left            =   4440
         TabIndex        =   129
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd02 
         Caption         =   "登"
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
         Left            =   4080
         TabIndex        =   147
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtMenu02 
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
         Left            =   720
         MaxLength       =   10
         TabIndex        =   91
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkMENU02_D 
         BackColor       =   &H80000005&
         Caption         =   "削"
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
         Left            =   120
         TabIndex        =   90
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cboMENU02 
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
         Index           =   2
         ItemData        =   "F1010702.frx":0018
         Left            =   2760
         List            =   "F1010702.frx":001A
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   92
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FrmMENU02 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   1
      Left            =   6840
      TabIndex        =   85
      Top             =   2280
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdSel02 
         Caption         =   "選"
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
         Left            =   4440
         TabIndex        =   128
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd02 
         Caption         =   "登"
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
         Left            =   4080
         TabIndex        =   146
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtMenu02 
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
         Left            =   720
         MaxLength       =   10
         TabIndex        =   87
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkMENU02_D 
         BackColor       =   &H80000005&
         Caption         =   "削"
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
         Left            =   120
         TabIndex        =   86
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cboMENU02 
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
         Index           =   1
         ItemData        =   "F1010702.frx":001C
         Left            =   2760
         List            =   "F1010702.frx":001E
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   88
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FrmMENU02 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   0
      Left            =   6840
      TabIndex        =   81
      Top             =   1560
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdSel02 
         Caption         =   "選"
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
         Left            =   4440
         TabIndex        =   127
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd02 
         Caption         =   "登"
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
         Left            =   4080
         TabIndex        =   145
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox chkMENU02_D 
         BackColor       =   &H80000005&
         Caption         =   "削"
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
         Left            =   120
         TabIndex        =   84
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtMenu02 
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
         Left            =   720
         MaxLength       =   10
         TabIndex        =   83
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cboMENU02 
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
         ItemData        =   "F1010702.frx":0020
         Left            =   2760
         List            =   "F1010702.frx":0022
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   82
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6375
      LargeChange     =   9
      Left            =   6480
      Max             =   36
      Min             =   1
      SmallChange     =   9
      TabIndex        =   80
      Top             =   1680
      Value           =   1
      Width           =   255
   End
   Begin VB.Frame FrmMENU01 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   8
      Left            =   0
      TabIndex        =   73
      Top             =   7320
      Width           =   6495
      Begin VB.CommandButton cmdSel01 
         Caption         =   "選"
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
         Index           =   8
         Left            =   6000
         TabIndex        =   126
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         Height          =   735
         Index           =   8
         Left            =   720
         TabIndex        =   76
         Top             =   0
         Width           =   1575
         Begin VB.OptionButton OptMENU01_M 
            BackColor       =   &H80000005&
            Caption         =   "MENU"
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
            Index           =   8
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton OptMENU01_S 
            BackColor       =   &H80000005&
            Caption         =   "作業"
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
            Index           =   8
            Left            =   840
            TabIndex        =   77
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox txtMenu01 
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
         Index           =   8
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   75
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cboMENU01 
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
         Index           =   8
         ItemData        =   "F1010702.frx":0024
         Left            =   4320
         List            =   "F1010702.frx":0026
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   74
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkMENU01_D 
         BackColor       =   &H80000005&
         Caption         =   "削"
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
         Index           =   8
         Left            =   120
         TabIndex        =   79
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdUpd01 
         Caption         =   "登"
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
         Index           =   8
         Left            =   5640
         TabIndex        =   144
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame FrmMENU01 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   7
      Left            =   0
      TabIndex        =   66
      Top             =   6600
      Width           =   6495
      Begin VB.CommandButton cmdSel01 
         Caption         =   "選"
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
         Index           =   7
         Left            =   6000
         TabIndex        =   125
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         Height          =   735
         Index           =   7
         Left            =   720
         TabIndex        =   69
         Top             =   0
         Width           =   1575
         Begin VB.OptionButton OptMENU01_S 
            BackColor       =   &H80000005&
            Caption         =   "作業"
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
            Index           =   7
            Left            =   840
            TabIndex        =   70
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptMENU01_M 
            BackColor       =   &H80000005&
            Caption         =   "MENU"
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
            Index           =   7
            Left            =   120
            TabIndex        =   71
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox txtMenu01 
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
         Index           =   7
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   68
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cboMENU01 
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
         Index           =   7
         ItemData        =   "F1010702.frx":0028
         Left            =   4320
         List            =   "F1010702.frx":002A
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   67
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkMENU01_D 
         BackColor       =   &H80000005&
         Caption         =   "削"
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
         Index           =   7
         Left            =   120
         TabIndex        =   72
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdUpd01 
         Caption         =   "登"
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
         Index           =   7
         Left            =   5640
         TabIndex        =   143
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame FrmMENU01 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   6
      Left            =   0
      TabIndex        =   59
      Top             =   5880
      Width           =   6495
      Begin VB.CommandButton cmdSel01 
         Caption         =   "選"
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
         Index           =   6
         Left            =   6000
         TabIndex        =   124
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         Height          =   735
         Index           =   6
         Left            =   720
         TabIndex        =   62
         Top             =   0
         Width           =   1575
         Begin VB.OptionButton OptMENU01_S 
            BackColor       =   &H80000005&
            Caption         =   "作業"
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
            Index           =   6
            Left            =   840
            TabIndex        =   63
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptMENU01_M 
            BackColor       =   &H80000005&
            Caption         =   "MENU"
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
            Index           =   6
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox txtMenu01 
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
         Index           =   6
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   61
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cboMENU01 
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
         Index           =   6
         ItemData        =   "F1010702.frx":002C
         Left            =   4320
         List            =   "F1010702.frx":002E
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   60
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkMENU01_D 
         BackColor       =   &H80000005&
         Caption         =   "削"
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
         Index           =   6
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdUpd01 
         Caption         =   "登"
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
         Index           =   6
         Left            =   5640
         TabIndex        =   142
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame FrmMENU01 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   5
      Left            =   0
      TabIndex        =   52
      Top             =   5160
      Width           =   6495
      Begin VB.CommandButton cmdSel01 
         Caption         =   "選"
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
         Left            =   6000
         TabIndex        =   123
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         Height          =   735
         Index           =   5
         Left            =   720
         TabIndex        =   55
         Top             =   0
         Width           =   1575
         Begin VB.OptionButton OptMENU01_S 
            BackColor       =   &H80000005&
            Caption         =   "作業"
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
            Left            =   840
            TabIndex        =   56
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptMENU01_M 
            BackColor       =   &H80000005&
            Caption         =   "MENU"
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
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox txtMenu01 
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   54
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cboMENU01 
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
         Index           =   5
         ItemData        =   "F1010702.frx":0030
         Left            =   4320
         List            =   "F1010702.frx":0032
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   53
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkMENU01_D 
         BackColor       =   &H80000005&
         Caption         =   "削"
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
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdUpd01 
         Caption         =   "登"
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
         Left            =   5640
         TabIndex        =   141
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame FrmMENU01 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   4
      Left            =   0
      TabIndex        =   45
      Top             =   4440
      Width           =   6495
      Begin VB.CommandButton cmdSel01 
         Caption         =   "選"
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
         Left            =   6000
         TabIndex        =   122
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtMenu01 
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   49
         Top             =   240
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         Height          =   735
         Index           =   4
         Left            =   720
         TabIndex        =   46
         Top             =   0
         Width           =   1575
         Begin VB.OptionButton OptMENU01_S 
            BackColor       =   &H80000005&
            Caption         =   "作業"
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
            Left            =   840
            TabIndex        =   47
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptMENU01_M 
            BackColor       =   &H80000005&
            Caption         =   "MENU"
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
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CheckBox chkMENU01_D 
         BackColor       =   &H80000005&
         Caption         =   "削"
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
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cboMENU01 
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
         Index           =   4
         ItemData        =   "F1010702.frx":0034
         Left            =   4320
         List            =   "F1010702.frx":0036
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   50
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpd01 
         Caption         =   "登"
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
         Left            =   5640
         TabIndex        =   140
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame FrmMENU01 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   3
      Left            =   0
      TabIndex        =   38
      Top             =   3720
      Width           =   6495
      Begin VB.CommandButton cmdSel01 
         Caption         =   "選"
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
         Left            =   6000
         TabIndex        =   121
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd01 
         Caption         =   "登"
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
         Left            =   5640
         TabIndex        =   139
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtMenu01 
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   42
         Top             =   240
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         Height          =   735
         Index           =   3
         Left            =   720
         TabIndex        =   39
         Top             =   0
         Width           =   1575
         Begin VB.OptionButton OptMENU01_S 
            BackColor       =   &H80000005&
            Caption         =   "作業"
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
            Left            =   840
            TabIndex        =   40
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptMENU01_M 
            BackColor       =   &H80000005&
            Caption         =   "MENU"
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
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CheckBox chkMENU01_D 
         BackColor       =   &H80000005&
         Caption         =   "削"
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
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cboMENU01 
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
         Index           =   3
         ItemData        =   "F1010702.frx":0038
         Left            =   4320
         List            =   "F1010702.frx":003A
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   43
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FrmMENU01 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   2
      Left            =   0
      TabIndex        =   31
      Top             =   3000
      Width           =   6495
      Begin VB.CommandButton cmdSel01 
         Caption         =   "選"
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
         Left            =   6000
         TabIndex        =   120
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd01 
         Caption         =   "登"
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
         Left            =   5640
         TabIndex        =   138
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         Height          =   735
         Index           =   2
         Left            =   720
         TabIndex        =   34
         Top             =   0
         Width           =   1575
         Begin VB.OptionButton OptMENU01_S 
            BackColor       =   &H80000005&
            Caption         =   "作業"
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
            TabIndex        =   35
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptMENU01_M 
            BackColor       =   &H80000005&
            Caption         =   "MENU"
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
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox txtMenu01 
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   33
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkMENU01_D 
         BackColor       =   &H80000005&
         Caption         =   "削"
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
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cboMENU01 
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
         Index           =   2
         ItemData        =   "F1010702.frx":003C
         Left            =   4320
         List            =   "F1010702.frx":003E
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   32
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FrmMENU01 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   1
      Left            =   0
      TabIndex        =   24
      Top             =   2280
      Width           =   6495
      Begin VB.CommandButton cmdSel01 
         Caption         =   "選"
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
         Left            =   6000
         TabIndex        =   119
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd01 
         Caption         =   "登"
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
         Left            =   5640
         TabIndex        =   137
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtMenu01 
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   29
         Top             =   240
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         Height          =   735
         Index           =   1
         Left            =   720
         TabIndex        =   25
         Top             =   0
         Width           =   1575
         Begin VB.OptionButton OptMENU01_S 
            BackColor       =   &H80000005&
            Caption         =   "作業"
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
            TabIndex        =   26
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptMENU01_M 
            BackColor       =   &H80000005&
            Caption         =   "MENU"
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
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CheckBox chkMENU01_D 
         BackColor       =   &H80000005&
         Caption         =   "削"
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
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cboMENU01 
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
         Index           =   1
         ItemData        =   "F1010702.frx":0040
         Left            =   4320
         List            =   "F1010702.frx":0042
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FrmMENU01 
      BackColor       =   &H80000005&
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   17
      Top             =   1560
      Width           =   6495
      Begin VB.CommandButton cmdSel01 
         Caption         =   "選"
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
         Left            =   6000
         TabIndex        =   118
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdUpd01 
         Caption         =   "登"
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
         Left            =   5640
         TabIndex        =   136
         Top             =   240
         Width           =   375
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         Height          =   735
         Index           =   0
         Left            =   720
         TabIndex        =   21
         Top             =   0
         Width           =   1575
         Begin VB.OptionButton OptMENU01_S 
            BackColor       =   &H80000005&
            Caption         =   "作業"
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
            Left            =   840
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptMENU01_M 
            BackColor       =   &H80000005&
            Caption         =   "MENU"
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
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CheckBox chkMENU01_D 
         BackColor       =   &H80000005&
         Caption         =   "削"
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
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtMenu01 
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   19
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cboMENU01 
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
         ItemData        =   "F1010702.frx":0044
         Left            =   4320
         List            =   "F1010702.frx":0046
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.ComboBox cboNAIGAI 
      Height          =   360
      ItemData        =   "F1010702.frx":0048
      Left            =   7320
      List            =   "F1010702.frx":004F
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox cboJIGYOBU 
      Height          =   360
      ItemData        =   "F1010702.frx":0061
      Left            =   5760
      List            =   "F1010702.frx":0068
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   1  'ｵﾝ
      Index           =   1
      Left            =   2760
      MaxLength       =   10
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton Command 
      Caption         =   "前画面"
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8520
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   8520
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   8520
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
      Left            =   7800
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   8520
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   8520
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   8520
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   8520
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   8520
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "Ｇ削除"
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   8520
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   8520
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   8520
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
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   8520
      Width           =   855
   End
   Begin VB.Label LblLEVEL02 
      Alignment       =   2  '中央揃え
      BackColor       =   &H80000005&
      Caption         =   "作業／要因"
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
      Index           =   2
      Left            =   9360
      TabIndex        =   157
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label LblLEVEL02 
      Alignment       =   2  '中央揃え
      BackColor       =   &H80000005&
      Caption         =   "表示名称"
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
      Left            =   7560
      TabIndex        =   156
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label LblLEVEL01 
      Alignment       =   2  '中央揃え
      BackColor       =   &H80000005&
      Caption         =   "表示名称"
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
      Left            =   2520
      TabIndex        =   154
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   7080
      X2              =   7200
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "メニューグループ"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label LblLEVEL01 
      Alignment       =   2  '中央揃え
      BackColor       =   &H80000005&
      Caption         =   "作業／要因"
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
      Index           =   2
      Left            =   4080
      TabIndex        =   155
      Top             =   1320
      Width           =   1695
   End
End
Attribute VB_Name = "F1010702"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const ptxMENU_GRP_No% = 0
Private Const ptxMENU_GRP% = 1
Private Const Text_Max% = 1


Private Sel_Code1 As String * 3
Private Sel_Code2 As String * 3

Private Type MENU_TBL_Tag
    Code    As String * 1
    NAME    As String * 4
    TYPE    As String * 1
    YOIN    As String * 1
End Type

Private MENU_TBL()  As MENU_TBL_Tag

Private Sel_CODE_TYPE   As String * 1

Private Const Base_Color& = &H80000005
Private Const Select_Color& = &H80FF80


Private Sub cboNAIGAI_KeyDown(KeyCode As Integer, Shift As Integer)
            
            If TBL_Set_Proc() Then
                Form_RTN = True
                F1010702.Hide
            End If

End Sub

Private Sub cmdSel01_Click(Index As Integer)
    
Dim i   As Integer
Dim sts As Integer
    
    If Len(Trim(Text(ptxMENU_GRP_No).Text)) = 0 Then
        Exit Sub
    End If
    
    If chkMENU01_D(Index).Value = True Then
        Exit Sub
    End If
    
    Sel_Code1 = Format((VScroll1.Value - 1) + Index, "000")
    
    Call UniCode_Conv(K0_tmpMENU.MENU_LV1, Sel_Code1)
    Call UniCode_Conv(K0_tmpMENU.MENU_LV2, "")
    Call UniCode_Conv(K0_tmpMENU.MENU_LV3, "")
    sts = BTRV(BtOpGetEqual, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), K0_tmpMENU, Len(K0_tmpMENU), 0)
    Select Case sts
        Case BtNoErr
            If StrConv(tmpMENUREC.MENU_KBN, vbUnicode) = "1" Then
                If StrConv(tmpMENUREC.PARAM_F, vbUnicode) <> "1" Then
                    Exit Sub
                End If
            End If
        Case BtErrKeyNotFound
            Exit Sub
        Case Else
            Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ(一時ファイル)")
            Form_RTN = True
            F1010702.Hide
    End Select
    
    For i = 0 To 8
        FrmMENU01(i).BackColor = Base_Color
    Next i
    
    Frame2.Visible = False
    
    If StrConv(tmpMENUREC.PARAM, vbUnicode) = "0" Then
        LblLEVEL02(1).Visible = False
        LblLEVEL02(2).Visible = False
        
        For i = 0 To 8
            FrmMENU02(i).Visible = False
        Next i
    
            
    
    Else
        FrmMENU01(Index).BackColor = Select_Color
        
        LblLEVEL02(1).Visible = True
        LblLEVEL02(2).Visible = True
        For i = 0 To 8
            FrmMENU02(i).Visible = True
        Next i
        
        VScroll2.Value = 1
        VScroll2.Visible = True
    
        Sel_CODE_TYPE = StrConv(tmpMENUREC.CODE_TYPE, vbUnicode) '実行中の要因をセット
    
        Call Display_R_Proc
    
    End If
End Sub

Private Sub cmdSel02_Click(Index As Integer)
    
Dim i   As Integer
Dim sts As Integer
    
    If Len(Trim(Text(ptxMENU_GRP_No).Text)) = 0 Then
        Exit Sub
    End If
    
    If chkMENU02_D(Index).Value = True Then
        Exit Sub
    End If
    
    Sel_Code2 = Format((VScroll2.Value - 1) + Index, "000")
    
    Call UniCode_Conv(K0_tmpMENU.MENU_LV1, Sel_Code1)
    Call UniCode_Conv(K0_tmpMENU.MENU_LV2, Sel_Code2)
    Call UniCode_Conv(K0_tmpMENU.MENU_LV3, "")
    sts = BTRV(BtOpGetEqual, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), K0_tmpMENU, Len(K0_tmpMENU), 0)
    Select Case sts
        Case BtNoErr
            If StrConv(tmpMENUREC.MENU_KBN, vbUnicode) = "1" Then
                If StrConv(tmpMENUREC.PARAM_F, vbUnicode) <> "1" Then
                    Exit Sub
                End If
            End If
        Case BtErrKeyNotFound
            Exit Sub
        Case Else
            Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ(一時ファイル)")
            Form_RTN = True
            F1010702.Hide
    End Select
    
    For i = 0 To 8
        FrmMENU02(i).BackColor = Base_Color
    Next i
    
    For i = 0 To 8
        txtMTS_CODE(i).Text = ""
        txtSS_CODE(i).Text = ""
        txtMTS_NAME(i).Text = ""
    Next i
    
    If StrConv(tmpMENUREC.PARAM_F, vbUnicode) <> "1" Then
        Frame2.Visible = False
    Else
        FrmMENU02(Index).BackColor = Select_Color
        
        VScroll3.Value = 1
        Frame2.Visible = True
    
        Call Display_MTS_Proc
    
    End If

End Sub

Private Sub cmdUpd01_Click(Index As Integer)

Dim Edit    As String * 2
Dim sts     As Integer
Dim com     As Integer

Dim ans     As Integer

    If Len(Trim(Text(ptxMENU_GRP_No).Text)) = 0 Then
        Beep
        MsgBox "メニューグループ№を入力して下さい。"
        Text(ptxMENU_GRP_No).SetFocus
        Exit Sub
    End If
    
    If chkMENU01_D(Index).Value Then
    Else
        If Not OptMENU01_M(Index).Value And Not OptMENU01_S(Index).Value Then
            Beep
            MsgBox "ＭＥＮＵ／作業の選択をして下さい。"
            Exit Sub
        End If

        If Len(Trim(txtMenu01(Index).Text)) = 0 Then
            Beep
            MsgBox "表示内容を入力して下さい。"
            Exit Sub
        End If

        If cboMENU01(Index).ListCount <= 0 Then
            Beep
            MsgBox "作業／要因が有りません。"
            Exit Sub
        End If
    
        If OptMENU01_S(Index).Value Then
            Edit = Right(cboMENU01(Index).Text, 2)
            Call UniCode_Conv(K0_YOIN.CODE_TYPE, Left(Edit, 1))
            Call UniCode_Conv(K0_YOIN.YOIN_CODE, Right(Edit, 1))
            sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Beep
                    MsgBox "作業／要因が有りません。選択内容を確認して下さい。"
                    Exit Sub
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ")
                    Form_RTN = True
                    F1010702.Hide
            End Select
    
    
        End If
    End If

    Do
        Call UniCode_Conv(K0_tmpMENU.MENU_LV1, Format((VScroll1.Value - 1) + Index, "000"))
        Call UniCode_Conv(K0_tmpMENU.MENU_LV2, "")
        Call UniCode_Conv(K0_tmpMENU.MENU_LV3, "")
        sts = BTRV(BtOpGetEqual + BtSNoWait, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), K0_tmpMENU, Len(K0_tmpMENU), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<tmpMENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Form_RTN = True
                    F1010702.Hide
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ(一時ファイル)")
                Form_RTN = True
                F1010702.Hide
        End Select
    Loop
                                                
    Call UniCode_Conv(tmpMENUREC.MENU_LV1, Format((VScroll1.Value - 1) + Index, "000"))
    Call UniCode_Conv(tmpMENUREC.MENU_LV2, "")
    Call UniCode_Conv(tmpMENUREC.MENU_LV3, "")
    
    
    If chkMENU01_D(Index).Value Then
        Call UniCode_Conv(tmpMENUREC.DEL_FLG, "1")
    Else
        Call UniCode_Conv(tmpMENUREC.DEL_FLG, "0")
    End If

    If OptMENU01_M(Index).Value Then
        Call UniCode_Conv(tmpMENUREC.MENU_KBN, "0")
    Else
        Call UniCode_Conv(tmpMENUREC.MENU_KBN, "1")
    End If
    
    Call UniCode_Conv(tmpMENUREC.DISPLAY_ITEM, Trim(txtMenu01(Index).Text))
    
    If OptMENU01_M(Index).Value Then
        Call UniCode_Conv(tmpMENUREC.CODE_TYPE, Right(cboMENU01(Index).Text, 1))
        Call UniCode_Conv(tmpMENUREC.YOIN_CODE, "")
        Call UniCode_Conv(tmpMENUREC.PARAM_F, "")
        Call UniCode_Conv(tmpMENUREC.PARAM, "")
    Else
        
        Edit = Right(cboMENU01(Index).Text, 2)
        
        Call UniCode_Conv(tmpMENUREC.CODE_TYPE, Left(Edit, 1))
        Call UniCode_Conv(tmpMENUREC.YOIN_CODE, Right(Edit, 1))
        Call UniCode_Conv(tmpMENUREC.PARAM_F, StrConv(YOINREC.PARAM_F, vbUnicode))
        Call UniCode_Conv(tmpMENUREC.PARAM, "")
    
    End If

    Do
        sts = BTRV(com, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), K0_tmpMENU, Len(K0_tmpMENU), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<tmpMENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Form_RTN = True
                    F1010702.Hide
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ(一時ファイル)")
                Form_RTN = True
                F1010702.Hide
        End Select
    Loop
End Sub

Private Sub cmdUpd02_Click(Index As Integer)

Dim Edit    As String * 2
Dim sts     As Integer
Dim com     As Integer

Dim ans     As Integer


    If Len(Trim(Text(ptxMENU_GRP_No).Text)) = 0 Then
        Beep
        MsgBox "メニューグループ№を入力して下さい。"
        Text(ptxMENU_GRP_No).SetFocus
        Exit Sub
    End If

    If chkMENU02_D(Index).Value Then
    Else

        If Len(Trim(txtMenu02(Index).Text)) = 0 Then
            Beep
            MsgBox "表示内容を入力して下さい。"
            Exit Sub
        End If

        If cboMENU02(Index).ListCount <= 0 Then
            Beep
            MsgBox "作業／要因が有りません。"
            Exit Sub
        End If
    
        Edit = Right(cboMENU02(Index).Text, 2)
        Call UniCode_Conv(K0_YOIN.CODE_TYPE, Left(Edit, 1))
        Call UniCode_Conv(K0_YOIN.YOIN_CODE, Right(Edit, 1))
        sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Beep
                MsgBox "作業／要因が有りません。選択内容を確認して下さい。"
                Exit Sub
            Case Else
                Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ")
                Form_RTN = True
                F1010702.Hide
        End Select
    
    
    End If

    Do
        Call UniCode_Conv(K0_tmpMENU.MENU_LV1, Sel_Code1)
        Call UniCode_Conv(K0_tmpMENU.MENU_LV2, Format((VScroll2.Value - 1) + Index, "000"))
        Call UniCode_Conv(K0_tmpMENU.MENU_LV3, "")
        sts = BTRV(BtOpGetEqual + BtSNoWait, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), K0_tmpMENU, Len(K0_tmpMENU), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<tmpMENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Form_RTN = True
                    F1010702.Hide
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ(一時ファイル)")
                Form_RTN = True
                F1010702.Hide
        End Select
    Loop

    Call UniCode_Conv(tmpMENUREC.MENU_LV1, Sel_Code1)
    Call UniCode_Conv(tmpMENUREC.MENU_LV2, Format((VScroll2.Value - 1) + Index, "000"))
    Call UniCode_Conv(tmpMENUREC.MENU_LV3, "")
    
    If chkMENU02_D(Index).Value Then
        Call UniCode_Conv(tmpMENUREC.DEL_FLG, "1")
    Else
        Call UniCode_Conv(tmpMENUREC.DEL_FLG, "0")
    End If
    
    If StrConv(YOINREC.PARAM_F, vbUnicode) = "1" Then
        Call UniCode_Conv(tmpMENUREC.MENU_KBN, "0")
    Else
        Call UniCode_Conv(tmpMENUREC.MENU_KBN, "1")
    End If
    Call UniCode_Conv(tmpMENUREC.DISPLAY_ITEM, txtMenu02(Index).Text)
    
    Edit = Right(cboMENU02(Index).Text, 2)
    Call UniCode_Conv(tmpMENUREC.CODE_TYPE, Left(Edit, 1))
    Call UniCode_Conv(tmpMENUREC.YOIN_CODE, Right(Edit, 1))
    Call UniCode_Conv(tmpMENUREC.PARAM_F, StrConv(YOINREC.PARAM_F, vbUnicode))
    
    If StrConv(YOINREC.PARAM_F, vbUnicode) = "2" Then
        Call UniCode_Conv(tmpMENUREC.PARAM, StrConv(YOINREC.Soko_No, vbUnicode))
    Else
        Call UniCode_Conv(tmpMENUREC.PARAM, "")
    End If
    
    Do
        sts = BTRV(com, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), K0_tmpMENU, Len(K0_tmpMENU), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<tmpMENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Form_RTN = True
                    F1010702.Hide
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ(一時ファイル)")
                Form_RTN = True
                F1010702.Hide
        End Select
    Loop

End Sub

Private Sub cmdUpd03_Click(Index As Integer)

Dim sts         As Integer
Dim com         As Integer

Dim ans         As Integer

Dim CODE_TYPE   As String * 1
Dim YOIN_CODE   As String * 1
        

    If Len(Trim(Text(ptxMENU_GRP_No).Text)) = 0 Then
        Beep
        MsgBox "メニューグループ№を入力して下さい。"
        Text(ptxMENU_GRP_No).SetFocus
        Exit Sub
    End If
    
    If Len(Trim(txtMTS_CODE(Index))) = 0 And _
        Len(Trim(txtSS_CODE(Index))) = 0 Then
    Else
        Call UniCode_Conv(K0_MTS.MUKE_CODE, txtMTS_CODE(Index))
        Call UniCode_Conv(K0_MTS.SS_CODE, txtSS_CODE(Index))
        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                txtMTS_NAME(Index).Text = ""
                Beep
                MsgBox "向け先が有りません。入力内容を確認して下さい。"
                Exit Sub
            Case Else
                Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                Form_RTN = True
                F1010702.Hide
        End Select
    End If
    '上位メニューの内容を獲得
    Call UniCode_Conv(K0_tmpMENU.MENU_LV1, Sel_Code1)
    Call UniCode_Conv(K0_tmpMENU.MENU_LV2, Sel_Code2)
    Call UniCode_Conv(K0_tmpMENU.MENU_LV3, "")
    sts = BTRV(BtOpGetEqual + BtSNoWait, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), K0_tmpMENU, Len(K0_tmpMENU), 0)
    Select Case sts
        Case BtNoErr
            CODE_TYPE = StrConv(tmpMENUREC.CODE_TYPE, vbUnicode)
            YOIN_CODE = StrConv(tmpMENUREC.YOIN_CODE, vbUnicode)
        Case BtErrKeyNotFound
            Beep
            MsgBox "上位メニューの設定に不具合が有ります。入力内容を確認して下さい。"
            Exit Sub
        Case Else
            Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ(一時ファイル)")
            Form_RTN = True
            F1010702.Hide
    End Select
    '未入力データも登録ボタンが押されたら登録しておく。最終更新時に削除する。
    
    Do
        Call UniCode_Conv(K0_tmpMENU.MENU_LV1, Sel_Code1)
        Call UniCode_Conv(K0_tmpMENU.MENU_LV2, Sel_Code2)
        Call UniCode_Conv(K0_tmpMENU.MENU_LV3, Format((VScroll3.Value - 1) + Index, "000"))
        sts = BTRV(BtOpGetEqual + BtSNoWait, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), K0_tmpMENU, Len(K0_tmpMENU), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<tmpMENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Form_RTN = True
                    F1010702.Hide
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ(一時ファイル)")
                Form_RTN = True
                F1010702.Hide
        End Select
    Loop

    Call UniCode_Conv(tmpMENUREC.MENU_LV1, Sel_Code1)
    Call UniCode_Conv(tmpMENUREC.MENU_LV2, Sel_Code2)
    Call UniCode_Conv(tmpMENUREC.MENU_LV3, Format((VScroll3.Value - 1) + Index, "000"))
    Call UniCode_Conv(tmpMENUREC.DEL_FLG, "0")
    Call UniCode_Conv(tmpMENUREC.MENU_KBN, "1")
    Call UniCode_Conv(tmpMENUREC.DISPLAY_ITEM, txtMTS_NAME(Index).Text)
    
    Call UniCode_Conv(tmpMENUREC.CODE_TYPE, CODE_TYPE)
    Call UniCode_Conv(tmpMENUREC.YOIN_CODE, YOIN_CODE)
    Call UniCode_Conv(tmpMENUREC.PARAM_F, "0")
    Call UniCode_Conv(tmpMENUREC.PARAM, StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode))

    Do
        sts = BTRV(com, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), K0_tmpMENU, Len(K0_tmpMENU), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<tmpMENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Form_RTN = True
                    F1010702.Hide
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ(一時ファイル)")
                Form_RTN = True
                F1010702.Hide
        End Select
    Loop


End Sub

Private Sub Command_Click(Index As Integer)
    
Dim ans As Integer

    Select Case Index
    
    
        Case 0
            If Len(Trim(Text(ptxMENU_GRP_No).Text)) = "" Then
                Beep
                MsgBox "入力した項目はエラーです。（必須入力）"
                Text(ptxMENU_GRP_No).SetFocus
                Exit Sub
            End If
        
            If Update_Proc() Then
                Form_RTN = True
                F1010702.Hide
                Exit Sub
            End If
            
            Text(ptxMENU_GRP_No).SetFocus
        
        Case 3
            If Len(Trim(Text(ptxMENU_GRP_No).Text)) = "" Then
                Beep
                MsgBox "入力した項目はエラーです。（必須入力）"
                Text(ptxMENU_GRP).SetFocus
                Exit Sub
            End If
        
            ans = MsgBox("グループ単位での削除を行いますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Delete_Proc() Then
                    Form_RTN = True
                    F1010702.Hide
                    Exit Sub
                End If
            End If
        
            Text(ptxMENU_GRP_No).SetFocus
        
        Case 11
            Call Clear_Proc
            Form_RTN = False
            F1010702.Hide
    
    End Select
End Sub

Private Sub Form_Activate()
    
    Text(ptxMENU_GRP_No).SetFocus

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

Dim i   As Integer
Dim c   As String * 128
    
    Form_RTN = True
    
    '作業タイプ設定
    i = 0
    Do
        If GetIni("ACTION", "ACTION_CD" & Format(i + 1, "00"), "SYS", c) Then
            Exit Do
        End If
        If Trim(c) = "NON" Then
            Exit Do
        End If
    
        ReDim Preserve MENU_TBL(i)
        MENU_TBL(i).Code = Trim(c)
        
        If GetIni("ACTION", "ACTION_NM" & Format(i + 1, "00"), "SYS", c) Then
            MsgBox "作業情報の獲得に失敗しました。" & "[F101070]" & "[ACTION_NM" & Format(i, "00") & "]"
            Exit Do
        End If
        MENU_TBL(i).NAME = Trim(c)
        
        If GetIni("ACTION", "ACTION_TYPE" & Format(i + 1, "00"), "SYS", c) Then
            MsgBox "作業情報の獲得に失敗しました。" & "[F101070]" & "[ACTION_TYPE" & Format(i, "00") & "]"
            Exit Do
        End If
        MENU_TBL(i).TYPE = Trim(c)
        
        If GetIni("ACTION", "ACTION_YOIN" & Format(i + 1, "00"), "SYS", c) Then
            MsgBox "作業情報の獲得に失敗しました。" & "[F101070]" & "[ACTION_YOIN" & Format(i, "00") & "]"
            Exit Do
        End If
        MENU_TBL(i).YOIN = Trim(c)
        i = i + 1
    Loop
    
    
    
    
    Form_RTN = False

End Sub


Private Sub OptMENU01_M_Click(Index As Integer)
    
Dim i   As Integer
    
    cboMENU01(Index).Clear
    
    For i = 0 To UBound(MENU_TBL)
        
        If MENU_TBL(i).TYPE = "0" Then
            cboMENU01(Index).AddItem MENU_TBL(i).NAME & "     " & MENU_TBL(i).Code
        End If
    
    Next i
    
    cboMENU01(Index).ListIndex = 0
    
    If Len(Trim(txtMenu01(Index).Text)) = 0 Then
        txtMenu01(Index).Text = Left(cboMENU01(Index).Text, Len(cboMENU01(Index).Text) - 1)
    End If
    
    cboMENU01(Index).SetFocus

End Sub

Private Sub OptMENU01_S_Click(Index As Integer)


Dim sts     As Integer
Dim com     As Integer

    cboMENU01(Index).Clear
    
    com = BtOpGetFirst
    
    Do
        
        sts = BTRV(com, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(YOINREC.REGI_F, vbUnicode) <= "1" Then
                    cboMENU01(Index).AddItem StrConv(YOINREC.YOIN_DNAME, vbUnicode) & "     " & StrConv(YOINREC.CODE_TYPE, vbUnicode) & StrConv(YOINREC.YOIN_CODE, vbUnicode)
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "要因マスタ")
                Form_RTN = True
                F1010702.Hide
        End Select
    
        com = BtOpGetNext
    
    Loop

    If cboMENU01(Index).ListCount > 0 Then
        cboMENU01(Index).ListIndex = 0
        
        If Len(Trim(txtMenu01(Index).Text)) = 0 Then
            txtMenu01(Index).Text = Left(cboMENU01(Index).Text, Len(cboMENU01(Index).Text) - 1)
        End If
        
        cboMENU01(Index).SetFocus
    End If


End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If


End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    
    If KeyCode <> vbKeyReturn Then Exit Sub

    Select Case Index
        
        Case ptxMENU_GRP_No
            If Len(Trim(Text(ptxMENU_GRP_No).Text)) = 0 Then
                Beep
                MsgBox "入力した項目はエラーです。（必須入力）"
                Text(ptxMENU_GRP).SetFocus
                Exit Sub
            End If
        
            Text(ptxMENU_GRP).Text = ""
                        
            
            If Not cboJIGYOBU.Enabled Then
            
                If TBL_Set_Proc() Then
                    Form_RTN = True
                    F1010702.Hide
                End If
            
            End If
                                                
            Text(ptxMENU_GRP).SetFocus
                        
    End Select
    
End Sub

Private Function Display_L_Proc() As Integer
'----------------------------------------------------------------------------
'                   メニュー内容表示
'----------------------------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim Now_Pos     As Integer
Dim com         As Integer
Dim sts         As Integer
Dim Pos         As Integer
    
    Display_L_Proc = True
    
    Pos = 0
    For i = (VScroll1.Value - 1) To (VScroll1.Value - 1) + 8
        DoEvents
        Call UniCode_Conv(K0_tmpMENU.MENU_LV1, Format(i, "000"))
        Call UniCode_Conv(K0_tmpMENU.MENU_LV2, "")
        Call UniCode_Conv(K0_tmpMENU.MENU_LV3, "")
        
        sts = BTRV(BtOpGetEqual, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), K0_tmpMENU, Len(K0_tmpMENU), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(tmpMENUREC.DEL_FLG, "")
                Call UniCode_Conv(tmpMENUREC.CODE_TYPE, "")
                Call UniCode_Conv(tmpMENUREC.MENU_KBN, "")
                Call UniCode_Conv(tmpMENUREC.DISPLAY_ITEM, "")
                Call UniCode_Conv(tmpMENUREC.CODE_TYPE, "")
                Call UniCode_Conv(tmpMENUREC.YOIN_CODE, "")
                Call UniCode_Conv(tmpMENUREC.PARAM_F, "")
                Call UniCode_Conv(tmpMENUREC.PARAM, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ(一時ファイル)")
                Exit Function
        End Select
        
        If Len(Trim(StrConv(tmpMENUREC.MENU_LV1, vbUnicode))) = 0 Then
            chkMENU01_D(Pos).Value = False
            OptMENU01_M(Pos).Value = False
            OptMENU01_S(Pos).Value = False
            txtMenu01(Pos).Text = ""
            cboMENU01(Pos).Clear
        Else
            If StrConv(tmpMENUREC.DEL_FLG, vbUnicode) = "1" Then
                chkMENU01_D(Pos).Value = True
            Else
                chkMENU01_D(Pos).Value = False
            End If
            
            cboMENU01(Pos).Clear
            Now_Pos = 0
            
            Select Case StrConv(tmpMENUREC.MENU_KBN, vbUnicode)
                Case "0"
                    OptMENU01_M(Pos).Value = True
                    OptMENU01_S(Pos).Value = False
            
                    For j = 0 To UBound(MENU_TBL)
        
                        If MENU_TBL(j).TYPE = "0" Then
                            cboMENU01(Pos).AddItem MENU_TBL(j).NAME & "     " & MENU_TBL(j).Code
                        End If
    
                        If MENU_TBL(j).Code = StrConv(tmpMENUREC.CODE_TYPE, vbUnicode) Then
                            Now_Pos = cboMENU01(Pos).ListCount - 1
                        End If
                    Next j
                Case "1"
                    OptMENU01_M(Pos).Value = False
                    OptMENU01_S(Pos).Value = True
    
                    com = BtOpGetFirst
                
                    Do
        
                        sts = BTRV(com, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                        Select Case sts
                            Case BtNoErr
                                If StrConv(YOINREC.REGI_F, vbUnicode) <= "1" Then
                                    cboMENU01(Pos).AddItem StrConv(YOINREC.YOIN_DNAME, vbUnicode) & StrConv(YOINREC.CODE_TYPE, vbUnicode) & StrConv(YOINREC.YOIN_CODE, vbUnicode)
                                End If
                        
                                If StrConv(YOINREC.CODE_TYPE, vbUnicode) = StrConv(tmpMENUREC.CODE_TYPE, vbUnicode) And _
                                    StrConv(YOINREC.YOIN_CODE, vbUnicode) = StrConv(tmpMENUREC.YOIN_CODE, vbUnicode) Then
                                    Now_Pos = cboMENU01(Pos).ListCount - 1
                                End If
                        
                            Case BtErrEOF
                                Exit Do
                            Case Else
                                Call File_Error(sts, com, "メニュー管理マスタ")
                                Exit Function
                        End Select
    
                        com = BtOpGetNext
    
                    Loop
                
                Case Else
                    OptMENU01_M(Pos).Value = False
                    OptMENU01_S(Pos).Value = False
            End Select
            
            txtMenu01(Pos).Text = Trim(StrConv(tmpMENUREC.DISPLAY_ITEM, vbUnicode))
            
        End If
    
    
        If cboMENU01(Pos).ListCount > 0 Then
            cboMENU01(Pos).ListIndex = Now_Pos
        End If
    
        Pos = Pos + 1
    
    Next i
    Display_L_Proc = False

End Function
Private Function Display_R_Proc() As Integer
'----------------------------------------------------------------------------
'                   作業内容（レベル２）表示
'----------------------------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim Now_Pos     As Integer
Dim com         As Integer
Dim sts         As Integer
Dim Pos         As Integer
    
    Display_R_Proc = True
    
    Pos = 0
    For i = (VScroll2.Value - 1) To (VScroll2.Value - 1) + 8
        DoEvents
        
        Call UniCode_Conv(K0_tmpMENU.MENU_LV1, Sel_Code1)
        
        Call UniCode_Conv(K0_tmpMENU.MENU_LV2, Format(i, "000"))
        Call UniCode_Conv(K0_tmpMENU.MENU_LV3, "")
            
        sts = BTRV(BtOpGetEqual, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), K0_tmpMENU, Len(K0_tmpMENU), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                
                Call UniCode_Conv(tmpMENUREC.DEL_FLG, "")
                Call UniCode_Conv(tmpMENUREC.CODE_TYPE, "")
                Call UniCode_Conv(tmpMENUREC.MENU_KBN, "")
                Call UniCode_Conv(tmpMENUREC.DISPLAY_ITEM, "")
                Call UniCode_Conv(tmpMENUREC.CODE_TYPE, "")
                Call UniCode_Conv(tmpMENUREC.YOIN_CODE, "")
                Call UniCode_Conv(tmpMENUREC.PARAM_F, "")
                Call UniCode_Conv(tmpMENUREC.PARAM, "")
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ(一時ファイル)")
                Exit Function
        End Select
        
        '上位のメニューに該当する作業を全てセットする。
        Call UniCode_Conv(K0_YOIN.CODE_TYPE, Sel_CODE_TYPE)
        Call UniCode_Conv(K0_YOIN.YOIN_CODE, "")
        
        cboMENU02(Pos).Clear
        
        com = BtOpGetGreater
        
        Now_Pos = 0
        
        Do
            sts = BTRV(com, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
            Select Case sts
                Case BtNoErr
                                        
                    If StrConv(YOINREC.CODE_TYPE, vbUnicode) <> Sel_CODE_TYPE Then
                        Exit Do
                    End If
                    
                    If StrConv(YOINREC.REGI_F, vbUnicode) <= "1" Then
                        cboMENU02(Pos).AddItem StrConv(YOINREC.YOIN_DNAME, vbUnicode) & "     " & StrConv(YOINREC.CODE_TYPE, vbUnicode) & StrConv(YOINREC.YOIN_CODE, vbUnicode)
                    End If
                
                Case BtErrEOF
                    
                    Exit Do
                
                Case Else
                    Call File_Error(sts, com, "要因マスタ")
                    Exit Function
            End Select

            If StrConv(tmpMENUREC.YOIN_CODE, vbUnicode) = StrConv(YOINREC.YOIN_CODE, vbUnicode) Then
                Now_Pos = cboMENU02(Pos).ListCount - 1
            End If
            
            com = BtOpGetNext
    
        Loop
        
        If Len(Trim(StrConv(tmpMENUREC.MENU_LV2, vbUnicode))) = 0 Then
            chkMENU02_D(Pos).Value = False
            txtMenu02(Pos).Text = ""
'            cboMENU02(Pos).Clear
           If cboMENU02(Pos).ListCount <> 0 Then
               cboMENU02(Pos).ListIndex = 0
           End If
        Else
            If StrConv(tmpMENUREC.DEL_FLG, vbUnicode) = "1" Then
                chkMENU02_D(Pos).Value = True
            Else
                chkMENU02_D(Pos).Value = False
            End If
        
            txtMenu02(Pos).Text = Trim(StrConv(tmpMENUREC.DISPLAY_ITEM, vbUnicode))
        
    
            If cboMENU02(Pos).ListCount > 0 Then
                cboMENU02(Pos).ListIndex = Now_Pos
            End If
    
    
        End If
        
        Pos = Pos + 1
    
    Next i
    
    Display_R_Proc = False

End Function



Private Function Display_MTS_Proc() As Integer
'----------------------------------------------------------------------------
'                   向け先（レベル３）表示
'----------------------------------------------------------------------------
Dim Pos     As Integer
Dim i       As Integer
Dim sts     As Integer
    
    Display_MTS_Proc = True
    
    Pos = 0
    For i = (VScroll3.Value - 1) To (VScroll3.Value - 1) + 8
        DoEvents
        
        Call UniCode_Conv(K0_tmpMENU.MENU_LV1, Sel_Code1)
        
        Call UniCode_Conv(K0_tmpMENU.MENU_LV2, Sel_Code2)
        Call UniCode_Conv(K0_tmpMENU.MENU_LV3, Format(i, "000"))
            
        sts = BTRV(BtOpGetEqual, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), K0_tmpMENU, Len(K0_tmpMENU), 0)
        Select Case sts
            Case BtNoErr
            
                Call UniCode_Conv(K0_MTS.MUKE_CODE, Left(StrConv(tmpMENUREC.PARAM, vbUnicode), 8))
                Call UniCode_Conv(K0_MTS.SS_CODE, Right(StrConv(tmpMENUREC.PARAM, vbUnicode), 8))
                sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                Select Case sts
                    Case BtNoErr
                        txtMTS_CODE(i).Text = StrConv(MTSREC.MUKE_CODE, vbUnicode)
                        txtSS_CODE(i).Text = StrConv(MTSREC.SS_CODE, vbUnicode)
                        txtMTS_NAME(i).Text = StrConv(MTSREC.MUKE_DNAME, vbUnicode)
                    Case BtErrKeyNotFound
                        txtMTS_CODE(i).Text = ""
                        txtSS_CODE(i).Text = ""
                        txtMTS_NAME(i).Text = ""
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                        Exit Function
                End Select
            
            Case BtErrKeyNotFound
                
                Call UniCode_Conv(tmpMENUREC.DEL_FLG, "")
                Call UniCode_Conv(tmpMENUREC.CODE_TYPE, "")
                Call UniCode_Conv(tmpMENUREC.MENU_KBN, "")
                Call UniCode_Conv(tmpMENUREC.DISPLAY_ITEM, "")
                Call UniCode_Conv(tmpMENUREC.CODE_TYPE, "")
                Call UniCode_Conv(tmpMENUREC.YOIN_CODE, "")
                Call UniCode_Conv(tmpMENUREC.PARAM_F, "")
                Call UniCode_Conv(tmpMENUREC.PARAM, "")
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ(一時ファイル)")
                Exit Function
        End Select
        
    
        Pos = Pos + 1
    
    Next i

End Function


Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   メニュー情報更新
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer
Dim ans         As Integer

Dim MENU_LV1    As Integer
Dim MENU_LV2    As Integer
Dim MENU_LV3    As Integer

    Update_Proc = True

    
    
    Call UniCode_Conv(K0_MENU.JGYOBU, Right(cboJIGYOBU.Text, 1))
    
    Call UniCode_Conv(K0_MENU.NAIGAI, Right(cboNAIGAI.Text, 1))
        
    Call UniCode_Conv(K0_MENU.MENU_GRP_NO, Text(ptxMENU_GRP_No).Text)
    Call UniCode_Conv(K0_MENU.MENU_LV1, "")
    Call UniCode_Conv(K0_MENU.MENU_LV2, "")
    Call UniCode_Conv(K0_MENU.MENU_LV3, "")
    
    com = BtOpGetGreaterEqual
    '該当グループ全件削除
    Do
        DoEvents
        Do
            sts = BTRV(com + BtSNoWait, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
            Select Case sts
                Case BtNoErr
                    If StrConv(MENUREC.JGYOBU, vbUnicode) <> Right(cboJIGYOBU.Text, 1) Or _
                         StrConv(MENUREC.NAIGAI, vbUnicode) <> Right(cboNAIGAI.Text, 1) Or _
                        Trim(StrConv(MENUREC.MENU_GRP_NO, vbUnicode)) <> Trim(Text(ptxMENU_GRP_No).Text) Then
                        sts = BtErrEOF
                        Exit Do
                    End If
                    
                    Do
                        sts = BTRV(BtOpDelete, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<MENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpDelete, "メニュー管理マスタ")
                                Exit Function
                        End Select
                    Loop
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<MENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "メニュー管理マスタ")
                    Exit Function
            End Select
        Loop
        
        If sts = BtErrEOF Then
            Exit Do
        End If
        
    Loop
    'データ登録

    com = BtOpGetFirst

    Do
    
        Do
            sts = BTRV(com + BtSNoWait, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), K0_tmpMENU, Len(K0_tmpMENU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<tmpMENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ(一時ファイル)")
                    Exit Function
            End Select
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
    
        If StrConv(tmpMENUREC.DEL_FLG, vbUnicode) = "1" Then
        Else
            Call UniCode_Conv(MENUREC.JGYOBU, Right(cboJIGYOBU.Text, 1))
            Call UniCode_Conv(MENUREC.NAIGAI, Right(cboNAIGAI.Text, 1))
            Call UniCode_Conv(MENUREC.MENU_GRP_NO, Text(ptxMENU_GRP_No).Text)
            Call UniCode_Conv(MENUREC.MENU_GRP, Text(ptxMENU_GRP).Text)
            
            Call UniCode_Conv(MENUREC.MENU_LV1, StrConv(tmpMENUREC.MENU_LV1, vbUnicode))
            Call UniCode_Conv(MENUREC.MENU_LV2, StrConv(tmpMENUREC.MENU_LV2, vbUnicode))
            Call UniCode_Conv(MENUREC.MENU_LV3, StrConv(tmpMENUREC.MENU_LV3, vbUnicode))

            Call UniCode_Conv(MENUREC.MENU_KBN, StrConv(tmpMENUREC.MENU_KBN, vbUnicode))
            Call UniCode_Conv(MENUREC.DISPLAY_ITEM, StrConv(tmpMENUREC.DISPLAY_ITEM, vbUnicode))
            Call UniCode_Conv(MENUREC.CODE_TYPE, StrConv(tmpMENUREC.CODE_TYPE, vbUnicode))
            Call UniCode_Conv(MENUREC.YOIN_CODE, StrConv(tmpMENUREC.YOIN_CODE, vbUnicode))
            Call UniCode_Conv(MENUREC.PARAM, StrConv(tmpMENUREC.PARAM, vbUnicode))
    
            Do
            
                sts = BTRV(BtOpInsert, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                       Beep
                        ans = MsgBox("他端末でデータ使用中です。<MENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpInsert, "メニュー管理マスタ")
                        Exit Function
                End Select
            
            Loop
            
            Do
                sts = BTRV(BtOpDelete, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), K0_tmpMENU, Len(K0_tmpMENU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<tmpMENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpDelete, "メニュー管理マスタ(一時ファイル)")
                        Exit Function
                End Select
        
            Loop
        End If
    
        com = BtOpGetNext
    Loop


    Call Clear_Proc

    Update_Proc = False

End Function

Private Function TBL_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   メニュー情報の一時ファイルへの退避
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim i           As Integer
Dim j           As Integer
Dim k           As Integer

Dim ans         As Integer


    TBL_Set_Proc = True

    com = BtOpGetFirst

    '該当グループ全件削除
    Do
        DoEvents
        Do
            sts = BTRV(com + BtSNoWait, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), K0_tmpMENU, Len(K0_tmpMENU), 0)
            Select Case sts
                Case BtNoErr
                    
                    Do
                        sts = BTRV(BtOpDelete, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), K0_tmpMENU, Len(K0_tmpMENU), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<MENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpDelete, "メニュー管理マスタ")
                                Exit Function
                        End Select
                    Loop
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<MENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "メニュー管理マスタ")
                    Exit Function
            End Select
        Loop
        
        If sts = BtErrEOF Then
            Exit Do
        End If
        
    Loop


    Call UniCode_Conv(K0_MENU.JGYOBU, Right(cboJIGYOBU.Text, 1))
    Call UniCode_Conv(K0_MENU.NAIGAI, Right(cboNAIGAI.Text, 1))
    Call UniCode_Conv(K0_MENU.MENU_GRP_NO, Text(ptxMENU_GRP_No).Text)
    Call UniCode_Conv(K0_MENU.MENU_LV1, "")
    Call UniCode_Conv(K0_MENU.MENU_LV2, "")
    Call UniCode_Conv(K0_MENU.MENU_LV3, "")

    com = BtOpGetGreaterEqual

    Do
        sts = BTRV(com, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(MENUREC.JGYOBU, vbUnicode) <> Right(cboJIGYOBU.Text, 1) Or _
                    StrConv(MENUREC.NAIGAI, vbUnicode) <> Right(cboNAIGAI.Text, 1) Or _
                    Trim(StrConv(MENUREC.MENU_GRP_NO, vbUnicode)) <> Trim(Text(ptxMENU_GRP_No).Text) Then
                
                    Exit Do
                
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ")
                Form_RTN = True
                F1010702.Hide
        End Select
    
        If Len(Trim(Text(ptxMENU_GRP).Text)) = 0 Then
            Text(ptxMENU_GRP).Text = Trim(StrConv(MENUREC.MENU_GRP, vbUnicode))
        End If
    
        Call UniCode_Conv(tmpMENUREC.MENU_LV1, StrConv(MENUREC.MENU_LV1, vbUnicode))
        Call UniCode_Conv(tmpMENUREC.MENU_LV2, StrConv(MENUREC.MENU_LV2, vbUnicode))
        Call UniCode_Conv(tmpMENUREC.MENU_LV3, StrConv(MENUREC.MENU_LV3, vbUnicode))
        Call UniCode_Conv(tmpMENUREC.DEL_FLG, "0")
        Call UniCode_Conv(tmpMENUREC.MENU_KBN, StrConv(MENUREC.MENU_KBN, vbUnicode))
        Call UniCode_Conv(tmpMENUREC.DISPLAY_ITEM, StrConv(MENUREC.DISPLAY_ITEM, vbUnicode))
        Call UniCode_Conv(tmpMENUREC.CODE_TYPE, StrConv(MENUREC.CODE_TYPE, vbUnicode))
        Call UniCode_Conv(tmpMENUREC.YOIN_CODE, StrConv(MENUREC.YOIN_CODE, vbUnicode))
    
        sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
        Select Case sts
            Case BtNoErr
                Call UniCode_Conv(tmpMENUREC.PARAM_F, StrConv(YOINREC.PARAM_F, vbUnicode))
            Case BtErrKeyNotFound
                Call UniCode_Conv(tmpMENUREC.PARAM_F, "0")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "要因マスタ")
                Form_RTN = True
                F1010702.Hide
        End Select
    
        Call UniCode_Conv(tmpMENUREC.PARAM, StrConv(MENUREC.PARAM, vbUnicode))
    
        Do
            sts = BTRV(BtOpInsert, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), K0_tmpMENU, Len(K0_tmpMENU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<MENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpDelete, "メニュー管理マスタ")
                    Exit Function
            End Select
        Loop
    
    
    
        com = BtOpGetNext
    Loop
    
    
    
    If Display_L_Proc() Then
        Exit Function
    End If

    TBL_Set_Proc = False

End Function


Private Sub Clear_Proc()
'----------------------------------------------------------------------------
'                   画面クリア
'----------------------------------------------------------------------------
Dim i   As Integer

    
    Text(ptxMENU_GRP_No).Text = ""
    Text(ptxMENU_GRP).Text = ""
    cboJIGYOBU.ListIndex = 0
    cboNAIGAI.ListIndex = 0
    
    For i = 0 To 8
        
        FrmMENU01(i).BackColor = Base_Color
        FrmMENU02(i).BackColor = Base_Color
        
        chkMENU01_D(i).Value = False
        txtMenu01(i).Text = ""
        cboMENU01(i).Clear
        chkMENU02_D(i).Value = False
        txtMenu02(i).Text = ""
        cboMENU02(i).Clear
    
        FrmMENU02(i).Visible = False
    
    
        txtMTS_CODE(i).Text = ""
        txtSS_CODE(i).Text = ""
        txtMTS_NAME(i).Text = ""
    
    
    Next i

    LblLEVEL02(1).Visible = False
    LblLEVEL02(2).Visible = False
    
    VScroll2.Visible = False
    Frame2.Visible = False
    
End Sub


Private Sub txtMenu01_GotFocus(Index As Integer)
    
    If cboMENU01(Index).ListCount <> 0 Then
        txtMenu01(Index).Text = Left(cboMENU01(Index).Text, Len(cboMENU01(Index).Text) - 2)
    End If
 
End Sub

Private Sub txtMTS_CODE_GotFocus(Index As Integer)
    
    If txtMTS_CODE(Index).TabStop = True Then
        txtMTS_CODE(Index).Text = Trim(txtMTS_CODE(Index).Text)
        txtMTS_CODE(Index).SelStart = 0
        txtMTS_CODE(Index).SelLength = Len(txtMTS_CODE(Index).Text)
    End If


End Sub



Private Sub txtMTS_CODE_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    txtSS_CODE(Index).SetFocus

End Sub

Private Sub txtSS_CODE_GotFocus(Index As Integer)
    
    If txtSS_CODE(Index).TabStop = True Then
        txtSS_CODE(Index).Text = Trim(txtSS_CODE(Index).Text)
        txtSS_CODE(Index).SelStart = 0
        txtSS_CODE(Index).SelLength = Len(txtSS_CODE(Index).Text)
    End If


End Sub

Private Sub txtSS_CODE_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
Dim sts As Integer
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    Call UniCode_Conv(K0_MTS.MUKE_CODE, txtMTS_CODE(Index).Text)
    Call UniCode_Conv(K0_MTS.SS_CODE, txtSS_CODE(Index).Text)
    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    Select Case sts
        Case BtNoErr
            txtMTS_NAME(Index).Text = StrConv(MTSREC.MUKE_DNAME, vbUnicode)
        Case BtErrKeyNotFound
            txtMTS_NAME(Index).Text = ""
            Beep
            MsgBox "入力した項目はエラーです。（必須入力）"
            txtMTS_CODE(Index).SetFocus
            Exit Sub
        Case Else
            Call File_Error(sts, BtOpGetEqual, "要因マスタ")
            Form_RTN = True
            F1010702.Hide
    End Select

End Sub

Private Sub VScroll1_Change()
    
    If Display_L_Proc() Then
        Form_RTN = True
        F1010702.Hide
    End If

End Sub

Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   メニュー情報削除
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer
Dim ans         As Integer

    Delete_Proc = True
    
    Call UniCode_Conv(K0_MENU.JGYOBU, Right(cboJIGYOBU.Text, 1))
    
    Call UniCode_Conv(K0_MENU.NAIGAI, Right(cboNAIGAI.Text, 1))
        
    Call UniCode_Conv(K0_MENU.MENU_GRP_NO, Text(ptxMENU_GRP_No).Text)
    Call UniCode_Conv(K0_MENU.MENU_LV1, "")
    Call UniCode_Conv(K0_MENU.MENU_LV2, "")
    Call UniCode_Conv(K0_MENU.MENU_LV3, "")
    
    com = BtOpGetGreaterEqual
    '該当グループ全件削除
    Do
        DoEvents
        Do
            sts = BTRV(com + BtSNoWait, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
            Select Case sts
                Case BtNoErr
                    If StrConv(MENUREC.JGYOBU, vbUnicode) <> Right(cboJIGYOBU.Text, 1) Or _
                         StrConv(MENUREC.NAIGAI, vbUnicode) <> Right(cboNAIGAI.Text, 1) Or _
                        Trim(StrConv(MENUREC.MENU_GRP_NO, vbUnicode)) <> Trim(Text(ptxMENU_GRP_No).Text) Then
                        sts = BtErrEOF
                        Exit Do
                    End If
                    
                    Do
                        sts = BTRV(BtOpDelete, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<MENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpDelete, "メニュー管理マスタ")
                                Exit Function
                        End Select
                    Loop
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<MENU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "メニュー管理マスタ")
                    Exit Function
            End Select
        Loop
        
        If sts = BtErrEOF Then
            Exit Do
        End If
        
    Loop

    Call Clear_Proc

    Delete_Proc = False


End Function

Private Sub VScroll2_Change()
    
    If Display_R_Proc() Then
        Form_RTN = True
        F1010702.Hide
    End If


End Sub

Private Sub VScroll3_Change()
Exit Sub
    If Display_MTS_Proc() Then
        Form_RTN = True
        F1010702.Hide
    End If


End Sub
