VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form SEI00161 
   Caption         =   "[請求システム]見積書作成処理"
   ClientHeight    =   12375
   ClientLeft      =   2025
   ClientTop       =   -3210
   ClientWidth     =   18360
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
   ScaleHeight     =   12375
   ScaleWidth      =   18360
   StartUpPosition =   2  '画面の中央
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   62
      Left            =   12840
      TabIndex        =   69
      Top             =   11160
      Width           =   1335
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   57
      Left            =   12840
      TabIndex        =   63
      Top             =   10800
      Width           =   1335
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   52
      Left            =   12840
      TabIndex        =   57
      Top             =   10440
      Width           =   1335
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   47
      Left            =   12840
      TabIndex        =   51
      Top             =   10080
      Width           =   1335
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   42
      Left            =   12840
      TabIndex        =   45
      Top             =   9720
      Width           =   1335
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   37
      Left            =   12840
      TabIndex        =   39
      Top             =   9360
      Width           =   1335
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   32
      Left            =   12840
      TabIndex        =   33
      Top             =   9000
      Width           =   1335
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   27
      Left            =   12840
      TabIndex        =   27
      Top             =   8640
      Width           =   1335
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   22
      Left            =   12840
      TabIndex        =   21
      Top             =   8280
      Width           =   1335
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   17
      Left            =   12840
      TabIndex        =   15
      Top             =   7920
      Width           =   1335
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   12
      Left            =   12840
      TabIndex        =   10
      Text            =   "99999999.99"
      Top             =   7560
      Width           =   1335
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
      Height          =   375
      Index           =   60
      Left            =   10680
      TabIndex        =   67
      Top             =   11160
      Width           =   1335
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
      Height          =   375
      Index           =   55
      Left            =   10680
      TabIndex        =   61
      Top             =   10800
      Width           =   1335
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
      Height          =   375
      Index           =   50
      Left            =   10680
      TabIndex        =   55
      Top             =   10440
      Width           =   1335
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
      Height          =   375
      Index           =   45
      Left            =   10680
      TabIndex        =   49
      Top             =   10080
      Width           =   1335
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
      Height          =   375
      Index           =   40
      Left            =   10680
      TabIndex        =   43
      Top             =   9720
      Width           =   1335
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
      Height          =   375
      Index           =   35
      Left            =   10680
      TabIndex        =   37
      Top             =   9360
      Width           =   1335
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
      Height          =   375
      Index           =   30
      Left            =   10680
      TabIndex        =   31
      Top             =   9000
      Width           =   1335
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
      Height          =   375
      Index           =   25
      Left            =   10680
      TabIndex        =   25
      Top             =   8640
      Width           =   1335
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
      Height          =   375
      Index           =   20
      Left            =   10680
      TabIndex        =   19
      Top             =   8280
      Width           =   1335
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
      Height          =   375
      Index           =   15
      Left            =   10680
      TabIndex        =   13
      Top             =   7920
      Width           =   1335
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
      Height          =   375
      Index           =   10
      Left            =   10680
      TabIndex        =   8
      Text            =   "999.99"
      Top             =   7560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   63
      Left            =   14160
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   11160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   58
      Left            =   14160
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   10800
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   53
      Left            =   14160
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   10440
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   48
      Left            =   14160
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   10080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   43
      Left            =   14160
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   9720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   38
      Left            =   14160
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   9360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   33
      Left            =   14160
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   28
      Left            =   14160
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   8640
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   23
      Left            =   14160
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8280
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   18
      Left            =   14160
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   13
      Left            =   14160
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Height          =   240
      Index           =   10
      Left            =   8520
      TabIndex        =   66
      Top             =   11280
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   240
      Index           =   9
      Left            =   8520
      TabIndex        =   60
      Top             =   10920
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   240
      Index           =   8
      Left            =   8520
      TabIndex        =   54
      Top             =   10560
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   240
      Index           =   7
      Left            =   8520
      TabIndex        =   48
      Top             =   10200
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   240
      Index           =   6
      Left            =   8520
      TabIndex        =   42
      Top             =   9840
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   240
      Index           =   5
      Left            =   8520
      TabIndex        =   36
      Top             =   9480
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   240
      Index           =   4
      Left            =   8520
      TabIndex        =   30
      Top             =   9120
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   240
      Index           =   3
      Left            =   8520
      TabIndex        =   24
      Top             =   8760
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   240
      Index           =   2
      Left            =   8520
      TabIndex        =   18
      Top             =   8400
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   240
      Index           =   1
      Left            =   8520
      TabIndex        =   12
      Top             =   8040
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Height          =   240
      Index           =   0
      Left            =   8520
      TabIndex        =   7
      Top             =   7680
      Width           =   375
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
      IMEMode         =   3  'ｵﾌ固定
      Index           =   66
      Left            =   17520
      MaxLength       =   1
      TabIndex        =   80
      Top             =   960
      Width           =   255
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
      IMEMode         =   1  'ｵﾝ
      Index           =   65
      Left            =   14040
      MaxLength       =   20
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin TrueDBGrid80.TDBDropDown TDBDropDown2 
      Height          =   2055
      Left            =   4560
      TabIndex        =   120
      Top             =   3840
      Width           =   1935
      _ExtentX        =   3413
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   64
      Left            =   15480
      TabIndex        =   71
      Top             =   11160
      Visible         =   0   'False
      Width           =   1455
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
      Height          =   375
      Index           =   61
      Left            =   12000
      TabIndex        =   68
      Top             =   11160
      Width           =   855
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   59
      Left            =   15480
      TabIndex        =   65
      Top             =   10800
      Visible         =   0   'False
      Width           =   1455
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
      Height          =   375
      Index           =   56
      Left            =   12000
      TabIndex        =   62
      Top             =   10800
      Width           =   855
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   54
      Left            =   15480
      TabIndex        =   59
      Top             =   10440
      Visible         =   0   'False
      Width           =   1455
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
      Height          =   375
      Index           =   51
      Left            =   12000
      TabIndex        =   56
      Top             =   10440
      Width           =   855
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   49
      Left            =   15480
      TabIndex        =   53
      Top             =   10080
      Visible         =   0   'False
      Width           =   1455
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
      Height          =   375
      Index           =   46
      Left            =   12000
      TabIndex        =   50
      Top             =   10080
      Width           =   855
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   44
      Left            =   15480
      TabIndex        =   47
      Top             =   9720
      Visible         =   0   'False
      Width           =   1455
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
      Height          =   375
      Index           =   41
      Left            =   12000
      TabIndex        =   44
      Top             =   9720
      Width           =   855
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   39
      Left            =   15480
      TabIndex        =   41
      Top             =   9360
      Visible         =   0   'False
      Width           =   1455
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
      Height          =   375
      Index           =   36
      Left            =   12000
      TabIndex        =   38
      Top             =   9360
      Width           =   855
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   34
      Left            =   15480
      TabIndex        =   35
      Top             =   9000
      Visible         =   0   'False
      Width           =   1455
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
      Height          =   375
      Index           =   31
      Left            =   12000
      TabIndex        =   32
      Top             =   9000
      Width           =   855
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   29
      Left            =   15480
      TabIndex        =   29
      Top             =   8640
      Visible         =   0   'False
      Width           =   1455
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
      Height          =   375
      Index           =   26
      Left            =   12000
      TabIndex        =   26
      Top             =   8640
      Width           =   855
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   24
      Left            =   15480
      TabIndex        =   23
      Top             =   8280
      Visible         =   0   'False
      Width           =   1455
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
      Height          =   375
      Index           =   21
      Left            =   12000
      TabIndex        =   20
      Top             =   8280
      Width           =   855
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   19
      Left            =   15480
      TabIndex        =   17
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
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
      Height          =   375
      Index           =   16
      Left            =   12000
      TabIndex        =   14
      Top             =   7920
      Width           =   855
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
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   14
      Left            =   15480
      TabIndex        =   11
      Top             =   7560
      Visible         =   0   'False
      Width           =   1455
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
      Height          =   375
      Index           =   11
      Left            =   12000
      TabIndex        =   9
      Top             =   7560
      Width           =   855
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
      Index           =   1
      Left            =   2520
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   4
      Top             =   1440
      Width           =   5055
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
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   3
      Top             =   1440
      Width           =   975
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
      Height          =   375
      Index           =   69
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   10200
      Visible         =   0   'False
      Width           =   855
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
      Height          =   375
      Index           =   68
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   855
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
      Height          =   375
      Index           =   67
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   9240
      Visible         =   0   'False
      Width           =   855
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1575
      Index           =   0
      Left            =   600
      TabIndex        =   72
      Top             =   8070
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2778
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"SEI00161.frx":0000
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
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   1
      Text            =   "AD-HHSJ10P"
      Top             =   960
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
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   1080
      Width           =   4335
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
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
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
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
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
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
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
      Left            =   12360
      Locked          =   -1  'True
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
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
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   0
      Top             =   480
      Width           =   855
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
      Left            =   13320
      TabIndex        =   93
      ToolTipText     =   "商品化単価を品目マスターに登録します"
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
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
      Left            =   11760
      TabIndex        =   92
      ToolTipText     =   "商品化単価見積書(EXCEL)を作成します"
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   600
      Width           =   2415
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
      Left            =   1440
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   2295
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
      Left            =   5760
      TabIndex        =   91
      ToolTipText     =   "商品化単価を計算します(F9)"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "見積書発行/保存"
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
      Left            =   3480
      TabIndex        =   90
      ToolTipText     =   "商品化構成を保存します"
      Top             =   0
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   11040
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   88
      Top             =   0
      Visible         =   0   'False
      Width           =   255
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
      Left            =   1800
      TabIndex        =   89
      ToolTipText     =   "商品化構成を読み込みます（Ｆ5）"
      Top             =   0
      Width           =   1455
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
      Left            =   480
      TabIndex        =   87
      Top             =   0
      Width           =   1215
   End
   Begin TrueDBGrid80.TDBDropDown TDBDropDown1 
      Height          =   2055
      Left            =   1440
      TabIndex        =   101
      Top             =   3720
      Width           =   1935
      _ExtentX        =   3413
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
      Height          =   4575
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   17895
      _ExtentX        =   31565
      _ExtentY        =   8070
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "№"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   1
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "種別"
      Columns(1).DataField=   ""
      Columns(1).DropDown=   "TDBDropDown1"
      Columns(1).DropDown.vt=   8
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   1
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "事業部"
      Columns(2).DataField=   ""
      Columns(2).DropDown=   "TDBDropDown2"
      Columns(2).DropDown.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "構成品番"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "品　名"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "提出品名"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "員数"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "単位"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "提出単価"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "提出金額"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "仕入＠"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "販売＠"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "合計金額"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "印刷順"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   14
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=14"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=900"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=794"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8705"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2196"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2090"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=256"
      Splits(0)._ColumnProps(10)=   "Column(1).Button=1"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=1958"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1852"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=512"
      Splits(0)._ColumnProps(16)=   "Column(2).Button=1"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=2831"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2725"
      Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=512"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=3757"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=3651"
      Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=8704"
      Splits(0)._ColumnProps(27)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(29)=   "Column(5).Width=5159"
      Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=5054"
      Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=512"
      Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(34)=   "Column(6).Width=1931"
      Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=1826"
      Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=514"
      Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(39)=   "Column(7).Width=1005"
      Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=900"
      Splits(0)._ColumnProps(42)=   "Column(7)._ColStyle=512"
      Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(44)=   "Column(8).Width=2143"
      Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=2037"
      Splits(0)._ColumnProps(47)=   "Column(8)._ColStyle=514"
      Splits(0)._ColumnProps(48)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(49)=   "Column(9).Width=2117"
      Splits(0)._ColumnProps(50)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(9)._WidthInPix=2011"
      Splits(0)._ColumnProps(52)=   "Column(9)._ColStyle=8706"
      Splits(0)._ColumnProps(53)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(54)=   "Column(10).Width=1879"
      Splits(0)._ColumnProps(55)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(10)._WidthInPix=1773"
      Splits(0)._ColumnProps(57)=   "Column(10)._ColStyle=514"
      Splits(0)._ColumnProps(58)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(59)=   "Column(11).Width=2143"
      Splits(0)._ColumnProps(60)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(11)._WidthInPix=2037"
      Splits(0)._ColumnProps(62)=   "Column(11)._ColStyle=514"
      Splits(0)._ColumnProps(63)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(64)=   "Column(12).Width=2117"
      Splits(0)._ColumnProps(65)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(66)=   "Column(12)._WidthInPix=2011"
      Splits(0)._ColumnProps(67)=   "Column(12)._ColStyle=8706"
      Splits(0)._ColumnProps(68)=   "Column(12).AllowFocus=0"
      Splits(0)._ColumnProps(69)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(70)=   "Column(13).Width=1482"
      Splits(0)._ColumnProps(71)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(72)=   "Column(13)._WidthInPix=1376"
      Splits(0)._ColumnProps(73)=   "Column(13).Order=14"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowAddNew     =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   2
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14,.alignment=2"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=0"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14,.alignment=0"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=66,.parent=13,.alignment=0"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=14,.alignment=2"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=0"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14,.alignment=2"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=0,.bgcolor=&H80000016&"
      _StyleDefs(53)  =   ":id=46,.locked=-1"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14,.alignment=2"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=78,.parent=13,.alignment=0"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=75,.parent=14,.alignment=2"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=76,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=77,.parent=17"
      _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=1"
      _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14,.alignment=2"
      _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
      _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
      _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=82,.parent=13,.alignment=0"
      _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=79,.parent=14,.alignment=2"
      _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=80,.parent=15"
      _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=81,.parent=17"
      _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=86,.parent=13,.alignment=1"
      _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=83,.parent=14,.alignment=2"
      _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=84,.parent=15"
      _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=85,.parent=17"
      _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=90,.parent=13,.alignment=1,.bgcolor=&H80000016&"
      _StyleDefs(74)  =   ":id=90,.locked=-1"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=87,.parent=14,.alignment=2"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=88,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=89,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=54,.parent=13,.alignment=1,.bgcolor=&H80000005&"
      _StyleDefs(79)  =   ":id=54,.locked=0"
      _StyleDefs(80)  =   "Splits(0).Columns(10).HeadingStyle:id=51,.parent=14,.alignment=2"
      _StyleDefs(81)  =   "Splits(0).Columns(10).FooterStyle:id=52,.parent=15"
      _StyleDefs(82)  =   "Splits(0).Columns(10).EditorStyle:id=53,.parent=17"
      _StyleDefs(83)  =   "Splits(0).Columns(11).Style:id=58,.parent=13,.alignment=1,.bgcolor=&H80000005&"
      _StyleDefs(84)  =   ":id=58,.locked=0"
      _StyleDefs(85)  =   "Splits(0).Columns(11).HeadingStyle:id=55,.parent=14,.alignment=2"
      _StyleDefs(86)  =   "Splits(0).Columns(11).FooterStyle:id=56,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(11).EditorStyle:id=57,.parent=17"
      _StyleDefs(88)  =   "Splits(0).Columns(12).Style:id=62,.parent=13,.alignment=1,.bgcolor=&H8000000F&"
      _StyleDefs(89)  =   ":id=62,.locked=-1"
      _StyleDefs(90)  =   "Splits(0).Columns(12).HeadingStyle:id=59,.parent=14,.alignment=2"
      _StyleDefs(91)  =   "Splits(0).Columns(12).FooterStyle:id=60,.parent=15"
      _StyleDefs(92)  =   "Splits(0).Columns(12).EditorStyle:id=61,.parent=17"
      _StyleDefs(93)  =   "Splits(0).Columns(13).Style:id=98,.parent=13"
      _StyleDefs(94)  =   "Splits(0).Columns(13).HeadingStyle:id=95,.parent=14"
      _StyleDefs(95)  =   "Splits(0).Columns(13).FooterStyle:id=96,.parent=15"
      _StyleDefs(96)  =   "Splits(0).Columns(13).EditorStyle:id=97,.parent=17"
      _StyleDefs(97)  =   "Named:id=33:Normal"
      _StyleDefs(98)  =   ":id=33,.parent=0"
      _StyleDefs(99)  =   "Named:id=34:Heading"
      _StyleDefs(100) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(101) =   ":id=34,.wraptext=-1"
      _StyleDefs(102) =   "Named:id=35:Footing"
      _StyleDefs(103) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(104) =   "Named:id=36:Selected"
      _StyleDefs(105) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(106) =   "Named:id=37:Caption"
      _StyleDefs(107) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(108) =   "Named:id=38:HighlightRow"
      _StyleDefs(109) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(110) =   "Named:id=39:EvenRow"
      _StyleDefs(111) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(112) =   "Named:id=40:OddRow"
      _StyleDefs(113) =   ":id=40,.parent=33"
      _StyleDefs(114) =   "Named:id=41:RecordSelector"
      _StyleDefs(115) =   ":id=41,.parent=34"
      _StyleDefs(116) =   "Named:id=42:FilterBar"
      _StyleDefs(117) =   ":id=42,.parent=33"
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   975
      Index           =   1
      Left            =   600
      TabIndex        =   73
      Top             =   10110
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1720
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"SEI00161.frx":00BE
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
      Height          =   375
      Index           =   9
      Left            =   6090
      MaxLength       =   8
      TabIndex        =   81
      Top             =   8520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "合計金額"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   17
      Left            =   8040
      TabIndex        =   105
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "提出単価"
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
      Height          =   375
      Index           =   14
      Left            =   12840
      TabIndex        =   132
      Top             =   7200
      Width           =   1335
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
      Height          =   375
      Index           =   13
      Left            =   10680
      TabIndex        =   131
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label lblZEN_GOUKEI_KIN 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11880
      TabIndex        =   130
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblZEN_T_GOUKEI_KIN 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10440
      TabIndex        =   129
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "前回合計金額"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   8040
      TabIndex        =   128
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblGOUKEI_T_KIN 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10440
      TabIndex        =   127
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "提出金額"
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
      Height          =   375
      Index           =   10
      Left            =   14160
      TabIndex        =   126
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Line Line13 
      X1              =   8160
      X2              =   9240
      Y1              =   11160
      Y2              =   11160
   End
   Begin VB.Line Line12 
      X1              =   8160
      X2              =   9240
      Y1              =   10800
      Y2              =   10800
   End
   Begin VB.Line Line11 
      X1              =   8160
      X2              =   9240
      Y1              =   10440
      Y2              =   10440
   End
   Begin VB.Line Line10 
      X1              =   9240
      X2              =   8160
      Y1              =   10080
      Y2              =   10080
   End
   Begin VB.Line Line9 
      X1              =   9240
      X2              =   8160
      Y1              =   9720
      Y2              =   9720
   End
   Begin VB.Line Line8 
      X1              =   8160
      X2              =   9240
      Y1              =   9360
      Y2              =   9360
   End
   Begin VB.Line Line7 
      X1              =   9240
      X2              =   8160
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Line Line6 
      X1              =   9240
      X2              =   8160
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Line Line5 
      X1              =   9240
      X2              =   8160
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line4 
      X1              =   9240
      X2              =   8160
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line3 
      X1              =   8160
      X2              =   9240
      Y1              =   11520
      Y2              =   11520
   End
   Begin VB.Line Line2 
      X1              =   8160
      X2              =   8160
      Y1              =   7560
      Y2              =   11520
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "見積書表示ﾌﾗｸﾞ"
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
      Height          =   375
      Index           =   9
      Left            =   8160
      TabIndex        =   125
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "3：打切り"
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
      Left            =   16320
      TabIndex        =   124
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "部材担当者"
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
      Left            =   12840
      TabIndex        =   123
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "品名"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   4080
      TabIndex        =   122
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
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
      Index           =   2
      Left            =   9240
      TabIndex        =   121
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblGOUKEI_KIN 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11880
      TabIndex        =   83
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "金　額"
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
      Height          =   375
      Index           =   15
      Left            =   15480
      TabIndex        =   119
      Top             =   7200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "管理費"
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
      Height          =   375
      Index           =   10
      Left            =   9240
      TabIndex        =   118
      Top             =   11160
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "梱包ASSY"
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
      Height          =   375
      Index           =   9
      Left            =   9240
      TabIndex        =   117
      Top             =   10800
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "副資材"
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
      Height          =   375
      Index           =   8
      Left            =   9240
      TabIndex        =   116
      Top             =   10440
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "梱包材"
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
      Height          =   375
      Index           =   7
      Left            =   9240
      TabIndex        =   115
      Top             =   10080
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "設置工事説明書"
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
      Height          =   375
      Index           =   6
      Left            =   9240
      TabIndex        =   114
      Top             =   9720
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "品番表示ﾗﾍﾞﾙ"
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
      Height          =   375
      Index           =   5
      Left            =   9240
      TabIndex        =   113
      Top             =   9360
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "PE資材"
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
      Height          =   375
      Index           =   4
      Left            =   9240
      TabIndex        =   112
      Top             =   9000
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "PE加工"
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
      Height          =   375
      Index           =   3
      Left            =   9240
      TabIndex        =   111
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "PF加工"
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
      Height          =   375
      Index           =   2
      Left            =   9240
      TabIndex        =   110
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "単位"
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
      Height          =   375
      Index           =   19
      Left            =   12000
      TabIndex        =   109
      Top             =   7200
      Width           =   855
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
      Height          =   375
      Index           =   16
      Left            =   9240
      TabIndex        =   108
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "商品化工料"
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
      Height          =   375
      Index           =   1
      Left            =   9240
      TabIndex        =   107
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BorderStyle     =   1  '実線
      Caption         =   "中西工料"
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
      Height          =   375
      Index           =   0
      Left            =   9240
      TabIndex        =   106
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "品名ｶﾃｺﾞﾘｰ"
      BeginProperty Font 
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
      Left            =   240
      TabIndex        =   104
      Top             =   1560
      Width           =   1215
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
      Height          =   195
      Index           =   100
      Left            =   630
      TabIndex        =   102
      Top             =   9930
      Width           =   1095
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
      Left            =   600
      TabIndex        =   100
      Top             =   7890
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   17520
      Y1              =   2400
      Y2              =   2400
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
      Left            =   720
      TabIndex        =   99
      Top             =   600
      Width           =   735
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
      Left            =   12120
      TabIndex        =   98
      Top             =   1080
      Width           =   255
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
      Left            =   11520
      TabIndex        =   97
      Top             =   1080
      Width           =   255
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
      Left            =   10920
      TabIndex        =   96
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "親品番"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   95
      Top             =   1080
      Width           =   855
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
      Height          =   315
      Index           =   0
      Left            =   720
      TabIndex        =   94
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
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
      Left            =   5130
      TabIndex        =   103
      Top             =   8520
      Visible         =   0   'False
      Width           =   975
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
Attribute VB_Name = "SEI00161"
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

Private Const ptxCATEGORY_CODE% = 8         '品名ｶﾃｺﾞﾘｰｺｰﾄﾞ

Private Const ptxTANKA_KIRIKAE_DT% = 9      '単価切替日



'--------------------------------
Private Const ptxNAKANISHI_QTY% = 10        '中西工料　数量
Private Const ptxNAKANISHI_TANI% = 11       '中西工料　単位
Private Const ptxNAKANISHI_T_TAN% = 12      '中西工料　提出単価
Private Const ptxNAKANISHI_T_KIN% = 13      '中西工料　提出金額
Private Const ptxNAKANISHI_KIN% = 14        '中西工料　金額


Private Const ptxSHOHIN_QTY% = 15           '商品化工料　数量
Private Const ptxSHOHIN_TANI% = 16          '商品化工料　単位
Private Const ptxSHOHIN_T_TAN% = 17         '商品化工料　提出単価
Private Const ptxSHOHIN_T_KIN% = 18         '商品化工料　提出金額
Private Const ptxSHOHIN_KIN% = 19           '商品化工料　金額

Private Const ptxPF_KAKOU_QTY% = 20         'PF加工　数量
Private Const ptxPF_KAKOU_TANI% = 21        'PF加工　単位
Private Const ptxPF_KAKOU_T_TAN% = 22       'PF加工　提出単価
Private Const ptxPF_KAKOU_T_KIN% = 23       'PF加工　提出金額
Private Const ptxPF_KAKOU_KIN% = 24         'PF加工　金額

Private Const ptxPE_KAKOU_QTY% = 25         'PE加工　数量
Private Const ptxPE_KAKOU_TANI% = 26        'PE加工　単位
Private Const ptxPE_KAKOU_T_TAN% = 27       'PE加工　提出単価
Private Const ptxPE_KAKOU_T_KIN% = 28       'PE加工　提出金額
Private Const ptxPE_KAKOU_KIN% = 29         'PE加工　金額

Private Const ptxPE_SHIZAI_QTY% = 30        'PE資材　数量
Private Const ptxPE_SHIZAI_TANI% = 31       'PE資材　単位
Private Const ptxPE_SHIZAI_T_TAN% = 32      'PE資材　提出単価
Private Const ptxPE_SHIZAI_T_KIN% = 33      'PE資材　提出金額
Private Const ptxPE_SHIZAI_KIN% = 34        'PE資材　金額

Private Const ptxHINBAN_LABEL_QTY% = 35     '品番表示ﾗﾍﾞﾙ　数量
Private Const ptxHINBAN_LABEL_TANI% = 36    '品番表示ﾗﾍﾞﾙ　単位
Private Const ptxHINBAN_LABEL_T_TAN% = 37   '品番表示ﾗﾍﾞﾙ　提出単価
Private Const ptxHINBAN_LABEL_T_KIN% = 38   '品番表示ﾗﾍﾞﾙ　提出金額
Private Const ptxHINBAN_LABEL_KIN% = 39     '品番表示ﾗﾍﾞﾙ　金額

Private Const ptxKOUJI_SETSU_QTY% = 40      '設置工事説明書　数量
Private Const ptxKOUJI_SETSU_TANI% = 41     '設置工事説明書　単位
Private Const ptxKOUJI_SETSU_T_TAN% = 42    '設置工事説明書　提出単価
Private Const ptxKOUJI_SETSU_T_KIN% = 43    '設置工事説明書　提出金額
Private Const ptxKOUJI_SETSU_KIN% = 44      '設置工事説明書　金額


Private Const ptxKONPOU_QTY% = 45           '梱包材　数量
Private Const ptxKONPOU_TANI% = 46          '梱包材　単位
Private Const ptxKONPOU_T_TAN% = 47         '梱包材　提出単価
Private Const ptxKONPOU_T_KIN% = 48         '梱包材　提出金額
Private Const ptxKONPOU_KIN% = 49           '梱包材　金額

Private Const ptxFUKU_SHIZAI_QTY% = 50      '副資材　数量
Private Const ptxFUKU_SHIZAI_TANI% = 51     '副資材　単位
Private Const ptxFUKU_SHIZAI_T_TAN% = 52    '副資材　提出単価
Private Const ptxFUKU_SHIZAI_T_KIN% = 53    '副資材　提出金額
Private Const ptxFUKU_SHIZAI_KIN% = 54      '副資材　金額

Private Const ptxKONPOU_ASSY_QTY% = 55      '梱包ASSY　数量
Private Const ptxKONPOU_ASSY_TANI% = 56     '梱包ASSY　単位
Private Const ptxKONPOU_ASSY_T_TAN% = 57    '梱包ASSY　提出単価
Private Const ptxKONPOU_ASSY_T_KIN% = 58    '梱包ASSY　提出金額
Private Const ptxKONPOU_ASSY_KIN% = 59      '梱包ASSY　金額

Private Const ptxKANRI_QTY% = 60            '管理費　数量
Private Const ptxKANRI_TANI% = 61           '管理費　単位
Private Const ptxKANRI_T_TAN% = 62          '管理費　提出単価
Private Const ptxKANRI_T_KIN% = 63          '管理費　提出金額
Private Const ptxKANRI_KIN% = 64            '管理費　金額






'---------------------------------------    2017.07.08
Private Const ptxBuzai_Tanto_Name% = 65     '部材担当者
Private Const ptxNAI_BUHIN% = 66            '国内供給区分（3：打切り）


'Private Const ptxNAKANISHI_T_KIN% = 37      '中西工料　提出金額
'Private Const ptxSHOHIN_T_KIN% = 38         '商品化工料　提出金額
'Private Const ptxPF_KAKOU_T_KIN% = 39       'PF加工　提出金額
'Private Const ptxPE_KAKOU_T_KIN% = 40       'PE加工　提出金額
'Private Const ptxPE_SHIZAI_T_KIN% = 41      'PF資材　提出金額
'Private Const ptxHINBAN_LABEL_T_KIN% = 42   '品番表示ﾗﾍﾞﾙ　提出金額
'Private Const ptxKOUJI_SETSU_T_KIN% = 43    '設置工事説明書　提出金額
'Private Const ptxKONPOU_T_KIN% = 44         '梱包材　提出金額
'Private Const ptxFUKU_SHIZAI_T_KIN% = 45    '副資材　提出金額
'Private Const ptxKONPOU_ASSY_T_KIN% = 46    '梱包ASSY　提出金額
'Private Const ptxKANRI_T_KIN% = 47         '管理費　提出金額



'---------------------------------------    2017.07.08





Private Const ptxS_CLASS_CODE% = 67        '商品化ｸﾗｽ
Private Const ptxF_CLASS_CODE% = 68        '付加ｸﾗｽ
Private Const ptxN_CLASS_CODE% = 69        '内職ｸﾗｽ





'------------------------------------   'コンボ定義
Private Const pcmbSHIMUKE% = 0          '仕向け先
Private Const pcmbCATEGORY_Name% = 1    '品名ｶﾃｺﾞﾘｰ


'------------------------------------   'リッチテキストボックス定義
Private Const prchBIKOU% = 0            '備考
Private Const prchM_BIKOU% = 1          '見積書備考



'------------------------------------   'ﾁｪｯｸボックス定義   2017.07.08
Private Const chkNAKANISHI_F% = 0           '中西工料
Private Const chkSHOHIN_F% = 1              '商品化工料
Private Const chkPF_KAKOU_F% = 2            'PF加工
Private Const chkPE_KAKOU_F% = 3            'PE加工
Private Const chkPE_SHIZAI_F% = 4           'PF資材
Private Const chkHINBAN_LABEL_F% = 5        '品番表示ﾗﾍﾞﾙ
Private Const chkKOUJI_SETSU_F% = 6         '設置工事説明書
Private Const chkKONPOU_F% = 7              '梱包材
Private Const chkFUKU_SHIZAI_F% = 8         '副資材
Private Const chkKONPOU_ASSY_F% = 9         '梱包ASSY
Private Const chkKANRI_F% = 10              '管理費
    


'------------------------------------   '構成品
Private Const pGrdKOUSEI% = 0


Private Const Min_Row% = 1              '最小行数

Dim Max_Row   As Integer                'グリッド最大表示件数

Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 13             '最大列数　9-->13 2017.07.08



Private Const ColNO% = 0                '№
Private Const ColKO_SYUBETSU% = 1       '種別
Private Const ColKO_JGYOBU% = 2         '事業部
Private Const ColKO_S_HIN_GAI% = 3      '指図票品番
Private Const ColKO_HIN_NAME% = 4       '品名
Private Const ColKO_T_HIN_NAME% = 5     '提出品名   2017.07.08
Private Const ColKO_QTY% = 6            '員数
Private Const ColKO_TANI% = 7           '単位   2017.07.08
Private Const ColKO_T_TANKA% = 8        '提出単価   2017.07.08
Private Const ColKO_T_KINGAKU% = 9      '提出金額   2017.07.08
Private Const ColG_ST_SHITAN% = 10      '仕入＠
Private Const ColG_ST_URITAN% = 11      '売上＠

Private Const ColG_KINGAKU% = 12        '合計金額

Private Const ColPRINT_SEQ% = 13        '印刷順
                                        
                                        

'-----------------------------------    ドロップダウン
Dim SYUBETSU        As New XArrayDB
Dim JGYOBU          As New XArrayDB



'-----------------------------------    見積書表示フラグ  2017.09.28
Dim Print_FLG(0 To 10)  As Integer

'-----------------------------------    印刷対象　種別      2017.09.28
Dim Print_SYUBETSU()    As String * 2
'-----------------------------------    印刷順(構成)  2017.11.10
Private Type Print_SEQ_Tag
    C_Code      As String * 10
    Print_SEQ   As String * 2
End Type

Private Print_SEQ() As Print_SEQ_Tag
'-----------------------------------    印刷順(作業工程)  2017.11.10
Private KOUTEI()    As String * 2

Dim svHin_Gai       As String           '品番
Dim svSHIMUKE_CODE  As String           '仕向け先
Dim svCATEGORY_CODE As String           '品名ｶﾃｺﾞﾘｰｺｰﾄﾞ




Dim EXCEL_TEMPLATE  As String           'EXCELﾃﾝﾌﾟﾚｰﾄ

Dim Save_Dir        As String           'EXCEL保存先ﾌｫﾙﾀﾞ

Dim HIN_INV         As Boolean          '未登録品番の登録可否

'--------------------------------------- EXCEL用定数
Private Const xlCalculationManual% = -4135
Private Const xlLeft% = -4131
Private Const xlCenter% = -4108
Private Const xlBottom% = -4107
Private Const xlNone% = -4142
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
'--------------------------------------- EXCEL用定数

'--------------------------------------- EXCEL列位置
Private Const xlHin_Name% = 3           '品名
Private Const xlKO_QTY% = 8             '数量
Private Const xlTANI% = 9               '単位
Private Const xlG_ST_URITAN% = 10       '単価
Private Const xlG_KINGAKU% = 11         '金額
'--------------------------------------- EXCEL列位置
 




'Private Const LAST_UPDATE_DAY$ = "[SEI0016] 2017.11.22 15:00"
Private Const LAST_UPDATE_DAY$ = "[SEI0016] 2017.11.29 14:30"

Private Sub Combo1_Change(Index As Integer)
Dim i   As Integer
    
    
    Select Case Index
    
        Case pcmbSHIMUKE
        
        
            If Trim(svHin_Gai) = (Text1(ptxHin_Gai).Text) Then
            Else
                For i = 2 To 5
                    Command1(i).Enabled = False
                Next i
            End If
    
    
    
                        '品名ｶﾃｺﾞﾘｰのセット
'            If ITEM_CATEGORY_Set_Proc() Then
'                Unload Me
'            End If
    
        Case pcmbCATEGORY_Name
    
            If Trim(Right(Combo1(Index).Text, 8)) = Trim(Text1(ptxCATEGORY_CODE).Text) Then
            Else
                Text1(ptxCATEGORY_CODE).Text = Trim(Right(Combo1(Index).Text, 8))
            End If
    End Select

End Sub

Private Sub Combo1_GotFocus(Index As Integer)


    Select Case Index
        Case pcmbSHIMUKE
            svSHIMUKE_CODE = Right(Combo1(pcmbSHIMUKE).Text, 2)
    
        Case pcmbCATEGORY_Name
            svCATEGORY_CODE = Text1(ptxCATEGORY_CODE).Text
    
    End Select

End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If


    Select Case Index
        Case pcmbSHIMUKE
            svSHIMUKE_CODE = Right(Combo1(pcmbSHIMUKE).Text, 2)
    
        Case pcmbCATEGORY_Name
            If Trim(Right(Combo1(Index).Text, 8)) = Trim(Text1(ptxCATEGORY_CODE).Text) Then
            Else
                Text1(ptxCATEGORY_CODE).Text = Trim(Right(Combo1(Index).Text, 8))
            
            End If
    End Select



End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim i   As Integer
    
    
    Select Case Index
        Case pcmbSHIMUKE
        
            If Trim(svSHIMUKE_CODE) = Right(Combo1(pcmbSHIMUKE).Text, 2) Then
            Else
                For i = 2 To 5
                    Command1(i).Enabled = False
                Next i
            End If
                        '品名ｶﾃｺﾞﾘｰのセット
            If ITEM_CATEGORY_Set_Proc() Then
                Unload Me
            End If
        
        
            '品名カテゴリィ
            For i = 0 To Combo1(pcmbCATEGORY_Name).ListCount - 1
                If Trim(Text1(ptxCATEGORY_CODE).Text) = Trim(Right(Combo1(pcmbCATEGORY_Name).List(i), 8)) Then
                    Text1(ptxCATEGORY_CODE).Text = Trim(Right(Combo1(pcmbCATEGORY_Name).List(i), 8))
                    Combo1(pcmbCATEGORY_Name).ListIndex = i
                    Exit For
                End If
            Next i
            If i > Combo1(pcmbCATEGORY_Name).ListCount - 1 Then
                Combo1(pcmbCATEGORY_Name).ListIndex = 0
            End If
        
        
        
        
        Case pcmbCATEGORY_Name
            If Trim(Right(Combo1(Index).Text, 8)) = Trim(Text1(ptxCATEGORY_CODE).Text) Then
            Else
                Text1(ptxCATEGORY_CODE).Text = Trim(Right(Combo1(Index).Text, 8))
            End If
    End Select
End Sub

Private Sub Command1_Click(Index As Integer)


Dim ans     As Integer
Dim i       As Integer

Dim MESG    As String

Dim yn      As Integer

Dim gyo_su  As Integer



    Select Case Index
    
        Case 0      '終了
            Unload Me
    
        Case 1      '検索（表示）
        
        
            If Detail_Disp_Proc() Then
                Unload Me
            End If
        
            Text1(ptxCATEGORY_CODE).SetFocus
        
        
        Case 2      '保存
            
            
            
            
            For i = ptxTanto_Code To ptxN_CLASS_CODE
            
                If Error_Check_Proc(i) Then
                    Exit Sub
                End If
            
            
            Next i
            If Grid_Error_Check_Proc() Then
                Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
                
            
            '    TDBGrid1(pGrdKOUSEI).Refresh
                TDBGrid1(pGrdKOUSEI).Update
            
                TDBGrid1(pGrdKOUSEI).SetFocus
                Exit Sub
            End If
        
            
            
            
            If LenB(StrConv(RTrim(RichTextBox1(prchBIKOU).Text), vbFromUnicode)) > 120 Then
                yn = MsgBox("備考が桁数オーバーしています(最大120文字)、オーバした文字は切り捨てられます。", vbYesNo, "確認入力")
                If yn = vbNo Then
                    RichTextBox1(prchBIKOU).SetFocus
                    Exit Sub
                End If
            End If
            
            gyo_su = SendMessage(RichTextBox1(prchBIKOU).hwnd, EM_GETLINECOUNT, 0&, 0&)
            If gyo_su > 5 Then
                MsgBox "備考最大印字行数は５行です。内容を確認して下さい。"
                RichTextBox1(prchBIKOU).SetFocus
                Exit Sub
            End If
            
'>>>>>>>>>  2017.07.08
'            MESG = "商品化構成データを保存します。" & vbCrLf
'            MESG = MESG & "　　種別／事業部／品番／員数" & vbCrLf
'            MESG = MESG & "　　指図票備考" & vbCrLf
'            MESG = MESG & "よろしいですか？" & vbCrLf
        
        
            If TANKA_KEISAN_Proc() Then     '2017.11.21
                Unload Me                   '2017.11.21
            End If                          '2017.11.21
        
        
            
            MESG = "商品化構成データを保存し、単価を更新します。" & vbCrLf & vbCrLf
            MESG = MESG & "　種別／事業部／品番／員数／指図票備考" & vbCrLf & vbCrLf
            
            
            MESG = MESG & "　合計金額:" & lblGOUKEI_KIN & vbCrLf
            MESG = MESG & "　設定日:" & Format(Now, "YYYY/MM/DD") & vbCrLf
            MESG = MESG & "　担当者:" & Text1(ptxTanto_Code).Text & " " & Text1(ptxTanto_Name).Text
            MESG = MESG & "よろしいですか？" & vbCrLf
        
        
        
        
'>>>>>>>>>  2017.07.08
        
        
            ans = MsgBox(MESG, vbYesNo + vbDefaultButton2 + vbExclamation, "商品化構成の保存確認")
            If ans = vbYes Then
'                If Tanka_Update_Proc() Then     '2017.07.08
'                    Unload Me                   '2017.07.08
'                End If                          '2017.07.08
                
                
                
                If Update_Proc() Then
                    Unload Me
                End If
            
            
            
            
                
                '印刷順に並べ替え   2017.11.10
                KOUSEI.QuickSort Min_Row, KOUSEI.UpperBound(1), ColPRINT_SEQ, XORDER_ASCEND, XTYPE_STRING, ColKO_S_HIN_GAI, XORDER_ASCEND, XTYPE_STRING
                
                Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
                
                TDBGrid1(pGrdKOUSEI).ReBind
                TDBGrid1(pGrdKOUSEI).Update
                TDBGrid1(pGrdKOUSEI).MoveFirst
                '印刷順に並べ替え   2017.11.10
            
            
            
            
                If Estimate_Proc() Then         '2017.07.08
                    Unload Me                   '2017.07.08
                End If                          '2017.07.08
            
            
            
            
                If Detail_Disp_Proc() Then
                    Unload Me
                End If
            
            End If
        
            Command1(4).Enabled = True          '2013.01.17
                    
            Text1(ptxTanto_Code).SetFocus
        
        Case 3      '単価計算
        
            For i = ptxTanto_Code To ptxN_CLASS_CODE
            
                If Error_Check_Proc(i) Then
                    Exit Sub
                End If
            
            Next i
        
            If Grid_Error_Check_Proc() Then
                Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
                
            
            '    TDBGrid1(pGrdKOUSEI).Refresh
                TDBGrid1(pGrdKOUSEI).Update
            
                TDBGrid1(pGrdKOUSEI).SetFocus
                Exit Sub
            End If
        
            If TANKA_KEISAN_Proc() Then
                Unload Me
            End If
        
            Command1(4).Enabled = True          '2013.01.17
        
        Case 4      '見積書発行
            
            If Estimate_Proc() Then
                Unload Me
            End If
        
        Case 5      '単価登録
            
            For i = ptxTanto_Code To ptxKANRI_KIN
            
                If Error_Check_Proc(i) Then
                    Exit Sub
                End If
            
            
            Next i
            
            
            If Grid_Error_Check_Proc() Then
                Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
                
            
            '    TDBGrid1(pGrdKOUSEI).Refresh
                TDBGrid1(pGrdKOUSEI).Update
            
                TDBGrid1(pGrdKOUSEI).SetFocus
                Exit Sub
            End If
            
            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2011.12.21
            If TANKA_KEISAN_Proc() Then
                Unload Me
            End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2011.12.21
            
            
            MESG = "単価を登録します。よろしいですか？" & vbCrLf
            MESG = MESG & "合計金額:" & lblGOUKEI_KIN & vbCrLf
            MESG = MESG & "設定日:" & Format(Now, "YYYY/MM/DD") & vbCrLf
            MESG = MESG & "担当者:" & Text1(ptxTanto_Code).Text & " " & Text1(ptxTanto_Name).Text

            
            
            
            ans = MsgBox(MESG, vbYesNo + vbDefaultButton1 + vbExclamation, "確認入力")
            If ans = vbYes Then
                If Tanka_Update_Proc() Then
                    Unload Me
                End If
            
                If Detail_Disp_Proc() Then
                    Unload Me
                End If
            
            
            End If
                    
            Command1(4).Enabled = True          '2013.01.17
            
            Text1(ptxTanto_Code).SetFocus
    
    
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

Dim wkVAL   As Variant  '2017.09.28
Dim i       As Integer  '2017.09.28


'    If App.PrevInstance Then
'        Beep
'        MsgBox "同一プログラム実行中です。"
'        End
'    End If


    
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]大阪事　見積書作成処理", Me.hwnd, 0)
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
                                
                                
                                
'---------------------------------------------- '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止します。"
        End
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
'------------------------------------------------------ 見積書　表示フラグ  2017.09.28
    If GetIni(App.EXEName, "PRINT_FLG", App.EXEName, c) Then
        c = "0,0,0,0,0,0,0,0,0,0,0"
    End If


    wkVAL = Split(c, ",", -1)

    If UBound(wkVAL) <> 10 Then
        MsgBox "見積書表示フラグ　初期値 SEI0016.INI [PRINT_FLG] を正しく設定して下さい。"
    End If

    
    For i = 0 To UBound(wkVAL)
    
        Print_FLG(i) = wkVAL(i)
            
    
    Next i

'------------------------------------------------------ 印刷対象種別  2017.09.28
    
    Erase Print_SYUBETSU
    
    If GetIni(App.EXEName, "Print_SYUBETSU", App.EXEName, c) Then
        c = "*"
    End If


    wkVAL = Split(c, ",", -1)


    
    For i = 0 To UBound(wkVAL)
    
        ReDim Preserve Print_SYUBETSU(0 To i)
        Print_SYUBETSU(i) = wkVAL(i)
            
    
    Next i


'------------------------------------------------------ EXCEL用項目

    If GetIni(App.EXEName, "EXCEL_TEMPLATE", App.EXEName, c) Then
        EXCEL_TEMPLATE = ""
    Else
        EXCEL_TEMPLATE = Trim(c)
    End If


    If GetIni(App.EXEName, "save_dir", App.EXEName, c) Then
        Save_Dir = ""
    Else
        Save_Dir = Trim(c)
    End If



'------------------------------------------------------ 印刷順      2017.11.10
    Erase Print_SEQ

    i = 0
    Do
        i = i + 1
        If GetIni(App.EXEName, "C_CODE" & Format(i, "00"), App.EXEName, c) Then
        
        Else
            If Trim(c) = "**" Then
                Exit Do
            End If
            
            ReDim Preserve Print_SEQ(0 To i - 1)
            Print_SEQ(i - 1).C_Code = Format(i, "00")
            Print_SEQ(i - 1).Print_SEQ = Trim(c)
        End If
    
    Loop

    Erase KOUTEI

    i = 0
    Do
        i = i + 1
        If GetIni(App.EXEName, "KOUTEI" & Format(i, "00"), App.EXEName, c) Then
        Else
            
            If Trim(c) = "**" Then
                Exit Do
            End If
            
            ReDim Preserve KOUTEI(0 To i - 1)
            KOUTEI(i - 1) = Trim(c)
        End If
    
    
    Loop

                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品名カテゴリマスタＯＰＥＮ
    If ITEM_CATEGORY_Open(BtOpenRead) Then
        Unload Me
    End If
                                '構成マスタＯＰＥＮ
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenRead) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenRead) Then
        Unload Me
    End If
                                '品目単価変更履歴ＯＰＥＮ
    If ITEM_HST_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '大阪事　見積用品目マスタ
    If ITEM_O_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '(一次)大阪事　見積用品目マスタＯＰＥＮ
    If tmpITEM_O_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                
                                '構成マスタＯＰＥＮ
    If wP_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
    
    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc
    
    '仕向け先のセット
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD, 0) Then
        Unload Me
    End If
    Combo1(pcmbSHIMUKE).ListIndex = 0


    '品名ｶﾃｺﾞﾘｰのセット
    If ITEM_CATEGORY_Set_Proc() Then
        Unload Me
    End If




    '種別セット
    If SYUBETSU_Set_Proc() Then
        Unload Me
    End If

    '事魚部セット
    If JGYOBU_Set_Proc() Then
        Unload Me
    End If

'    Load SEI00162          2017.09.28

    SEI00161.Caption = SEI00161.Caption & " " & LAST_UPDATE_DAY

    Call Init_Proc

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
Dim yn  As Integer
                                            
                                            
    yn = MsgBox("終了しますか？", vbYesNo, "確認入力")
    If yn = vbNo Then
        Cancel = True
        Exit Sub
    End If
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
    
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    Set SEI00161 = Nothing
'    Set SEI00162 = Nothing     2017.09.28

    End
End Sub

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    SEI00161.MousePointer = vbHourglass

    Call Ctrl_Lock(SEI00161)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(SEI00161)


    SEI00161.MousePointer = vbDefault

End Sub


Private Sub SHORI_Click(Index As Integer)
    Select Case Index
        Case 0 To 5
            Command1(Index).Value = True

        Case 6      '画面印刷
        
            Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)

    End Select
                    
    
    


End Sub






Private Function Init_Proc(Optional Start_Pos As Integer = 0) As Integer
'----------------------------------------------------------------------------
'                   画面初期化
'----------------------------------------------------------------------------
Dim i           As Integer

Dim Row         As Integer

Dim c           As String * 128
                                
                                
                                

                                
                                
                                
    Init_Proc = True
                                
                                
    For i = Start_Pos To ptxN_CLASS_CODE
        Text1(i).Text = ""
    Next i
                                
                                
    For i = prchBIKOU To prchM_BIKOU
        RichTextBox1(i).Text = ""
    Next i
                                
                                
    For i = pcmbCATEGORY_Name To pcmbCATEGORY_Name
        Combo1(i).ListIndex = -1
    Next i
                                
    If SYUBETSU_Set_Proc() Then
        Exit Function
    End If
                                
                                
    If JGYOBU_Set_Proc() Then
        Exit Function
    End If
                                
    
    
    
    Init_Proc = True


End Function
Private Function ITEM_CATEGORY_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   品名カテゴリィーマスタをドロップダウンリストにセットする。
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer



Dim i           As Integer
    
    ITEM_CATEGORY_Set_Proc = True
    
    Call UniCode_Conv(K0_ITEM_CATEGORY.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM_CATEGORY.CATEGORY_CODE, "")


    Combo1(pcmbCATEGORY_Name).Clear


    Combo1(pcmbCATEGORY_Name).AddItem "なし" & Space(76) & Space(8)


    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
                If StrConv(ITEM_CATEGORYREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Then

                    Exit Do

                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品名カテゴリマスタ")
                Exit Function
        
        End Select

        
        Combo1(pcmbCATEGORY_Name).AddItem StrConv(ITEM_CATEGORYREC.CATEGORY_NAME, vbUnicode) & StrConv(ITEM_CATEGORYREC.CATEGORY_CODE, vbUnicode)
        
        com = BtOpGetNext
    
    Loop



    If Combo1(pcmbCATEGORY_Name).ListCount > 1 Then
        Combo1(pcmbCATEGORY_Name).ListIndex = 0
    End If

    ITEM_CATEGORY_Set_Proc = False
    



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
Private Function JGYOBU_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   事業部をドロップダウンリストにセットする。
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer



Dim i           As Integer
    
    JGYOBU_Set_Proc = True
    
    Set JGYOBU = Nothing
    
    i = 0
    Do
        If i > UBound(JGYOBU_T) Then
            Exit Do
        End If
        
        i = i + 1
        
        JGYOBU.ReDim 1, i, 0, 0
        JGYOBU(i, 0) = Trim(JGYOBU_T(i - 1).NAME) & "            " & Trim(JGYOBU_T(i - 1).CODE)
    
    Loop
    
    
    

    Set TDBDropDown2.Array = JGYOBU
    TDBDropDown2.ReBind

    JGYOBU_Set_Proc = False
    



End Function



Private Sub TDBGrid1_AfterColUpdate(Index As Integer, ByVal ColIndex As Integer)

Dim sts             As Integer
Dim Bookmark        As Variant
    
    
Dim i               As Integer
Dim j               As Integer
    
    
Dim wkDouble        As Double
    
    
 Dim c              As String * 128
    
    
    If TDBGrid1(pGrdKOUSEI).Bookmark = Null Then
        Exit Sub
    End If
    
    If TDBGrid1(pGrdKOUSEI).Bookmark <= 0 Then
        Exit Sub
    End If
    
                    
                    
                    
    Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
    TDBGrid1(pGrdKOUSEI).Update
                    
                    
                    
    Select Case ColIndex
        
        Case ColKO_SYUBETSU
        
                       
            If GetIni(App.EXEName, "C_CODE" & Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU), 2), App.EXEName, c) Then
                c = ""
            End If
                 
                    
            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColPRINT_SEQ) = Trim(c)
        
        Case ColKO_JGYOBU
        Case ColKO_S_HIN_GAI
        
            ' 指図票品番の削除
            If Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_S_HIN_GAI)) = "" And _
                Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU)) = "" Then
                
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_SYUBETSU) = ""
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU) = ""
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_S_HIN_GAI) = ""
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_NAME) = ""
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = ""
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN) = ""
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = ""
            
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_KINGAKU) = ""
                
'                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_UMU) = ""  2017.09.27
                
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_KINGAKU) = ""
                    
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_T_HIN_NAME) = ""    '2017.09.28
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_TANI) = ""          '2017.09.28
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_T_TANKA) = ""       '2017.09.28
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_T_KINGAKU) = ""     '2017.09.28
            
            
            
            
            Else
                
                
                '品番
                Call UniCode_Conv(K0_ITEM.JGYOBU, Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU), 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                
                '2013.01.17
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_S_HIN_GAI) = StrConv(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_S_HIN_GAI), vbUpperCase)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_S_HIN_GAI))
            
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        
                        If HIN_INV Then
                            '未登録品番　可　資材としておく
                            Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                            Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                            Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録品番")
                            Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "00000000.00")
                            Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
                        Else
                            MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]行目 入力した項目はエラーです。(品番)"
                            Exit Sub
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                        Unload Me
                
                End Select
                
                
                
                
                
                Call UniCode_Conv(ITEM_O_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
                Call UniCode_Conv(ITEM_O_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
                Call UniCode_Conv(ITEM_O_REC.HIN_GAI, Text1(ptxHin_Gai).Text)
                
'                Call UniCode_Conv(ITEM_O_REC.KO_JGYOBU, Right(KOUSEI(i, ColKO_JGYOBU), 1))                             '2017.10.16
                Call UniCode_Conv(ITEM_O_REC.KO_JGYOBU, Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU), 1))  '2017.10.16
                
                Call UniCode_Conv(ITEM_O_REC.KO_NAIGAI, NAIGAI_NAI)
                'Call UniCode_Conv(ITEM_O_REC.KO_HIN_GAI, KOUSEI(i, ColKO_S_HIN_GAI))                                   '2017.10.16
                Call UniCode_Conv(ITEM_O_REC.KO_HIN_GAI, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_S_HIN_GAI))        '2017.10.16
               
                sts = BTRV(BtOpGetEqual, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    
                        Call Rclr_ITEM_O_REC
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "大阪事　見積用品目マスタ")
                        Unload Me
                
                End Select
                
                
                '品名
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_HIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                '員数
                If KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = "" Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = "1.00"
                End If
                
                '仕入単価
                If Not IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN) = Format(0, "#0.00")
                Else
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN) = Format(Val(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)), "#0.00")
                End If
                
                '売上単価
                If Not IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = Format(0, "#0.00")
                Else
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = Format(Val(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
                End If
                
                
                '合計金額
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_KINGAKU) = ToHalfAdjust(CCur(CCur(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CCur(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN))), 2)
                        
'               >>>>>>> 削除    2017.09.28
                '子部品　有無
'                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_UMU) = ""  2017.09.28
'                '子部品　チェック
'                Call UniCode_Conv(K0_wP_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
'                Call UniCode_Conv(K0_wP_COMPO.JGYOBU, Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU), 1))
'                Call UniCode_Conv(K0_wP_COMPO.NAIGAI, NAIGAI_NAI)
'                Call UniCode_Conv(K0_wP_COMPO.HIN_GAI, KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_S_HIN_GAI))
'
'                Call UniCode_Conv(K0_wP_COMPO.DATA_KBN, P_KOSOU)
'                Call UniCode_Conv(K0_wP_COMPO.SEQNO, "000")
'
'                sts = BTRV(BtOpGetGreaterEqual, wP_COMPO_POS, wP_COMPO_K_REC, Len(wP_COMPO_K_REC), K0_wP_COMPO, Len(K0_wP_COMPO), 0)
'                Select Case sts
'                   Case BtNoErr
'
'
'                        If StrConv(wP_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
'                            StrConv(wP_COMPO_K_REC.JGYOBU, vbUnicode) <> Right(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_JGYOBU), 1) Or _
'                            StrConv(wP_COMPO_K_REC.NAIGAI, vbUnicode) <> NAIGAI_NAI Or _
'                            Trim(StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_S_HIN_GAI)) Then
'
'                        Else
'                            '子部品　有無
'                            KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_UMU) = "  ▽"
'
'                        End If
'
'
'                    Case BtErrEOF
'                    Case Else
'                        Call File_Error(sts, BtOpGetNext, "構成マスタ")
'                        Unload Me
'                End Select
'               >>>>>>> 削除    2017.09.28
'
            
            
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_T_HIN_NAME) = StrConv(ITEM_O_REC.T_HIN_NAME, vbUnicode)     '2017.09.28
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_TANI) = StrConv(ITEM_O_REC.TANI, vbUnicode)                 '2017.09.28
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_T_TANKA) = Format(Val(StrConv(ITEM_O_REC.TANI, vbUnicode)), "#0.00")      '2017.09.28
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_T_KINGAKU) = ToHalfAdjust(CCur(CCur(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CCur(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_T_TANKA))), 2)
            
            
            
            
            
            End If
                

        Case ColKO_QTY
            If KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = "" Then
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = "1.00"
            End If

            If Not IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) Then
                MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]行目 入力した項目はエラーです。(員数)"
            Else
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = Format(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY), "0.00")
    
                '合計金額
                If IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_KINGAKU) = ToHalfAdjust(CCur(CCur(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CCur(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN))), 2)
                Else
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_KINGAKU) = "0"
                End If
                '提出金額   '2017.09.28
                If IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_T_TANKA)) Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_T_KINGAKU) = ToHalfAdjust(CCur(CCur(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CCur(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_T_TANKA))), 2)
                Else
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_T_KINGAKU) = "0"
                End If
                
            End If


        Case ColKO_T_TANKA  '提出単価   2017.09.29
            If Not IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_T_TANKA)) Then
                MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]行目 入力した項目はエラーです。(提出単価)"
            Else
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_T_TANKA) = Format(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_T_TANKA), "#0.00")
                
                If Not IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) Then
                    KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY) = "1.00"
                End If
            
                
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_T_KINGAKU) = ToHalfAdjust(CCur(CCur(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CCur(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_T_TANKA))), 2)
            End If

        Case ColG_ST_SHITAN

            If Not IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN)) Then
                MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]行目 入力した項目はエラーです。(仕入単価)"
            Else
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN) = Format(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_SHITAN), "#0.00")
            End If


        Case ColG_ST_URITAN

            If Not IsNumeric(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN)) Then
                MsgBox "[" & Format(TDBGrid1(pGrdKOUSEI).Bookmark, "0") & "]行目 入力した項目はエラーです。(売上単価)"
            Else
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN) = Format(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN), "#0.00")
    
                '合計金額
                KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_KINGAKU) = ToHalfAdjust(CCur(CCur(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColKO_QTY)) * CCur(KOUSEI(TDBGrid1(pGrdKOUSEI).Bookmark, ColG_ST_URITAN))), 2)
            End If


    End Select


    Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
    

    TDBGrid1(pGrdKOUSEI).Refresh
    TDBGrid1(pGrdKOUSEI).Update

    TDBGrid1(pGrdKOUSEI).SetFocus


End Sub



Private Sub TDBGrid1_BeforeInsert(Index As Integer, Cancel As Integer)
    
    KOUSEI.ReDim Min_Row, KOUSEI.Count(1), Min_Col, Max_Col

End Sub

Private Sub TDBGrid1_DblClick(Index As Integer)

    If TDBGrid1(pGrdKOUSEI).Bookmark = Null Then
        Exit Sub
    End If
    
    If TDBGrid1(pGrdKOUSEI).Bookmark <= 0 Then
        Exit Sub
    End If

    Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
    TDBGrid1(pGrdKOUSEI).Update

'    SEI00162.Show vbModal  2017.09.28


    Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
    TDBGrid1(pGrdKOUSEI).ReBind
    
    TDBGrid1(pGrdKOUSEI).Update




End Sub

Private Sub Text1_Change(Index As Integer)
Dim i   As Integer
    
    
    Select Case Index
        Case ptxHin_Gai
            If Trim(svHin_Gai) = (Text1(ptxHin_Gai).Text) Then
            Else
                For i = 2 To 5
                    Command1(i).Enabled = False
                Next i
            
            
            
            
            End If
    
    
    
    
    End Select



End Sub

Private Sub Text1_GotFocus(Index As Integer)
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If


    
    Select Case Index
        Case ptxHin_Gai
            svHin_Gai = Text1(ptxHin_Gai).Text
        Case ptxCATEGORY_CODE
            svCATEGORY_CODE = Text1(ptxCATEGORY_CODE).Text
    End Select



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

Dim NEW_ITEM    As Integer          '2017.09.28
    
    
    
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
        
        
        
        
        
        
        
        
        Case ptxCATEGORY_CODE               ' 品名ｶﾃｺﾞﾘｰｺｰﾄﾞ
        
            For i = 0 To Combo1(pcmbCATEGORY_Name).ListCount - 1
                If Trim(Text1(Mode).Text) = Trim(Right(Combo1(pcmbCATEGORY_Name).List(i), 8)) Then
                    Combo1(pcmbCATEGORY_Name).ListIndex = i
                    Exit For
                End If
            Next i
            If i > Combo1(pcmbCATEGORY_Name).ListCount - 1 Then
                MsgBox "入力した項目はエラーです。(品名カテゴリー　未登録)"
                Text1(Mode).SetFocus
                Exit Function
            End If



'>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28
        Case ptxNAKANISHI_KIN           '中西工料　金額

            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(中西工料　金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If

        Case ptxSHOHIN_KIN              '商品化工料　金額

            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(商品化工料　金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If

        Case ptxPF_KAKOU_KIN            'PF加工　金額

            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(PF加工　金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If

        Case ptxPE_KAKOU_KIN            'PE加工　金額

            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(PE加工　金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If

        Case ptxPE_SHIZAI_KIN           'PF資材　金額
            
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(PF資材　金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If

        Case ptxHINBAN_LABEL_KIN        '品番表示ﾗﾍﾞﾙ　金額

            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(品番表示ﾗﾍﾞﾙ　金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If


        Case ptxKOUJI_SETSU_KIN         '設置工事説明書　金額

            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(設置工事説明書　金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If


        Case ptxKONPOU_KIN              '梱包材　金額
            
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(梱包材　金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If

        Case ptxFUKU_SHIZAI_KIN         '副資材　金額

            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(副資材　金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If

        Case ptxKONPOU_ASSY_KIN         '梱包ASSY　金額

            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(梱包ASSY　金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If


        Case ptxKANRI_KIN               '管理費　金額

            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(管理費　金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If


        Case ptxBuzai_Tanto_Name             '部材担当者
        
            If Text1(Mode).Text = "" Then
                MsgBox "入力した項目はエラーです。(部材担当者)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        Case ptxNAI_BUHIN               '国内供給区分（3：打切り）
            If Text1(Mode).Text <> "" And Text1(Mode).Text <> "3" Then
                MsgBox "入力した項目はエラーです。(打切り)"
                Text1(Mode).SetFocus
                Exit Function
            End If



        Case ptxNAKANISHI_T_KIN         '中西工料　提出金額
        
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(中西工料　提出金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        
        
        Case ptxSHOHIN_T_KIN            '商品化工料　提出金額
        
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(商品化工料　提出金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        Case ptxPF_KAKOU_T_KIN          'PF加工　提出金額
        
        
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(PF加工　提出金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        Case ptxPE_KAKOU_T_KIN          'PE加工　提出金額
        
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(PE加工　提出金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        
        Case ptxPE_SHIZAI_T_KIN         'PF資材　提出金額
        
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(PF資材　提出金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        Case ptxHINBAN_LABEL_T_KIN      '品番表示ﾗﾍﾞﾙ　提出金額
        
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(品番表示ﾗﾍﾞﾙ　提出金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        Case ptxKOUJI_SETSU_T_KIN       '設置工事説明書　提出金額
        
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(設置工事説明書　提出金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        
        Case ptxKONPOU_T_KIN            '梱包材　提出金額
        
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(梱包材　提出金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        Case ptxFUKU_SHIZAI_T_KIN       '副資材　提出金額
        
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(副資材　提出金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        Case ptxKONPOU_ASSY_T_KIN       '梱包ASSY　提出金額
        
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(梱包ASSY　提出金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If

        
        Case ptxKANRI_T_KIN              '管理費　提出金額
    
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0")
            Else
                MsgBox "入力した項目はエラーです。(管理費　提出金額)"
                Text1(Mode).SetFocus
                Exit Function
            End If

'>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28

'>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07
        Case ptxNAKANISHI_QTY           '中西工料　数量
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(中西工料　数量)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode).Text) And IsNumeric(Text1(Mode + 2).Text) Then
                Text1(Mode + 3).Text = Format(CDbl(Text1(Mode).Text) * CDbl(Text1(Mode + 2).Text), "#0")
            End If
        
        
        Case ptxNAKANISHI_TANI          '中西工料　単位
        Case ptxNAKANISHI_T_TAN         '中西工料　提出単価
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(中西工料　提出単価)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode - 2).Text) And IsNumeric(Text1(Mode).Text) Then
                Text1(Mode + 1).Text = Format(CDbl(Text1(Mode - 2).Text) * CDbl(Text1(Mode).Text), "#0")
            End If


        Case ptxSHOHIN_QTY              '商品化工料　数量
        
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(商品化工料　数量)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode).Text) And IsNumeric(Text1(Mode + 2).Text) Then
                Text1(Mode + 3).Text = Format(CDbl(Text1(Mode).Text) * CDbl(Text1(Mode + 2).Text), "#0")
            End If
        
        Case ptxSHOHIN_TANI             '商品化工料　単位
        Case ptxSHOHIN_T_TAN            '商品化工料　提出単価
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(商品化工料　提出単価)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode - 2).Text) And IsNumeric(Text1(Mode).Text) Then
                Text1(Mode + 1).Text = Format(CDbl(Text1(Mode - 2).Text) * CDbl(Text1(Mode).Text), "#0")
            End If


        Case ptxPF_KAKOU_QTY            'PF加工　数量
        
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(PF加工　数量)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode).Text) And IsNumeric(Text1(Mode + 2).Text) Then
                Text1(Mode + 3).Text = Format(CDbl(Text1(Mode).Text) * CDbl(Text1(Mode + 2).Text), "#0")
            End If
        
        
        Case ptxPF_KAKOU_TANI           'PF加工　単位
        Case ptxPF_KAKOU_T_TAN          'PF加工　提出単価

            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(PF加工　提出単価)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode - 2).Text) And IsNumeric(Text1(Mode).Text) Then
                Text1(Mode + 1).Text = Format(CDbl(Text1(Mode - 2).Text) * CDbl(Text1(Mode).Text), "#0")
            End If


        Case ptxPE_KAKOU_QTY            'PE加工　数量
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(PE加工　数量)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode).Text) And IsNumeric(Text1(Mode + 2).Text) Then
                Text1(Mode + 3).Text = Format(CDbl(Text1(Mode).Text) * CDbl(Text1(Mode + 2).Text), "#0")
            End If
        Case ptxPE_KAKOU_TANI           'PE加工　単位
        Case ptxPE_KAKOU_T_TAN          'PE加工　提出単価

            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(PE加工　提出単価)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode - 2).Text) And IsNumeric(Text1(Mode).Text) Then
                Text1(Mode + 1).Text = Format(CDbl(Text1(Mode - 2).Text) * CDbl(Text1(Mode).Text), "#0")
            End If


        Case ptxPE_SHIZAI_QTY           'PE資材　数量
        
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(PE資材　数量)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode).Text) And IsNumeric(Text1(Mode + 2).Text) Then
                Text1(Mode + 3).Text = Format(CDbl(Text1(Mode).Text) * CDbl(Text1(Mode + 2).Text), "#0")
            End If
        
        Case ptxPE_SHIZAI_TANI          'PE資材　単位
        Case ptxPE_SHIZAI_T_TAN         'PE資材　提出単価

            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(PE資材　提出単価)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode - 2).Text) And IsNumeric(Text1(Mode).Text) Then
                Text1(Mode + 1).Text = Format(CDbl(Text1(Mode - 2).Text) * CDbl(Text1(Mode).Text), "#0")
            End If


        Case ptxHINBAN_LABEL_QTY        '品番表示ﾗﾍﾞﾙ　数量
        
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(品番表示ﾗﾍﾞﾙ　数量)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode).Text) And IsNumeric(Text1(Mode + 2).Text) Then
                Text1(Mode + 3).Text = Format(CDbl(Text1(Mode).Text) * CDbl(Text1(Mode + 2).Text), "#0")
            End If
        
        
        Case ptxHINBAN_LABEL_TANI       '品番表示ﾗﾍﾞﾙ　単位
        Case ptxHINBAN_LABEL_T_TAN      '品番表示ﾗﾍﾞﾙ　提出単価

            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(品番表示ﾗﾍﾞﾙ　提出単価)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode - 2).Text) And IsNumeric(Text1(Mode).Text) Then
                Text1(Mode + 1).Text = Format(CDbl(Text1(Mode - 2).Text) * CDbl(Text1(Mode).Text), "#0")
            End If



        Case ptxKOUJI_SETSU_QTY         '設置工事説明書　数量
        
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(設置工事説明書　数量)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode).Text) And IsNumeric(Text1(Mode + 2).Text) Then
                Text1(Mode + 3).Text = Format(CDbl(Text1(Mode).Text) * CDbl(Text1(Mode + 2).Text), "#0")
            End If
        
        
        
        Case ptxKOUJI_SETSU_TANI        '設置工事説明書　単位
        Case ptxKOUJI_SETSU_T_TAN       '設置工事説明書　提出単価
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(設置工事説明書　提出単価)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode - 2).Text) And IsNumeric(Text1(Mode).Text) Then
                Text1(Mode + 1).Text = Format(CDbl(Text1(Mode - 2).Text) * CDbl(Text1(Mode).Text), "#0")
            End If


        Case ptxKONPOU_QTY              '梱包材　数量
        
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(梱包材　数量)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode).Text) And IsNumeric(Text1(Mode + 2).Text) Then
                Text1(Mode + 3).Text = Format(CDbl(Text1(Mode).Text) * CDbl(Text1(Mode + 2).Text), "#0")
            End If
        
        Case ptxKONPOU_TANI             '梱包材　単位
        Case ptxKONPOU_T_TAN            '梱包材　提出単価

            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(梱包材　提出単価)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode - 2).Text) And IsNumeric(Text1(Mode).Text) Then
                Text1(Mode + 1).Text = Format(CDbl(Text1(Mode - 2).Text) * CDbl(Text1(Mode).Text), "#0")
            End If



        Case ptxFUKU_SHIZAI_QTY         '副資材　数量
        
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(副資材　数量)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode).Text) And IsNumeric(Text1(Mode + 2).Text) Then
                Text1(Mode + 3).Text = Format(CDbl(Text1(Mode).Text) * CDbl(Text1(Mode + 2).Text), "#0")
            End If
        
        
        Case ptxFUKU_SHIZAI_TANI        '副資材　単位
        Case ptxFUKU_SHIZAI_T_TAN       '副資材　提出単価

            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(副資材　提出単価)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode - 2).Text) And IsNumeric(Text1(Mode).Text) Then
                Text1(Mode + 1).Text = Format(CDbl(Text1(Mode - 2).Text) * CDbl(Text1(Mode).Text), "#0")
            End If


        Case ptxKONPOU_ASSY_QTY         '梱包ASSY　数量
        
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(梱包ASSY　数量)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode).Text) And IsNumeric(Text1(Mode + 2).Text) Then
                Text1(Mode + 3).Text = Format(CDbl(Text1(Mode).Text) * CDbl(Text1(Mode + 2).Text), "#0")
            End If
        
        
        Case ptxKONPOU_ASSY_TANI        '梱包ASSY　単位
        Case ptxKONPOU_ASSY_T_TAN       '梱包ASSY　提出単価


            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(梱包ASSY　提出単価)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode - 2).Text) And IsNumeric(Text1(Mode).Text) Then
                Text1(Mode + 1).Text = Format(CDbl(Text1(Mode - 2).Text) * CDbl(Text1(Mode).Text), "#0")
            End If


        Case ptxKANRI_QTY               '管理費　数量
        
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(管理費　数量)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode).Text) And IsNumeric(Text1(Mode + 2).Text) Then
                Text1(Mode + 3).Text = Format(CDbl(Text1(Mode).Text) * CDbl(Text1(Mode + 2).Text), "#0")
            End If
        
        
        Case ptxKANRI_TANI              '管理費　単位
        Case ptxKANRI_T_TAN             '管理費　提出単価

            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0"
            End If
            
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(CDbl(Text1(Mode).Text), "#0.00")
            Else
                MsgBox "入力した項目はエラーです。(管理費　提出単価)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(Mode - 2).Text) And IsNumeric(Text1(Mode).Text) Then
                Text1(Mode + 1).Text = Format(CDbl(Text1(Mode - 2).Text) * CDbl(Text1(Mode).Text), "#0")
            End If





'>>>>>>>>>>>>>>>>>>>>>>>    2017.11.07


            
    
    
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
        
            '商品化ｸﾗｽ  2017.09.28
            'Text1(ptxS_CLASS_CODE).Text = Trim(StrConv(P_COMPO_O_REC.CLASS_CODE, vbUnicode))
            '付加ｸﾗｽ    2017.09.28
            'Text1(ptxF_CLASS_CODE).Text = Trim(StrConv(P_COMPO_O_REC.F_CLASS_CODE, vbUnicode))
            '内職ｸﾗｽ    2017.09.28
            'Text1(ptxN_CLASS_CODE).Text = Trim(StrConv(P_COMPO_O_REC.N_CLASS_CODE, vbUnicode))

        
        Case BtErrKeyNotFound
            
            FAST_FLG = False
            
            '備考
            RichTextBox1(prchBIKOU).Text = ""
        
            '商品化ｸﾗｽ  2017.09.28
            'Text1(ptxS_CLASS_CODE).Text = ""
            '付加ｸﾗｽ    2017.09.28
            'Text1(ptxF_CLASS_CODE).Text = ""
            '内職ｸﾗｽ    2017.09.28
            'Text1(ptxN_CLASS_CODE).Text = ""
        
        
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
            
            
            
            
            Row = Row + 1
                        
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
            
            
            
        Loop
    End If


    
    
    Call UniCode_Conv(K0_ITEM_O.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM_O.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM_O.HIN_GAI, Text1(ptxHin_Gai).Text)
    Call UniCode_Conv(K0_ITEM_O.SEQ_NO, "0000")

    com = BtOpGetGreater
    
    Do
        DoEvents
        
        sts = BTRV(com, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
        Select Case sts
            Case BtNoErr
            
                            
                If StrConv(ITEM_O_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                    StrConv(ITEM_O_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                    Trim(StrConv(ITEM_O_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHin_Gai).Text) Then
                
                    Exit Do
            
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock             '2008.01.15
                Call File_Error(sts, BtOpGetNext, "構成マスタ")
                Exit Function
        End Select
        
        
        
        
        If Trim(StrConv(ITEM_O_REC.KO_HIN_GAI, vbUnicode)) = "" Then
        
            Row = Row + 1
                        
            If Grid_Set_ITEM_O_Proc(Row) Then
                Exit Function
            End If
        End If
        
        com = BtOpGetNext
        
        
    Loop




    If Row < 49 Then
        For Row = Row + 1 To 50

            KOUSEI.ReDim Min_Row, Row, Min_Col, Max_Col
            KOUSEI(Row, ColNO) = Row
            KOUSEI(Row, ColKO_S_HIN_GAI) = ""
        Next Row
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

Dim sts         As Integer
Dim i           As Integer
Dim j           As Integer
    
Dim com         As Integer
    
Dim NEW_ITEM    As Integer      '2017.09.27
    
    
    Grid_Set_Proc = True

    

    KOUSEI.ReDim Min_Row, Row, Min_Col, Max_Col
    
    'No
    KOUSEI(Row, ColNO) = Row
    
    
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
    
    
    '事業部
    For i = 0 To UBound(JGYOBU_T)
    
        If Trim(JGYOBU_T(i).CODE) = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode) Then
            KOUSEI(Row, ColKO_JGYOBU) = Trim(JGYOBU_T(i).NAME) & "            " & Trim(JGYOBU_T(i).CODE)
            Exit For
        End If
    Next i
    
    '指図票品番
    KOUSEI(Row, ColKO_S_HIN_GAI) = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
        
    '品名
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Select Case sts
        Case BtNoErr
            KOUSEI(Row, ColKO_HIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
        Case BtErrKeyNotFound
            KOUSEI(Row, ColKO_HIN_NAME) = ""
            
        
            Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "00000000.00")
            Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    
    End Select
    
    
    
    '>>>>>>>>>>>>>>>>>>>    大阪事　見積用品目マスタ読込み  2017.09.28
    Call UniCode_Conv(K0_ITEM_O.JGYOBU, StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM_O.NAIGAI, StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM_O.HIN_GAI, StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode))
    
    Call UniCode_Conv(K0_ITEM_O.SEQ_NO, Format(Row, "000"))
    
    
    
    sts = BTRV(BtOpGetEqual, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
        
    Select Case sts
        Case BtNoErr
            NEW_ITEM = 0
        Case BtErrKeyNotFound
            NEW_ITEM = 1
        Case Else
            Call File_Error(sts, BtOpGetEqual, "大阪事　見積用品目マスタ")
            Exit Function
    
    End Select
    '>>>>>>>>>>>>>>>>>>>    大阪事　見積用品目マスタ読込み  2017.09.28
       
    
    
    
    
    
    
    
    
    
    '提出品名   2017.09.28
    If NEW_ITEM = 0 Then
        KOUSEI(Row, ColKO_T_HIN_NAME) = Trim(StrConv(ITEM_O_REC.T_HIN_NAME, vbUnicode))
    Else
        KOUSEI(Row, ColKO_T_HIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
    End If
    
    
    
    
    '員数
    If IsNumeric(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) Then
        KOUSEI(Row, ColKO_QTY) = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
    Else
        KOUSEI(Row, ColKO_QTY) = "1.00"
    End If
    
    '単位   2017.09.28
    If NEW_ITEM = 0 Then
        KOUSEI(Row, ColKO_TANI) = Trim(StrConv(ITEM_O_REC.TANI, vbUnicode))
    End If
    '提出単価   2017.09.28
    If NEW_ITEM = 0 Then
        If IsNumeric(StrConv(ITEM_O_REC.T_TANKA, vbUnicode)) Then
            KOUSEI(Row, ColKO_T_TANKA) = Format(CDbl(StrConv(ITEM_O_REC.T_TANKA, vbUnicode)), "#0.00")
        End If
    End If
    '提出金額   2017.09.28
    If NEW_ITEM = 0 Then
        If IsNumeric(StrConv(ITEM_O_REC.T_TANKA, vbUnicode)) Then
            KOUSEI(Row, ColKO_T_KINGAKU) = Format(CDbl(StrConv(ITEM_O_REC.T_KINGAKU, vbUnicode)), "#0")
        End If
    End If
    
    
    
    
    
    
    
    '仕入単価
    If IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
        KOUSEI(Row, ColG_ST_SHITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)), "#0.00")
    Else
        KOUSEI(Row, ColG_ST_SHITAN) = "0.00"
    End If
    
    '売上単価
    If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
        KOUSEI(Row, ColG_ST_URITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
    Else
        KOUSEI(Row, ColG_ST_URITAN) = "0.00"
    End If
    
    
    '合計金額
    KOUSEI(Row, ColG_KINGAKU) = ToHalfAdjust(CCur(CCur(KOUSEI(Row, ColKO_QTY)) * CCur(KOUSEI(Row, ColG_ST_URITAN))), 2)
    
    '印刷順 2017.11.10
    For i = 0 To UBound(Print_SEQ)
        If Trim(Print_SEQ(i).C_Code) = Trim(Right(KOUSEI(Row, ColKO_SYUBETSU), 2)) Then
            KOUSEI(Row, ColPRINT_SEQ) = Print_SEQ(i).Print_SEQ
            Exit For
        End If
    Next i
    
    
    
    '>>>>>>>>>>>>>>>>子部品　チェック   2017.09.28  DEL
'    '子部品　有無
'    KOUSEI(Row, ColKO_UMU) = ""
'
'
'
'    Call UniCode_Conv(K0_wP_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
'    Call UniCode_Conv(K0_wP_COMPO.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
'    Call UniCode_Conv(K0_wP_COMPO.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
'    Call UniCode_Conv(K0_wP_COMPO.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
'
'    Call UniCode_Conv(K0_wP_COMPO.DATA_KBN, P_KOSOU)
'    Call UniCode_Conv(K0_wP_COMPO.SEQNO, "000")
'
'    sts = BTRV(BtOpGetGreaterEqual, wP_COMPO_POS, wP_COMPO_K_REC, Len(wP_COMPO_K_REC), K0_wP_COMPO, Len(K0_wP_COMPO), 0)
'    Select Case sts
'       Case BtNoErr
'
'
'            If StrConv(wP_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
'                StrConv(wP_COMPO_K_REC.JGYOBU, vbUnicode) <> StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode) Or _
'                StrConv(wP_COMPO_K_REC.NAIGAI, vbUnicode) <> StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode) Or _
'                Trim(StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)) Then
'
'            Else
'                '子部品　有無
'                KOUSEI(Row, ColKO_UMU) = "  ▽"
'
'            End If
'
'
'        Case BtErrEOF
'        Case Else
'            Call File_Error(sts, BtOpGetNext, "構成マスタ")
'            Exit Function
'    End Select
    '>>>>>>>>>>>>>>>>子部品　チェック   2017.09.28  DEL
        
    
    
    
    
    
    
    Grid_Set_Proc = False
End Function


Private Function Grid_Set_ITEM_O_Proc(Row As Long) As Integer


'----------------------------------------------------------------------------
'                   大阪事　見積用品目マスタ==>Gridテーブル
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim i           As Integer
Dim j           As Integer
    
Dim com         As Integer
    
Dim NEW_ITEM    As Integer      '2017.09.27
    
    
    Grid_Set_ITEM_O_Proc = True

    

    KOUSEI.ReDim Min_Row, Row, Min_Col, Max_Col
    
    'No
    KOUSEI(Row, ColNO) = Row
    
    
    '種別
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(ITEM_O_REC.KO_SYUBETSU, vbUnicode))
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
    
    
    '事業部
    For i = 0 To UBound(JGYOBU_T)
    
        If Trim(JGYOBU_T(i).CODE) = StrConv(ITEM_O_REC.KO_JGYOBU, vbUnicode) Then
            KOUSEI(Row, ColKO_JGYOBU) = Trim(JGYOBU_T(i).NAME) & "            " & Trim(JGYOBU_T(i).CODE)
            Exit For
        End If
    Next i
    
    
    
    
       
    
    
    
    
    
    
    
    
    
    '提出品名   2017.09.28
    If NEW_ITEM = 0 Then
        KOUSEI(Row, ColKO_T_HIN_NAME) = Trim(StrConv(ITEM_O_REC.T_HIN_NAME, vbUnicode))
    Else
        KOUSEI(Row, ColKO_T_HIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
    End If
    
    
    
    
    '員数
    If IsNumeric(StrConv(ITEM_O_REC.KO_QTY, vbUnicode)) Then
        KOUSEI(Row, ColKO_QTY) = Format(CDbl(StrConv(ITEM_O_REC.KO_QTY, vbUnicode)), "#0.00")
    Else
        KOUSEI(Row, ColKO_QTY) = "1.00"
    End If
    
    '単位
    KOUSEI(Row, ColKO_TANI) = Trim(StrConv(ITEM_O_REC.TANI, vbUnicode))
    '提出単価
    If IsNumeric(StrConv(ITEM_O_REC.T_TANKA, vbUnicode)) Then
        KOUSEI(Row, ColKO_T_TANKA) = Format(CDbl(StrConv(ITEM_O_REC.T_TANKA, vbUnicode)), "#0.00")
    End If
    '提出金額
    If IsNumeric(StrConv(ITEM_O_REC.T_TANKA, vbUnicode)) Then
        KOUSEI(Row, ColKO_T_KINGAKU) = Format(CDbl(StrConv(ITEM_O_REC.T_KINGAKU, vbUnicode)), "#0")
    End If
    
    
    
    
    
    
    
    '仕入単価
    KOUSEI(Row, ColG_ST_SHITAN) = "0.00"
    
    '売上単価
    KOUSEI(Row, ColG_ST_URITAN) = "0.00"
    
    
    '合計金額
    KOUSEI(Row, ColG_KINGAKU) = ToHalfAdjust(CCur(CCur(KOUSEI(Row, ColKO_QTY)) * CCur(KOUSEI(Row, ColG_ST_URITAN))), 2)
    
    '印刷順 2017.11.10
    For i = 0 To UBound(Print_SEQ)
        If Trim(Print_SEQ(i).C_Code) = Trim(Right(KOUSEI(Row, ColKO_SYUBETSU), 2)) Then
            KOUSEI(Row, ColPRINT_SEQ) = Print_SEQ(i).Print_SEQ
            Exit For
        End If
    Next i
    
    
    
    
    
    
    
    
    
    Grid_Set_ITEM_O_Proc = False
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
Dim excelApplication    As Object
Dim excelWorkBook       As Object
Dim excelSheet          As Object

    
Dim i                   As Integer
Dim j                   As Integer
    
Dim Row                 As Integer
    
Dim Fsw                 As Boolean
    
Dim com                 As Integer
Dim sts                 As Integer
    
    
Dim LCnt                As Integer
Dim PCnt                As Integer
    
    
Dim Start_Line          As Integer
Dim End_Line            As Integer
Dim Total_Line          As Integer
    
Dim Clear_Start_Line    As Integer
Dim Clear_End_Line      As Integer
    
Dim Total_Kingaku       As Double
        
Dim Print_FLG           As Integer          '2017.10.06
    
    
    Estimate_Proc = True
    
    
    If Error_Check_Proc(ptxTanto_Code) Then
        Estimate_Proc = False
        Exit Function
    End If
    
    
    
    Call Input_Lock
    
    
    
    Set excelApplication = CreateObject("Excel.Application")
    

    If Trim(EXCEL_TEMPLATE) = "" Then
        Set excelWorkBook = excelApplication.Workbooks.Add
    
    Else
                                                        'ﾃﾝﾌﾟﾚｰﾄﾌﾞｯｸを開く
        Set excelWorkBook = excelApplication.Workbooks.Open(EXCEL_TEMPLATE)
    End If

    Set excelSheet = excelWorkBook.Worksheets(1)
    
    excelSheet.NAME = Trim(Text1(ptxHin_Gai).Text)
    
    
    
    'excelApplication.Visible = True
        
    excelApplication.Calculation = xlCalculationManual
    excelApplication.ScreenUpdating = False
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   開始処理


'-----------------------    パナソニック担当者  2017.09.28
    excelSheet.Application.Cells(6, 5).Value = Text1(ptxBuzai_Tanto_Name).Text

'-----------------------    担当者　2017.07.08
    excelSheet.Application.Cells(2, 11).Value = Text1(ptxTanto_Name).Text





'-----------------------    品名
'    excelSheet.Application.Cells(20, 3).Value = Text1(ptxHin_Name).Text                                                '2017.10.13
    excelSheet.Application.Cells(20, 3).Value = Trim(Text1(ptxHin_Gai).Text) & " " & Trim(Text1(ptxHin_Name).Text)      '2017.10.13

'-----------------------    合計金額
    excelSheet.Application.Cells(20, 10).Value = Val(lblGOUKEI_T_KIN.Caption)
'-----------------------    合計金額
    excelSheet.Application.Cells(44, 11).Value = Val(lblGOUKEI_T_KIN.Caption)



    Row = 22


    LCnt = 1
    PCnt = 1


    Start_Line = 1
    End_Line = 54


    Clear_Start_Line = 23
    Clear_End_Line = 43


    Total_Line = 20

    Total_Kingaku = 0
    For i = 1 To KOUSEI.UpperBound(1)
        
        
        If Trim(KOUSEI(i, ColKO_S_HIN_GAI)) <> "" Or Trim(KOUSEI(i, ColKO_T_HIN_NAME)) <> "" Then
        
 '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>印刷内容を全て画面より 2017.10.06
            Print_FLG = 0
            For j = 0 To UBound(Print_SYUBETSU)
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = Print_SYUBETSU(j) Then
        
                    Print_FLG = 1
            
                End If
            Next j
            If Print_FLG = 1 Then
    
    
   
                If LCnt > PCnt * 21 Then
                    PCnt = PCnt + 1
                    Row = Row + 33
                    LCnt = 0


                    Total_Line = Total_Line + 54
                    excelSheet.Application.Range(excelSheet.Application.Rows(Start_Line), excelSheet.Application.Rows(End_Line)).Copy
                    Start_Line = Start_Line + 54
                    End_Line = End_Line + 54
                    excelSheet.Application.Range(excelSheet.Application.Rows(Start_Line), excelSheet.Application.Rows(Start_Line)).pastespecial


                    Clear_Start_Line = Clear_Start_Line + 54
                    Clear_End_Line = Clear_End_Line + 54

                    excelSheet.Application.Range(excelSheet.Application.Cells(Clear_Start_Line, 3), excelSheet.Application.Cells(Clear_End_Line, 11)).ClearContents

                End If


                Row = Row + 1
                LCnt = LCnt + 1


                Fsw = True
   
        
        
                 excelSheet.Application.Cells(Row, xlHin_Name).Value = KOUSEI(i, ColKO_T_HIN_NAME)
                 excelSheet.Application.Cells(Row, xlKO_QTY).Value = KOUSEI(i, ColKO_QTY)
                 excelSheet.Application.Cells(Row, xlTANI).Value = Trim(KOUSEI(i, ColKO_TANI))
                 'excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value = KOUSEI(i, ColG_ST_URITAN)
                 excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value = KOUSEI(i, ColKO_T_TANKA)
                 excelSheet.Application.Cells(Row, xlG_KINGAKU).Value = KOUSEI(i, ColKO_T_KINGAKU)
        
                
                
            End If
        End If
    Next i
            
            
            
            
            
            
'>>>>>>>>>>>>>>>>>>>>>>>>>  2017.11.10 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'    For i = chkNAKANISHI_F To chkKANRI_F
'        If Check1(i).Value = 1 Then
'
'
'            If LCnt > PCnt * 21 Then
'                PCnt = PCnt + 1
'                Row = Row + 33
'                LCnt = 0
'
'
'                Total_Line = Total_Line + 54
'                excelSheet.Application.Range(excelSheet.Application.Rows(Start_Line), excelSheet.Application.Rows(End_Line)).Copy
'                Start_Line = Start_Line + 54
'                End_Line = End_Line + 54
'                excelSheet.Application.Range(excelSheet.Application.Rows(Start_Line), excelSheet.Application.Rows(Start_Line)).pastespecial
'
'
'                Clear_Start_Line = Clear_Start_Line + 54
'                Clear_End_Line = Clear_End_Line + 54
'
'                excelSheet.Application.Range(excelSheet.Application.Cells(Clear_Start_Line, 3), excelSheet.Application.Cells(Clear_End_Line, 11)).ClearContents
'
'            End If
'
'
'            Row = Row + 1
'            LCnt = LCnt + 1
'
'
'
'
'
'            excelSheet.Application.Cells(Row, xlHin_Name).Value = lblTitle(i).Caption
''>>>>>>>>   2017.10.16
'            Select Case i
'                Case 0
'                    excelSheet.Application.Cells(Row, xlTANI).Value = Text1(10).Text
'                Case 1
'                    excelSheet.Application.Cells(Row, xlTANI).Value = Text1(12).Text
'                Case 2
'                    excelSheet.Application.Cells(Row, xlTANI).Value = Text1(14).Text
'                Case 3
'                    excelSheet.Application.Cells(Row, xlTANI).Value = Text1(16).Text
'                Case 4
'                    excelSheet.Application.Cells(Row, xlTANI).Value = Text1(18).Text
'                Case 5
'                    excelSheet.Application.Cells(Row, xlTANI).Value = Text1(20).Text
'                Case 6
'                    excelSheet.Application.Cells(Row, xlTANI).Value = Text1(22).Text
'                Case 7
'                    excelSheet.Application.Cells(Row, xlTANI).Value = Text1(24).Text
'                Case 8
'                    excelSheet.Application.Cells(Row, xlTANI).Value = Text1(26).Text
'                Case 9
'                    excelSheet.Application.Cells(Row, xlTANI).Value = Text1(28).Text
'                Case 10
'                    excelSheet.Application.Cells(Row, xlTANI).Value = Text1(30).Text
'            End Select
''>>>>>>>>   2017.10.16
'
'            excelSheet.Application.Cells(Row, xlG_KINGAKU).Value = Text1(i + 37).Text
'
'
'
'
'        End If
'
'    Next i
'
'>>>>>>>>>>>>>>>>>>>>>>>>>  2017.11.10 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'>>>>>>>>>>>>>>>>>>>>>>>>>  2017.11.10 工数を任意の印刷順に変更<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
For i = 0 To UBound(KOUTEI)
    If Check1(KOUTEI(i) - 1).Value = vbChecked Then
        
        If LCnt > PCnt * 21 Then
            PCnt = PCnt + 1
            Row = Row + 33
            LCnt = 0


            Total_Line = Total_Line + 54
            excelSheet.Application.Range(excelSheet.Application.Rows(Start_Line), excelSheet.Application.Rows(End_Line)).Copy
            Start_Line = Start_Line + 54
            End_Line = End_Line + 54
            excelSheet.Application.Range(excelSheet.Application.Rows(Start_Line), excelSheet.Application.Rows(Start_Line)).pastespecial


            Clear_Start_Line = Clear_Start_Line + 54
            Clear_End_Line = Clear_End_Line + 54

            excelSheet.Application.Range(excelSheet.Application.Cells(Clear_Start_Line, 3), excelSheet.Application.Cells(Clear_End_Line, 11)).ClearContents

        End If


        Row = Row + 1
        LCnt = LCnt + 1


        excelSheet.Application.Cells(Row, xlHin_Name).Value = lblTitle(KOUTEI(i) - 1).Caption


        Select Case KOUTEI(i) - 1
            Case 0
                excelSheet.Application.Cells(Row, xlKO_QTY).Value = Text1(10).Text
                excelSheet.Application.Cells(Row, xlTANI).Value = Trim(Text1(11).Text)
                excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value = Text1(12).Text
                excelSheet.Application.Cells(Row, xlG_KINGAKU).Value = Val(Text1(13).Text)
            Case 1
                excelSheet.Application.Cells(Row, xlKO_QTY).Value = Text1(15).Text
                excelSheet.Application.Cells(Row, xlTANI).Value = Trim(Text1(16).Text)
                excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value = Text1(17).Text
                excelSheet.Application.Cells(Row, xlG_KINGAKU).Value = Val(Text1(18).Text)
            Case 2
                excelSheet.Application.Cells(Row, xlKO_QTY).Value = Text1(20).Text
                excelSheet.Application.Cells(Row, xlTANI).Value = Trim(Text1(21).Text)
                excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value = Text1(22).Text
                excelSheet.Application.Cells(Row, xlG_KINGAKU).Value = Val(Text1(23).Text)
            Case 3
                excelSheet.Application.Cells(Row, xlKO_QTY).Value = Text1(25).Text
                excelSheet.Application.Cells(Row, xlTANI).Value = Trim(Text1(26).Text)
                excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value = Text1(27).Text
                excelSheet.Application.Cells(Row, xlG_KINGAKU).Value = Val(Text1(28).Text)
            Case 4
                excelSheet.Application.Cells(Row, xlKO_QTY).Value = Text1(30).Text
                excelSheet.Application.Cells(Row, xlTANI).Value = Trim(Text1(31).Text)
                excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value = Text1(32).Text
                excelSheet.Application.Cells(Row, xlG_KINGAKU).Value = Val(Text1(33).Text)
            Case 5
                excelSheet.Application.Cells(Row, xlKO_QTY).Value = Text1(35).Text
                excelSheet.Application.Cells(Row, xlTANI).Value = Trim(Text1(36).Text)
                excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value = Text1(37).Text
                excelSheet.Application.Cells(Row, xlG_KINGAKU).Value = Val(Text1(38).Text)
            Case 6
                excelSheet.Application.Cells(Row, xlKO_QTY).Value = Text1(40).Text
                excelSheet.Application.Cells(Row, xlTANI).Value = Trim(Text1(41).Text)
                excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value = Text1(42).Text
                excelSheet.Application.Cells(Row, xlG_KINGAKU).Value = Val(Text1(43).Text)
            Case 7
                excelSheet.Application.Cells(Row, xlKO_QTY).Value = Text1(45).Text
                excelSheet.Application.Cells(Row, xlTANI).Value = Trim(Text1(46).Text)
                excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value = Text1(47).Text
                excelSheet.Application.Cells(Row, xlG_KINGAKU).Value = Val(Text1(48).Text)
            Case 8
                excelSheet.Application.Cells(Row, xlKO_QTY).Value = Text1(50).Text
                excelSheet.Application.Cells(Row, xlTANI).Value = Trim(Text1(51).Text)
                excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value = Text1(52).Text
                excelSheet.Application.Cells(Row, xlG_KINGAKU).Value = Val(Text1(53).Text)
            Case 9
                excelSheet.Application.Cells(Row, xlKO_QTY).Value = Text1(55).Text
                excelSheet.Application.Cells(Row, xlTANI).Value = Trim(Text1(56).Text)
                excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value = Text1(57).Text
                excelSheet.Application.Cells(Row, xlG_KINGAKU).Value = Val(Text1(58).Text)
            Case 10
                excelSheet.Application.Cells(Row, xlKO_QTY).Value = Text1(60).Text
                excelSheet.Application.Cells(Row, xlTANI).Value = Trim(Text1(61).Text)
                excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value = Text1(62).Text
                excelSheet.Application.Cells(Row, xlG_KINGAKU).Value = Val(Text1(63).Text)
        End Select

    End If

Next i
'>>>>>>>>>>>>>>>>>>>>>>>>>  2017.11.10 工数を任意の印刷順に変更<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

 
 
 
 
 
 
 
 '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>印刷内容を全て画面より 2017.10.06
            

'            Call UniCode_Conv(K0_wP_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
'            Call UniCode_Conv(K0_wP_COMPO.JGYOBU, Right(KOUSEI(i, ColKO_JGYOBU), 1))
'            Call UniCode_Conv(K0_wP_COMPO.NAIGAI, NAIGAI_NAI)
'            Call UniCode_Conv(K0_wP_COMPO.HIN_GAI, KOUSEI(i, ColKO_S_HIN_GAI))
'
'            Call UniCode_Conv(K0_wP_COMPO.DATA_KBN, P_KOSOU)
'            Call UniCode_Conv(K0_wP_COMPO.SEQNO, "000")
'
'            com = BtOpGetGreater
'
'            Fsw = False
'
'
'            Do
'                DoEvents
'                sts = BTRV(com, wP_COMPO_POS, wP_COMPO_K_REC, Len(wP_COMPO_K_REC), K0_wP_COMPO, Len(K0_wP_COMPO), 0)
'                Select Case sts
'                    Case BtNoErr
'
'
'                        If StrConv(wP_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
'                            StrConv(wP_COMPO_K_REC.JGYOBU, vbUnicode) <> Right(KOUSEI(i, ColKO_JGYOBU), 1) Or _
'                            StrConv(wP_COMPO_K_REC.NAIGAI, vbUnicode) <> NAIGAI_NAI Or _
'                            Trim(StrConv(wP_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(KOUSEI(i, ColKO_S_HIN_GAI)) Then
'
'                            Exit Do
'
'                        End If
'
'
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(wP_COMPO_K_REC.KO_JGYOBU, vbUnicode))
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(wP_COMPO_K_REC.KO_NAIGAI, vbUnicode))
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(wP_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
'
'
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                        Select Case sts
'                            Case BtNoErr
'                            Case BtErrKeyNotFound
'                                Call UniCode_Conv(ITEMREC.HIN_GAI, "未登録品番")
'                                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")
'                            Case Else
'                                Call Input_UnLock
'                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
'                                Exit Function
'                        End Select
'
'
'
'
'
'
'
'                        If LCnt > PCnt * 21 Then
'                            PCnt = PCnt + 1
'                            Row = Row + 33
'                            LCnt = 0
'
'
'                            Total_Line = Total_Line + 54
'                            excelSheet.Application.Range(excelSheet.Application.Rows(Start_Line), excelSheet.Application.Rows(End_Line)).Copy
'                            Start_Line = Start_Line + 54
'                            End_Line = End_Line + 54
'                            excelSheet.Application.Range(excelSheet.Application.Rows(Start_Line), excelSheet.Application.Rows(Start_Line)).pastespecial
'
'
'                            Clear_Start_Line = Clear_Start_Line + 54
'                            Clear_End_Line = Clear_End_Line + 54
'
'                            excelSheet.Application.Range(excelSheet.Application.Cells(Clear_Start_Line, 3), excelSheet.Application.Cells(Clear_End_Line, 11)).ClearContents
'
'                        End If
'
'
'                        Row = Row + 1
'                        LCnt = LCnt + 1
'
'
'                        Fsw = True
'
'
'                        excelSheet.Application.Cells(Row, xlHin_Name).Value = RTrim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
'                        excelSheet.Application.Cells(Row, xlTANI).Value = ""
'
'                        If IsNumeric(StrConv(wP_COMPO_K_REC.KO_QTY, vbUnicode)) And _
'                            IsNumeric(KOUSEI(i, ColKO_QTY)) Then
'                            excelSheet.Application.Cells(Row, xlKO_QTY).Value = ToHalfAdjust(CCur(CCur(StrConv(wP_COMPO_K_REC.KO_QTY, vbUnicode)) * CCur(KOUSEI(i, ColKO_QTY))), 2)
'                        Else
'                            excelSheet.Application.Cells(Row, xlKO_QTY).Value = ""
'                        End If
'
'
'                        If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
'                            excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value = CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode))
'                        Else
'                            excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value = ""
'                        End If
'
'                        If IsNumeric(excelSheet.Application.Cells(Row, xlKO_QTY).Value) And _
'                            IsNumeric(excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value) Then
'
'                            excelSheet.Application.Cells(Row, xlG_KINGAKU).Value = ToHalfAdjust(CCur(CCur(excelSheet.Application.Cells(Row, xlKO_QTY).Value) * CCur(excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value)), 2)
'                            Total_Kingaku = Total_Kingaku + excelSheet.Application.Cells(Row, xlG_KINGAKU).Value
'                        End If
'
'                    Case BtErrEOF
'                        Exit Do
'                    Case Else
'                        Call Input_UnLock             '2008.01.15
'                        Call File_Error(sts, BtOpGetNext, "構成マスタ")
'                        Exit Function
'                End Select
'
'
'            Loop
'
'            If Not Fsw Then
'
'                If LCnt > PCnt * 21 Then
'                    PCnt = PCnt + 1
'                    Row = Row + 33
'                    LCnt = 0
'
'                    Total_Line = Total_Line + 54
'                    excelSheet.Application.Range(excelSheet.Application.Rows(Start_Line), excelSheet.Application.Rows(End_Line)).Copy
'                    Start_Line = Start_Line + 54
'                    End_Line = End_Line + 54
'                    excelSheet.Application.Range(excelSheet.Application.Rows(Start_Line), excelSheet.Application.Rows(Start_Line)).pastespecial
'
'
'                    Clear_Start_Line = Clear_Start_Line + 54
'                    Clear_End_Line = Clear_End_Line + 54
'
'                    excelSheet.Application.Range(excelSheet.Application.Cells(Clear_Start_Line, 3), excelSheet.Application.Cells(Clear_End_Line, 11)).ClearContents
'
'                End If
'
'
'
'                Row = Row + 1
'                LCnt = LCnt + 1
'
'                Call UniCode_Conv(K0_ITEM.JGYOBU, Right(KOUSEI(i, ColKO_JGYOBU), 1))
'                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(i, ColKO_S_HIN_GAI))
'
'
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                Select Case sts
'                    Case BtNoErr
'                    Case BtErrKeyNotFound
'                        Call UniCode_Conv(ITEMREC.HIN_GAI, "未登録品番")
'                        Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")
'                    Case Else
'                        Call Input_UnLock
'                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
'                        Exit Function
'                End Select
'
'
'
'                excelSheet.Application.Cells(Row, xlHin_Name).Value = RTrim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
'                excelSheet.Application.Cells(Row, xlTANI).Value = ""
'
'                excelSheet.Application.Cells(Row, xlKO_QTY).Value = KOUSEI(i, ColKO_QTY)
'
'
'                If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
'                    excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value = CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode))
'                Else
'                    excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value = ""
'                End If
'
'                If IsNumeric(excelSheet.Application.Cells(Row, xlKO_QTY).Value) And _
'                    IsNumeric(excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value) Then
'
'                    excelSheet.Application.Cells(Row, xlG_KINGAKU).Value = ToHalfAdjust(CCur(CCur(excelSheet.Application.Cells(Row, xlKO_QTY).Value) * CCur(excelSheet.Application.Cells(Row, xlG_ST_URITAN).Value)), 2)
'                    Total_Kingaku = Total_Kingaku + excelSheet.Application.Cells(Row, xlG_KINGAKU).Value
'
'                End If
'
'
'            End If
'
'        End If
'
'
'    Next i
'
'
'
''    For i = ptxNAKANISHI_KIN To ptxKANRI_KIN               '2017.06.12
'    For i = ptxNAKANISHI_KIN To ptxKANRI_KIN Step 2         '2017.06.12
'
'        If IsNumeric(Text1(i).Text) Then
'            If CDbl(Text1(i).Text) <> 0 Then
'
'
'
'                Row = Row + 1
'
'
'
'
'                excelSheet.Application.Cells(Row, xlHin_Name).Value = lblTitle(i).Caption
'                excelSheet.Application.Cells(Row, xlTANI).Value = Text1(i - 1).Text
'                excelSheet.Application.Cells(Row, xlG_KINGAKU).Value = Text1(i).Text
'
'                Total_Kingaku = Total_Kingaku + excelSheet.Application.Cells(Row, xlG_KINGAKU).Value
'
'
'
'            End If
'        End If
'
'    Next i
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>印刷内容を全て画面より 2017.10.06

    'excelSheet.Application.Cells(Total_Line, xlG_KINGAKU - 1).Value = Total_Kingaku                    '2017.10.13


    excelSheet.PageSetup.printarea = excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(End_Line, 12)).address

    excelSheet.Application.Cells(1, 1).Select


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   終了処理
    excelApplication.Calculation = xlCalculationAutomatic
    
    
    excelApplication.ScreenUpdating = True
    excelApplication.Visible = True
    
    
    excelApplication.displayalerts = False
'    excelWorkBook.saveas FileName:=(Save_Dir & Trim(Text1(ptxHin_Gai).Text))       '2017.07.08
    excelWorkBook.saveas FileName:=(Save_Dir & Trim(Text1(ptxHin_Gai).Text)) & "-" & Format(Now, "YYYYMMDD hhmmss")      '2017.07.08
    
    
    
    
    Set excelSheet = Nothing
    
    Set excelWorkBook = Nothing
    

    
    Set excelApplication = Nothing

    
    
    Call Input_UnLock
    
    
    
    
    Estimate_Proc = False
End Function

Private Function Detail_Disp_Proc() As Integer
''----------------------------------------------------------------------------
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

Dim INV_F       As Boolean

Dim CATE_ST_SEC As Long


Dim NEW_ITEM    As Integer          '2017.09.28

    Detail_Disp_Proc = True
    
    
    If Error_Check_Proc(ptxTanto_Code) Then
        Detail_Disp_Proc = False
        Exit Function
    End If
    
    
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
            Detail_Disp_Proc = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function

    End Select
    
    
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2017.11.15
    If ITEM_O_RESET_PROC() Then
        Exit Function
    End If
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2017.11.15
    
    
    
    
    
    
    
    '>>>>>>>>>>大阪事 見積用品目マスタ  読込み  2017.09.28
    Call UniCode_Conv(K0_ITEM_O.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM_O.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM_O.HIN_GAI, Text1(ptxHin_Gai).Text)

    Call UniCode_Conv(K0_ITEM_O.SEQ_NO, "000")     '2017.11.08

    sts = BTRV(BtOpGetEqual, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
    Select Case sts
        Case BtNoErr
            NEW_ITEM = 0
        
        Case BtErrKeyNotFound

            NEW_ITEM = 1
        Case Else
            Call File_Error(sts, BtOpGetEqual, "大阪事 見積用品目マスタ")
            Exit Function

    End Select
    
    
'>>>>>>>>>>>>>  2017.09.27
    If NEW_ITEM = 1 Then
            
        Check1(0).Value = Print_FLG(0)
        Check1(1).Value = Print_FLG(1)
        Check1(2).Value = Print_FLG(2)
        Check1(3).Value = Print_FLG(3)
        Check1(4).Value = Print_FLG(4)
        Check1(5).Value = Print_FLG(5)
        Check1(6).Value = Print_FLG(6)
        Check1(7).Value = Print_FLG(7)
        Check1(8).Value = Print_FLG(8)
        Check1(9).Value = Print_FLG(9)
        Check1(10).Value = Print_FLG(10)
    
'>>>>>>>>>> 2017.10.20
    
        For i = ptxNAKANISHI_TANI To ptxKANRI_KIN
            Text1(i).Text = ""
        Next i
    
    
'>>>>>>>>>> 2017.10.20
    
    
    Else
        
        Check1(0).Value = Val(StrConv(ITEM_O_REC.NAKANISHI_F, vbUnicode))
        Check1(1).Value = Val(StrConv(ITEM_O_REC.SHOHIN_F, vbUnicode))
        Check1(2).Value = Val(StrConv(ITEM_O_REC.PF_KAKOU_F, vbUnicode))
        Check1(3).Value = Val(StrConv(ITEM_O_REC.PE_KAKOU_F, vbUnicode))
        Check1(4).Value = Val(StrConv(ITEM_O_REC.PE_SHIZAI_F, vbUnicode))
        Check1(5).Value = Val(StrConv(ITEM_O_REC.HINBAN_LABEL_F, vbUnicode))
        Check1(6).Value = Val(StrConv(ITEM_O_REC.KOUJI_SETSU_F, vbUnicode))
        Check1(7).Value = Val(StrConv(ITEM_O_REC.KONPOU_F, vbUnicode))
        Check1(8).Value = Val(StrConv(ITEM_O_REC.FUKU_SHIZAI_F, vbUnicode))
        Check1(9).Value = Val(StrConv(ITEM_O_REC.KONPOU_ASSY_F, vbUnicode))
        Check1(10).Value = Val(StrConv(ITEM_O_REC.KANRI_F, vbUnicode))
    
        
        Text1(ptxNAKANISHI_TANI).Text = StrConv(ITEM_O_REC.NAKANISHI_TANI, vbUnicode)       '中西工料　単位
                                                                                            '中西工料　金額
        Text1(ptxNAKANISHI_KIN).Text = Format(Val(StrConv(ITEM_O_REC.NAKANISHI_KIN, vbUnicode)), "#0")
            
        Text1(ptxSHOHIN_TANI).Text = StrConv(ITEM_O_REC.SHOHIN_TANI, vbUnicode)             '商品化工料　単位
                                                                                            '商品化工料　金額
        Text1(ptxSHOHIN_KIN).Text = Format(Val(StrConv(ITEM_O_REC.SHOHIN_KIN, vbUnicode)), "#0")
        
            
        Text1(ptxPF_KAKOU_TANI).Text = StrConv(ITEM_O_REC.PF_KAKOU_TANI, vbUnicode)         'PF加工　単位
                                                                                            'PF加工　金額
        Text1(ptxPF_KAKOU_KIN).Text = Format(Val(StrConv(ITEM_O_REC.PF_KAKOU_KIN, vbUnicode)), "#0")
            
        Text1(ptxPE_KAKOU_TANI).Text = StrConv(ITEM_O_REC.PE_KAKOU_TANI, vbUnicode)         'PE加工　単位
                                                                                            'PE加工　金額
        Text1(ptxPE_KAKOU_KIN).Text = Format(Val(StrConv(ITEM_O_REC.PE_KAKOU_KIN, vbUnicode)), "#0")
            
        Text1(ptxPE_SHIZAI_TANI).Text = StrConv(ITEM_O_REC.PE_SHIZAI_TANI, vbUnicode)       'PE資材　単位
                                                                                            'PE資材　金額
        Text1(ptxPE_SHIZAI_KIN).Text = Format(Val(StrConv(ITEM_O_REC.PE_SHIZAI_KIN, vbUnicode)), "#0")
    
        Text1(ptxHINBAN_LABEL_TANI).Text = StrConv(ITEM_O_REC.HINBAN_LABEL_TANI, vbUnicode) '品番表示ﾗﾍﾞﾙ　単位
                                                                                            '品番表示ﾗﾍﾞﾙ　金額
        Text1(ptxHINBAN_LABEL_KIN).Text = Format(Val(StrConv(ITEM_O_REC.HINBAN_LABEL_KIN, vbUnicode)), "#0")
    
        Text1(ptxKOUJI_SETSU_TANI).Text = StrConv(ITEM_O_REC.KOUJI_SETSU_TANI, vbUnicode)   '設置工事説明書　単位
                                                                                            '設置工事説明書　金額
        Text1(ptxKOUJI_SETSU_KIN).Text = Format(Val(StrConv(ITEM_O_REC.KOUJI_SETSU_KIN, vbUnicode)), "#0")
    
        Text1(ptxKONPOU_TANI).Text = StrConv(ITEM_O_REC.KONPOU_TANI, vbUnicode)             '梱包材　単位
                                                                                            '梱包材　金額
        Text1(ptxKONPOU_KIN).Text = Format(Val(StrConv(ITEM_O_REC.KONPOU_KIN, vbUnicode)), "#0")
    
        Text1(ptxFUKU_SHIZAI_TANI).Text = StrConv(ITEM_O_REC.FUKU_SHIZAI_TANI, vbUnicode)   '副資材　単位
                                                                                            '副資材　金額
        Text1(ptxFUKU_SHIZAI_KIN).Text = Format(Val(StrConv(ITEM_O_REC.FUKU_SHIZAI_KIN, vbUnicode)), "#0")
    
        Text1(ptxKONPOU_ASSY_TANI).Text = StrConv(ITEM_O_REC.KONPOU_ASSY_TANI, vbUnicode)   '梱包ASSY　単位
                                                                                            '梱包ASSY　金額
        Text1(ptxKONPOU_ASSY_KIN).Text = Format(Val(StrConv(ITEM_O_REC.KONPOU_ASSY_KIN, vbUnicode)), "#0")
    
        Text1(ptxKANRI_TANI).Text = StrConv(ITEM_O_REC.KANRI_TANI, vbUnicode)               '管理費　単位
                                                                                            '管理費　金額
        Text1(ptxKANRI_KIN).Text = Format(Val(StrConv(ITEM_O_REC.KANRI_KIN, vbUnicode)), "#0")
    
    
    
                                                                                            '中西工料　提出金額
        Text1(ptxNAKANISHI_T_KIN).Text = Format(Val(StrConv(ITEM_O_REC.NAKANISHI_T_KIN, vbUnicode)), "#0")
                                                                                            '商品化工料　提出金額
        Text1(ptxSHOHIN_T_KIN).Text = Format(Val(StrConv(ITEM_O_REC.SHOHIN_T_KIN, vbUnicode)), "#0")
                                                                                            'PF加工　提出金額
        Text1(ptxPF_KAKOU_T_KIN).Text = Format(Val(StrConv(ITEM_O_REC.PF_KAKOU_T_KIN, vbUnicode)), "#0")
                                                                                            'PE加工　提出金額
        Text1(ptxPE_KAKOU_T_KIN).Text = Format(Val(StrConv(ITEM_O_REC.PE_KAKOU_T_KIN, vbUnicode)), "#0")
                                                                                            'PE資材　提出金額
        Text1(ptxPE_SHIZAI_T_KIN).Text = Format(Val(StrConv(ITEM_O_REC.PE_SHIZAI_T_KIN, vbUnicode)), "#0")
                                                                                            '品番表示ﾗﾍﾞﾙ　提出金額
        Text1(ptxHINBAN_LABEL_T_KIN).Text = Format(Val(StrConv(ITEM_O_REC.HINBAN_LABEL_T_KIN, vbUnicode)), "#0")
                                                                                            '設置工事説明書　提出金額
        Text1(ptxKOUJI_SETSU_T_KIN).Text = Format(Val(StrConv(ITEM_O_REC.KOUJI_SETSU_T_KIN, vbUnicode)), "#0")
                                                                                            '梱包材　提出金額
        Text1(ptxKONPOU_T_KIN).Text = Format(Val(StrConv(ITEM_O_REC.KONPOU_T_KIN, vbUnicode)), "#0")
                                                                                            '副資材　提出金額
        Text1(ptxFUKU_SHIZAI_T_KIN).Text = Format(Val(StrConv(ITEM_O_REC.FUKU_SHIZAI_T_KIN, vbUnicode)), "#0")
                                                                                            '梱包ASSY　提出金額
        Text1(ptxKONPOU_ASSY_T_KIN).Text = Format(Val(StrConv(ITEM_O_REC.KONPOU_ASSY_T_KIN, vbUnicode)), "#0")
                                                                                            '管理費　提出金額
        Text1(ptxKANRI_T_KIN).Text = Format(Val(StrConv(ITEM_O_REC.KANRI_T_KIN, vbUnicode)), "#0")
    
'>>>>>>>>>>>>>  金額--＞提出金額    2017.07.08

'>>>>>>>>>>>>>  数量/提出単価        2017.11.08
                                                                                            '中西工料　数量
        Text1(ptxNAKANISHI_QTY).Text = Format(Val(StrConv(ITEM_O_REC.NAKANISHI_QTY, vbUnicode)), "#0.00")
                                                                                            '商品化工料　数量
        Text1(ptxSHOHIN_QTY).Text = Format(Val(StrConv(ITEM_O_REC.SHOHIN_QTY, vbUnicode)), "#0.00")
                                                                                            'PF加工　数量
        Text1(ptxPF_KAKOU_QTY).Text = Format(Val(StrConv(ITEM_O_REC.PF_KAKOU_QTY, vbUnicode)), "#0.00")
                                                                                            'PE加工　数量
        Text1(ptxPE_KAKOU_QTY).Text = Format(Val(StrConv(ITEM_O_REC.PE_KAKOU_QTY, vbUnicode)), "#0.00")
                                                                                            'PE資材　数量
        Text1(ptxPE_SHIZAI_QTY).Text = Format(Val(StrConv(ITEM_O_REC.PE_SHIZAI_QTY, vbUnicode)), "#0.00")
                                                                                            '品番表示ﾗﾍﾞﾙ　数量
        Text1(ptxHINBAN_LABEL_QTY).Text = Format(Val(StrConv(ITEM_O_REC.HINBAN_LABEL_QTY, vbUnicode)), "#0.00")
                                                                                            '設置工事説明書　数量
        Text1(ptxKOUJI_SETSU_QTY).Text = Format(Val(StrConv(ITEM_O_REC.KOUJI_SETSU_QTY, vbUnicode)), "#0.00")
                                                                                            '梱包材　数量
        Text1(ptxKONPOU_QTY).Text = Format(Val(StrConv(ITEM_O_REC.KONPOU_QTY, vbUnicode)), "#0.00")
                                                                                            '副資材　数量
        Text1(ptxFUKU_SHIZAI_QTY).Text = Format(Val(StrConv(ITEM_O_REC.FUKU_SHIZAI_QTY, vbUnicode)), "#0.00")
                                                                                            '梱包ASSY　数量
        Text1(ptxKONPOU_ASSY_QTY).Text = Format(Val(StrConv(ITEM_O_REC.KONPOU_ASSY_QTY, vbUnicode)), "#0.00")
                                                                                            '管理費　数量
        Text1(ptxKANRI_QTY).Text = Format(Val(StrConv(ITEM_O_REC.KANRI_QTY, vbUnicode)), "#0.00")


                                                                                            '中西工料　提出単価
        Text1(ptxNAKANISHI_T_TAN).Text = Format(Val(StrConv(ITEM_O_REC.NAKANISHI_T_TAN, vbUnicode)), "#0.00")
                                                                                            '商品化工料　提出単価
        Text1(ptxSHOHIN_T_TAN).Text = Format(Val(StrConv(ITEM_O_REC.SHOHIN_T_TAN, vbUnicode)), "#0.00")
                                                                                            'PF加工　提出単価
        Text1(ptxPF_KAKOU_T_TAN).Text = Format(Val(StrConv(ITEM_O_REC.PF_KAKOU_T_TAN, vbUnicode)), "#0.00")
                                                                                            'PE加工　提出単価
        Text1(ptxPE_KAKOU_T_TAN).Text = Format(Val(StrConv(ITEM_O_REC.PE_KAKOU_T_TAN, vbUnicode)), "#0.00")
                                                                                            'PE資材　提出単価
        Text1(ptxPE_SHIZAI_T_TAN).Text = Format(Val(StrConv(ITEM_O_REC.PE_SHIZAI_T_TAN, vbUnicode)), "#0.00")
                                                                                            '品番表示ﾗﾍﾞﾙ　提出単価
        Text1(ptxHINBAN_LABEL_T_TAN).Text = Format(Val(StrConv(ITEM_O_REC.HINBAN_LABEL_T_TAN, vbUnicode)), "#0.00")
                                                                                            '設置工事説明書　提出単価
        Text1(ptxKOUJI_SETSU_T_TAN).Text = Format(Val(StrConv(ITEM_O_REC.KOUJI_SETSU_T_TAN, vbUnicode)), "#0.00")
                                                                                            '梱包材　提出単価
        Text1(ptxKONPOU_T_TAN).Text = Format(Val(StrConv(ITEM_O_REC.KONPOU_T_TAN, vbUnicode)), "#0.00")
                                                                                            '副資材　提出単価
        Text1(ptxFUKU_SHIZAI_T_TAN).Text = Format(Val(StrConv(ITEM_O_REC.FUKU_SHIZAI_T_TAN, vbUnicode)), "#0.00")
                                                                                            '梱包ASSY　提出単価
        Text1(ptxKONPOU_ASSY_T_TAN).Text = Format(Val(StrConv(ITEM_O_REC.KONPOU_ASSY_T_TAN, vbUnicode)), "#0.00")
                                                                                            '管理費　提出単価
        Text1(ptxKANRI_T_TAN).Text = Format(Val(StrConv(ITEM_O_REC.KANRI_T_TAN, vbUnicode)), "#0.00")

'>>>>>>>>>>>>>  数量/提出単価        2017.11.08
    
    End If
    
    
    
    For i = 2 To 5      '2013.01.16 5-->6
        Command1(i).Enabled = True
    Next i
    
    
    '品名
    Text1(ptxHin_Name).Text = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
        
    '標準棚番
    Text1(ptxST_SOKO).Text = StrConv(ITEMREC.ST_SOKO, vbUnicode)
    Text1(ptxST_RETU).Text = StrConv(ITEMREC.ST_RETU, vbUnicode)
    Text1(ptxST_REN).Text = StrConv(ITEMREC.ST_REN, vbUnicode)
    Text1(ptxST_DAN).Text = StrConv(ITEMREC.ST_DAN, vbUnicode)
    
    
    
    '国内供給区分   2017.07.08
'    Text1(ptxNAI_BUHIN).Text = Trim(StrConv(ITEMREC.NAI_BUHIN, vbUnicode))
    
    
    
    '品名カテゴリィ
    Text1(ptxCATEGORY_CODE).Text = Trim(StrConv(ITEMREC.CATEGORY_CODE, vbUnicode))
    For i = 0 To Combo1(pcmbCATEGORY_Name).ListCount - 1
        If Trim(Text1(ptxCATEGORY_CODE).Text) = Trim(Right(Combo1(pcmbCATEGORY_Name).List(i), 8)) Then
            Text1(ptxCATEGORY_CODE).Text = Trim(Right(Combo1(pcmbCATEGORY_Name).List(i), 8))
            Combo1(pcmbCATEGORY_Name).ListIndex = i
            Exit For
        End If
    Next i
    If i > Combo1(pcmbCATEGORY_Name).ListCount - 1 Then
        Combo1(pcmbCATEGORY_Name).ListIndex = 0
    End If
    '見積書備考
    wkBikou = Replace(StrConv(ITEMREC.M_BIKOU, vbUnicode), Chr(0), " ")
    RichTextBox1(prchM_BIKOU).Text = RTrim(wkBikou)
    
    '単価切替日
    Text1(ptxTANKA_KIRIKAE_DT).Text = RTrim(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode))

'>>>>>>>>>>>>>  2017.09.27
    '打切り
    If StrConv(ITEMREC.NAI_BUHIN, vbUnicode) = "3" Then
        Text1(ptxNAI_BUHIN) = "3"
    Else
        Text1(ptxNAI_BUHIN) = ""
    End If
    '前回提出金額合計／合計金額
    If NEW_ITEM = 0 Then
    
        lblZEN_T_GOUKEI_KIN = Format(Val(StrConv(ITEM_O_REC.GOUKEI_T_KIN, vbUnicode)), "#0")
        lblZEN_GOUKEI_KIN = Format(Val(StrConv(ITEM_O_REC.GOUKEI_KIN, vbUnicode)), "#0")
    Else                            '2017.10.20
        lblZEN_T_GOUKEI_KIN = "0"   '2017.10.20
        lblZEN_GOUKEI_KIN = "0"     '2017.10.20
    End If
'>>>>>>>>>>>>>  2017.09.27
        
        
    
    
    
    
    
    
    
    
    '-----------------------------------    構成品表示
    If P_COMPO_Disp_Proc() Then
        Exit Function
    End If
    

    
    
    

    If TANKA_KEISAN_Proc() Then
        Unload Me
    End If

    DoEvents

    If CLng(lblZEN_T_GOUKEI_KIN.Caption) <> CLng(lblGOUKEI_T_KIN.Caption) Then      '2107.11.13
        MsgBox "前回合計金額と合計金額が異なっています。明細内容を確認して下さい。" '2107.11.13
    End If                                                                          '2107.11.13
    



    Detail_Disp_Proc = False

End Function

Private Function TANKA_KEISAN_Proc() As Integer

Dim i               As Integer
Dim j               As Integer

Dim wkKingaku       As Double
Dim wk_T_Kingaku    As Double


    TANKA_KEISAN_Proc = True
    
'    Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
'    TDBGrid1(pGrdKOUSEI).ReBind
'    TDBGrid1(pGrdKOUSEI).Update
    
    wkKingaku = 0
    wk_T_Kingaku = 0
    
    For i = 1 To KOUSEI.UpperBound(1)
    
    Debug.Print Trim(KOUSEI(i, ColNO)) & Trim(KOUSEI(i, ColKO_T_HIN_NAME))
    
        If Trim(KOUSEI(i, ColKO_S_HIN_GAI)) = "" And Trim(KOUSEI(i, ColKO_T_HIN_NAME)) = "" Then
            Exit For
        Else
            
            
            If IsNumeric(KOUSEI(i, ColG_KINGAKU)) Then
                wkKingaku = wkKingaku + Val(KOUSEI(i, ColG_KINGAKU))
            End If
            For j = 0 To UBound(Print_SYUBETSU)
            
                If Right(KOUSEI(i, ColKO_SYUBETSU), 2) = Print_SYUBETSU(j) Then
                    If IsNumeric(KOUSEI(i, ColKO_T_KINGAKU)) Then
                        wk_T_Kingaku = wk_T_Kingaku + Val(KOUSEI(i, ColKO_T_KINGAKU))
                        Exit For
                    End If
                End If
            Next j
        End If
    
    
    Next i
    
    
    For i = ptxNAKANISHI_KIN To ptxKANRI_KIN Step 2
        If IsNumeric(Text1(i).Text) Then
            wkKingaku = wkKingaku + Val(Text1(i).Text)
        End If
    Next i
    
    
    j = 8
    For i = chkNAKANISHI_F To chkKANRI_F
        j = j + 5
        If Check1(i).Value = 1 Then
            
            wk_T_Kingaku = wk_T_Kingaku + Val(Text1(j).Text)
        
        End If
    
    Next i
    
    
    lblGOUKEI_KIN.Caption = Format(wkKingaku, "#0")
    lblGOUKEI_T_KIN.Caption = Format(wk_T_Kingaku, "#0")
        
    
'    TDBGrid1(pGrdKOUSEI).ScrollBars = dbgAutomatic
    
    TANKA_KEISAN_Proc = False

End Function





Private Function Tanka_Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   単価登録処理
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim com     As Integer
    
Dim ans     As Integer
    
    Tanka_Update_Proc = True
    
    Text1(ptxHin_Gai).Text = Trim(StrConv(Text1(ptxHin_Gai).Text, vbUpperCase))
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)


    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                        
                Exit Do
            
            Case BtErrKeyNotFound
    
                Text1(ptxHin_Name).Text = ""
    
                MsgBox "入力した項目はエラーです。(品番)"
                Unload Me
                Exit Do
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Do
                End If
            
            
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                Exit Function
    
        End Select
    Loop
        
    Call UniCode_Conv(K0_ITEM_O.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM_O.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM_O.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM_O.SEQ_NO, "000")
        
    Do
    
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
            
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
            
                Exit Do
                            
            Case BtErrKeyNotFound
                
                com = BtOpInsert
            
                Call Rclr_ITEM_O_REC
            
                Exit Do
            
            
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "大阪事　見積用品目マスタ")
                Exit Function
        End Select

    Loop
    
    Call UniCode_Conv(ITEM_O_REC.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(ITEM_O_REC.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(ITEM_O_REC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
    
    
    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_TANI, Text1(ptxNAKANISHI_TANI).Text) '中西工料　単位
    If IsNumeric(Text1(ptxNAKANISHI_KIN).Text) Then                                '中西工料　金額
        Call UniCode_Conv(ITEM_O_REC.NAKANISHI_KIN, Format(CDbl(Text1(ptxNAKANISHI_KIN).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEM_O_REC.NAKANISHI_KIN, "00000000.00")
    End If
    
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_TANI, Text1(ptxSHOHIN_TANI).Text)       '中西工料　単位
    If IsNumeric(Text1(ptxSHOHIN_KIN).Text) Then                                '中西工料　金額
        Call UniCode_Conv(ITEM_O_REC.SHOHIN_KIN, Format(CDbl(Text1(ptxSHOHIN_KIN).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEM_O_REC.SHOHIN_KIN, "00000000.00")
    End If
    
    
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_TANI, Text1(ptxPF_KAKOU_TANI).Text)   'PF加工　単位
    If IsNumeric(Text1(ptxPF_KAKOU_KIN).Text) Then                              'PF加工　金額
        Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_KIN, Format(CDbl(Text1(ptxPF_KAKOU_KIN).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_KIN, "00000000.00")
    End If
    
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_TANI, Text1(ptxPE_KAKOU_TANI).Text)   'PE加工　単位
    If IsNumeric(Text1(ptxPE_KAKOU_KIN).Text) Then                              'PE加工　金額
        Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_KIN, Format(CDbl(Text1(ptxPE_KAKOU_KIN).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_KIN, "00000000.00")
    End If
    
    
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_TANI, Text1(ptxPE_SHIZAI_TANI).Text) 'PE資材　単位
    If IsNumeric(Text1(ptxPE_SHIZAI_KIN).Text) Then                             'PE資材　金額
        Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_KIN, Format(CDbl(Text1(ptxPE_SHIZAI_KIN).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_KIN, "00000000.00")
    End If
    
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_TANI, Text1(ptxHINBAN_LABEL_TANI).Text)   '品番表示ﾗﾍﾞﾙ　単位
    If IsNumeric(Text1(ptxHINBAN_LABEL_KIN).Text) Then                                  '品番表示ﾗﾍﾞﾙ　金額
        Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_KIN, Format(CDbl(Text1(ptxHINBAN_LABEL_KIN).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_KIN, "00000000.00")
    End If
    
    
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_TANI, Text1(ptxKOUJI_SETSU_TANI).Text)     '設置工事説明書　単位
    If IsNumeric(Text1(ptxKOUJI_SETSU_KIN).Text) Then                                   '設置工事説明書　金額
        Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_KIN, Format(CDbl(Text1(ptxKOUJI_SETSU_KIN).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_KIN, "00000000.00")
    End If
    
    Call UniCode_Conv(ITEM_O_REC.KONPOU_TANI, Text1(ptxKONPOU_TANI).Text)       '梱包材　単位
    If IsNumeric(Text1(ptxKONPOU_KIN).Text) Then                                '梱包材　金額
        Call UniCode_Conv(ITEM_O_REC.KONPOU_KIN, Format(CDbl(Text1(ptxKONPOU_KIN).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEM_O_REC.KONPOU_KIN, "00000000.00")
    End If
    
    
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_TANI, Text1(ptxFUKU_SHIZAI_TANI).Text) '副資材　単位
    If IsNumeric(Text1(ptxFUKU_SHIZAI_KIN).Text) Then                               '副資材　金額
        Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_KIN, Format(CDbl(Text1(ptxFUKU_SHIZAI_KIN).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_KIN, "00000000.00")
    End If

    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_TANI, Text1(ptxKONPOU_ASSY_TANI).Text) '梱包ASSY　単位
    If IsNumeric(Text1(ptxKONPOU_ASSY_KIN).Text) Then                               '梱包ASSY　金額
        Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_KIN, Format(CDbl(Text1(ptxKONPOU_ASSY_KIN).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_KIN, "00000000.00")
    End If


    Call UniCode_Conv(ITEM_O_REC.KANRI_TANI, Text1(ptxKANRI_TANI).Text)     '梱包ASSY　単位
    If IsNumeric(Text1(ptxKANRI_KIN).Text) Then                             '梱包ASSY　金額
        Call UniCode_Conv(ITEM_O_REC.KANRI_KIN, Format(CDbl(Text1(ptxKANRI_KIN).Text), "00000000.00"))
    Else
        Call UniCode_Conv(ITEM_O_REC.KANRI_KIN, "00000000.00")
    End If


    If IsNumeric(lblGOUKEI_KIN.Caption) Then                                '合計金額
        Call UniCode_Conv(ITEM_O_REC.GOUKEI_KIN, Format(CDbl(lblGOUKEI_KIN.Caption), "00000000.00"))
    Else
        Call UniCode_Conv(ITEM_O_REC.GOUKEI_KIN, "00000000.00")
    End If
                                                                            '担当者ｺｰﾄﾞ
    Call UniCode_Conv(ITEM_O_REC.INPUT_TANTO_CODE, Text1(ptxTanto_Code).Text)


    If com = BtOpInsert Then
        Call UniCode_Conv(ITEM_O_REC.INS_TANTO, "SEI16")
        Call UniCode_Conv(ITEM_O_REC.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))
    Else
        Call UniCode_Conv(ITEM_O_REC.UPD_TANTO, "SEI16")
        Call UniCode_Conv(ITEM_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
    End If


    Do
    
        sts = BTRV(com, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
            
        Select Case sts
            Case BtNoErr
            
                Exit Do
                            
            
            
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, com, "大阪事　見積用品目マスタ")
                Exit Function
        End Select

    Loop


    
    '見積書備考
    Call UniCode_Conv(ITEMREC.M_BIKOU, RichTextBox1(prchM_BIKOU).Text)


    Do
        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            
                Exit Do
            
            Case BtErrKeyNotFound
    
                Text1(ptxHin_Name).Text = ""
    
                MsgBox "入力した項目はエラーです。(品番)"
                Text1(ptxHin_Gai).SetFocus
                Exit Do
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Do
                End If
            
            
            Case Else
                Call File_Error(sts, BtOpUpdate, "品目マスタ")
                Exit Function
    
        End Select
    Loop



    Tanka_Update_Proc = False


End Function

Private Function Grid_Error_Check_Proc() As Integer
'----------------------------------------------------------------------------
'                   ｸﾞﾘｯﾄﾞ内容のエラーチェック処理
'----------------------------------------------------------------------------
Dim i               As Long

Dim sts             As Integer
    
Dim j               As Long
    
    
Dim NEW_ITEM        As Integer          '2017.09.28
    
    
    
    
    Grid_Error_Check_Proc = True
    
    
    
    Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
    
    
   TDBGrid1(pGrdKOUSEI).Update
    
    
    For i = 1 To KOUSEI.UpperBound(1)
        ' 指図票品番の削除
        If Trim(KOUSEI(i, ColKO_S_HIN_GAI)) = "" And Trim(KOUSEI(i, ColKO_T_HIN_NAME)) = "" Then
            
            KOUSEI(i, ColKO_SYUBETSU) = ""
            KOUSEI(i, ColKO_JGYOBU) = ""
            KOUSEI(i, ColKO_S_HIN_GAI) = ""
            KOUSEI(i, ColKO_HIN_NAME) = ""

            KOUSEI(i, ColKO_T_HIN_NAME) = ""        '2017.09.28

            KOUSEI(i, ColKO_QTY) = ""

            KOUSEI(i, ColKO_TANI) = ""              '2017.09.28
            KOUSEI(i, ColKO_T_TANKA) = ""           '2017.09.28
            KOUSEI(i, ColKO_T_KINGAKU) = ""         '2017.09.28


            KOUSEI(i, ColG_ST_SHITAN) = ""
            KOUSEI(i, ColG_ST_URITAN) = ""

            KOUSEI(i, ColG_KINGAKU) = ""
'
'            KOUSEI(i, ColKO_UMU) = ""  2017.09.28

        Else
            '品番
            Call UniCode_Conv(K0_ITEM.JGYOBU, Right(KOUSEI(i, ColKO_JGYOBU), 1))
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(i, ColKO_S_HIN_GAI))
    
    
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    KOUSEI(i, ColKO_HIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                    If KOUSEI(i, ColG_ST_SHITAN) = "" Then
                    
                        If IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
                            KOUSEI(i, ColG_ST_SHITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)), "#0.00")
                        Else
                            KOUSEI(i, ColG_ST_SHITAN) = Format(CDbl(0), "#0.00")
                        End If
                    End If
                
                    If KOUSEI(i, ColG_ST_URITAN) = "" Then
                    
                        If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                            KOUSEI(i, ColG_ST_URITAN) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
                        Else
                            KOUSEI(i, ColG_ST_URITAN) = Format(CDbl(0), "#0.00")
                        End If
                    End If
                
                
                    '>>>>>>>>>>>>>>>    大阪事　見積用品目マスタ    2017.09.28
                    'Call UniCode_Conv(K0_ITEM_O.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                    'Call UniCode_Conv(K0_ITEM_O.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                    'Call UniCode_Conv(K0_ITEM_O.HIN_GAI, Text1(ptxHin_Gai).Text)
                    '
                    'Call UniCode_Conv(K0_ITEM_O.SEQ_NO, KOUSEI(i, ColNO))
                    '
                    '
                    'sts = BTRV(BtOpGetEqual, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
                    'Select Case sts
                    '    Case BtNoErr
                    '        NEW_ITEM = 0
                    '    Case BtErrKeyNotFound
                    '        NEW_ITEM = 1
                    '    Case Else
                    '        Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                    '        Exit Function
                    'End Select
                    '>>>>>>>>>>>>>>>    大阪事　見積用品目マスタ    2017.09.28
                
                
                Case BtErrKeyNotFound
                    
                        If Trim(KOUSEI(i, ColKO_S_HIN_GAI)) = "" Then           '2017.11.14
                        Else
                        
                            If HIN_INV Then
                                '未登録品番　可　資材としておく
                                KOUSEI(i, ColKO_HIN_NAME) = "未登録品番"
                            Else
                                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(品番)"
                                Exit Function
                            End If
                        End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                    Exit Function
            End Select
                
                
                
            '>>>>>>>>>>>>>>>    大阪事　見積用品目マスタ    2017.09.28
            Call UniCode_Conv(K0_ITEM_O.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
            Call UniCode_Conv(K0_ITEM_O.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
            Call UniCode_Conv(K0_ITEM_O.HIN_GAI, Text1(ptxHin_Gai).Text)
                
            Call UniCode_Conv(K0_ITEM_O.SEQ_NO, KOUSEI(i, ColNO))
                
                
            sts = BTRV(BtOpGetEqual, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
            Select Case sts
                Case BtNoErr
                    NEW_ITEM = 0
                Case BtErrKeyNotFound
                    NEW_ITEM = 1
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                    Exit Function
            End Select
            '>>>>>>>>>>>>>>>    大阪事　見積用品目マスタ    2017.09.28
                
                            
                            
                
                
                
                
                '員数
            If IsNumeric(KOUSEI(i, ColKO_QTY)) Then
                KOUSEI(i, ColKO_QTY) = Format(CDbl(KOUSEI(i, ColKO_QTY)), "#0.00")
            Else
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(員数)"
                Exit Function
            End If
                
                
                
                '提出単価   ’2017.09.28
            If KOUSEI(i, ColKO_T_TANKA) = "" Then
                KOUSEI(i, ColKO_T_TANKA) = "0"
            End If
            If IsNumeric(KOUSEI(i, ColKO_T_TANKA)) Then
                
                KOUSEI(i, ColKO_T_TANKA) = Format(CDbl(KOUSEI(i, ColKO_T_TANKA)), "#0.00")
            
                KOUSEI(i, ColKO_T_KINGAKU) = ToHalfAdjust(CCur(CCur(KOUSEI(i, ColKO_QTY)) * CCur(KOUSEI(i, ColKO_T_TANKA))), 2)
            
                                
            
            Else
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(提出単価)"
                Exit Function
            End If
                
                
                '仕入＠
            If Trim(KOUSEI(i, ColG_ST_SHITAN)) = "" Then
                KOUSEI(i, ColG_ST_SHITAN) = "0.00"
            End If
            If IsNumeric(KOUSEI(i, ColG_ST_SHITAN)) Then
                KOUSEI(i, ColG_ST_SHITAN) = Format(CDbl(KOUSEI(i, ColG_ST_SHITAN)), "#0.00")
            Else
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(仕入単価)"
                Exit Function
            End If
                '販売＠
            If Trim(KOUSEI(i, ColG_ST_URITAN)) = "" Then
                KOUSEI(i, ColG_ST_URITAN) = "0.00"
            End If
            
            If IsNumeric(KOUSEI(i, ColG_ST_URITAN)) Then
                KOUSEI(i, ColG_ST_URITAN) = Format(CDbl(KOUSEI(i, ColG_ST_URITAN)), "#0.00")
            Else
                MsgBox "[" & Format(i, "0") & "]行目 入力した項目はエラーです。(販売単価)"
                Exit Function
            End If
    
            '合計金額
            KOUSEI(i, ColG_KINGAKU) = ToHalfAdjust(CCur(CCur(KOUSEI(i, ColKO_QTY)) * CCur(KOUSEI(i, ColG_ST_URITAN))), 2)
    
    
        End If
    
    
    
    
    
    Next i
    
    Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
    

'    TDBGrid1(pGrdKOUSEI).Refresh
    TDBGrid1(pGrdKOUSEI).Update

    TDBGrid1(pGrdKOUSEI).SetFocus
    
    
    

    Grid_Error_Check_Proc = False



End Function
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   構成マスタ出力
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim ans         As Integer


Dim i           As Integer
Dim j           As Integer

Dim MESG        As String

Dim D_SEQNO     As Integer


Dim NEW_ITEM    As Integer          '2017.09.28
Dim upd_com     As Integer          '2017.09.28


    Update_Proc = True
                                        
    Call Input_Lock
                                        
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
    
    
    
    
    
    
    
    
    '品目マスタ　（親）
                                                                                '事業部
    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
                                                                                '国内外
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)
    
    Do
    
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
            
                ans = MsgBox("データ内容が変更されています。確認して下さい。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "構成マスタ")
                GoTo Abort_Tran
        End Select

    Loop
    'カテゴリ
    Call UniCode_Conv(ITEMREC.CATEGORY_CODE, Text1(ptxCATEGORY_CODE).Text)
    '見積書備考
    Call UniCode_Conv(ITEMREC.M_BIKOU, RichTextBox1(prchM_BIKOU).Text)
    '国内供給区分           ’2017.09.28
    Call UniCode_Conv(ITEMREC.NAI_BUHIN, Text1(ptxNAI_BUHIN).Text)

    Call UniCode_Conv(ITEM_O_REC.UPD_TANTO, "SEI16")
    Call UniCode_Conv(ITEM_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))

    Do
        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            
                Exit Do
            
            Case BtErrKeyNotFound
    
                Text1(ptxHin_Name).Text = ""
    
                MsgBox "入力した項目はエラーです。(品番)"
                Text1(ptxHin_Gai).SetFocus
                Exit Do
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Do
                End If
            
            
            Case Else
                Call File_Error(sts, BtOpUpdate, "品目マスタ")
                Exit Function
    
        End Select
    Loop


    
    
        
    
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

    Call UniCode_Conv(P_COMPO_O_REC.UPD_TANTO, "SEI16")            '更新担当者ｺｰﾄﾞ
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
    TDBGrid1(pGrdKOUSEI).ReBind
    TDBGrid1(pGrdKOUSEI).Update


    D_SEQNO = 0


    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then

    Else


        For i = 1 To KOUSEI.UpperBound(1)
    
    
            If Trim(KOUSEI(i, ColKO_S_HIN_GAI)) = "" Then
            Else
                                                                                            '仕向け先ｺｰﾄﾞ
                Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
                                                                                            '事業部
                Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
                                                                                            '国内外
                Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
                Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, Text1(ptxHin_Gai).Text)
            
            
            
                
        
                Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_DOUKON)             'データ区分
                        
                D_SEQNO = D_SEQNO + 10
                        
                Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(D_SEQNO, "000"))  '追番
                                                                                '種別
                Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, Right(KOUSEI(i, ColKO_SYUBETSU), 2))
            
                                                                                            '子　事業部
                Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, Right(KOUSEI(i, ColKO_JGYOBU), 1))
                Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, NAIGAI_NAI)                      '子　国内外
                Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, KOUSEI(i, ColKO_S_HIN_GAI))     '子　品番
                                                                                            '員数
                Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(CDbl(KOUSEI(i, ColKO_QTY)), "000.00"))
            
                Call UniCode_Conv(P_COMPO_K_REC.FILLER, "")
            
                Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, "SEI16")                         '更新担当者ｺｰﾄﾞ
                                                                                            '更新日時
                Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
            
'>>>>>>>>>  2017.11.29
'                Call UniCode_Conv(K0_ITEM.JGYOBU, Right(KOUSEI(i, ColKO_JGYOBU), 1)) '事業部
'                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)                        '国内外
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(i, ColKO_S_HIN_GAI))       '品番
'
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'
'                Select Case sts
'                    Case BtNoErr
'                    Case BtErrKeyNotFound
'                        Call UniCode_Conv(ITEMREC.HIN_NAME, "")
'                    Case Else
'                        Call File_Error(sts, com + BtSNoWait, "品目マスタ")
'                        GoTo Abort_Tran
'               End Select
    
                Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, KOUSEI(i, ColKO_HIN_NAME))
    
'>>>>>>>>>  2017.11.29
            
            
            
            
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
    
    
                If Trim(KOUSEI(i, ColKO_S_HIN_GAI)) <> "" Then      '2017.11.15
                    '>>>>>>>>>>>>>  品目単価　更新
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Right(KOUSEI(i, ColKO_JGYOBU), 1)) '事業部
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)                        '国内外
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, KOUSEI(i, ColKO_S_HIN_GAI))       '品番
        
                    Do
                        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrKeyNotFound
                                Beep
                                ans = MsgBox("他端末でデータが変更されました。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                GoTo Abort_Tran
                    
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    GoTo Abort_Tran
                                End If
                            Case Else
                                Call File_Error(sts, com + BtSNoWait, "品目マスタ")
                                GoTo Abort_Tran
                       End Select
        
                    Loop
        
                    If Not IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
                        Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "00000000.00")
                    End If
                    If Not IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                        Call UniCode_Conv(ITEMREC.G_ST_URITAN, "00000000.00")
                    End If
        
        
        
                    If CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) <> CDbl(KOUSEI(i, ColG_ST_SHITAN)) Then
                        Call UniCode_Conv(ITEMREC.G_ST_SHITAN, Format(CDbl(KOUSEI(i, ColG_ST_SHITAN)), "00000000.00"))
                        Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, Format(Now, "YYYY/MM/DD"))
                                        
                        Call UniCode_Conv(ITEMREC.UPD_TANTO, "SEI16")
                        Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, Format(Now, "YYYYMMDDHHMMSS"))
        
                    End If
        
        
                    If CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) <> CDbl(KOUSEI(i, ColG_ST_URITAN)) Then
                        Call UniCode_Conv(ITEMREC.G_ST_URITAN, Format(CDbl(KOUSEI(i, ColG_ST_URITAN)), "00000000.00"))
                        Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, Format(Now, "YYYY/MM/DD"))
                                        
                        Call UniCode_Conv(ITEMREC.UPD_TANTO, "SEI16")
                        Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, Format(Now, "YYYYMMDDHHMMSS"))
        
                    End If
        
                    
                    '見積書備考
    '                Call UniCode_Conv(ITEMREC.M_BIKOU, RichTextBox1(prchM_BIKOU).Text)
                    
                    
                    
                    Do
                        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                    
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    GoTo Abort_Tran
                                End If
                            Case Else
                                Call File_Error(sts, com + BtSNoWait, "品目マスタ")
                                GoTo Abort_Tran
                       End Select
        
                    Loop
                End If
    
    
    
    
    
    
    
    
    
    
            End If
        Next i
    End If





'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>





    '---------------------------------------------------    '大阪事　見積用品目マスタ  該当データ全件削除
    '該当データ全件削除
    Call UniCode_Conv(K0_ITEM_O.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM_O.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM_O.HIN_GAI, Text1(ptxHin_Gai).Text)
       
    Call UniCode_Conv(K0_ITEM_O.SEQ_NO, "")
       
    com = BtOpGetGreaterEqual
       
    Do
        
        DoEvents
        
        Do
        
            sts = BTRV(com + BtSNoWait, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
                
            Select Case sts
                Case BtNoErr
                
                        If StrConv(ITEM_O_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                           StrConv(ITEM_O_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                                Trim(StrConv(ITEM_O_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHin_Gai).Text) Then
                        sts = BTRV(BtOpUnlock, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "大阪事　見積用品目マスタ")
                            GoTo Abort_Tran
                        End If
                        sts = BtErrEOF
                    End If
                    Exit Do
                Case BtErrEOF
                    Exit Do
                
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ITEM_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        GoTo Abort_Tran
                    End If
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "大阪事　見積用品目マスタ")
                    
                    GoTo Abort_Tran
            End Select
    
        Loop
            
        If sts = BtErrEOF Then
            Exit Do
        End If





        Do
            sts = BTRV(BtOpDelete, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
            Select Case sts
                Case BtNoErr
                
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_CMOPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "大阪事　見積用品目マスタ")
                        End If
                        GoTo Abort_Tran
                    End If
                
                
                Case Else
                    Call File_Error(sts, BtOpDelete, "大阪事　見積用品目マスタ")
                    GoTo Abort_Tran
            End Select
        Loop
    
        com = BtOpGetNext
    
    Loop




    '---------------------------------------------------    '大阪事　見積用品目マスタ ヘッダー出力
    
    
    Call Rclr_ITEM_O_REC
    
    Call UniCode_Conv(ITEM_O_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
    Call UniCode_Conv(ITEM_O_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
    Call UniCode_Conv(ITEM_O_REC.HIN_GAI, Text1(ptxHin_Gai).Text)
    Call UniCode_Conv(ITEM_O_REC.SEQ_NO, "000")
    
    
    Call UniCode_Conv(ITEM_O_REC.KO_JGYOBU, "")
    Call UniCode_Conv(ITEM_O_REC.KO_NAIGAI, "")
    Call UniCode_Conv(ITEM_O_REC.KO_HIN_GAI, "")
    
    
    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_TANI, Text1(ptxNAKANISHI_TANI).Text)         '中西工料　単位
    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_KIN, Text1(ptxNAKANISHI_KIN).Text)           '中西工料　金額

    Call UniCode_Conv(ITEM_O_REC.SHOHIN_TANI, Text1(ptxSHOHIN_TANI).Text)               '商品化工料　単位
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_KIN, Text1(ptxSHOHIN_KIN).Text)                 '商品化工料　金額

    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_TANI, Text1(ptxPF_KAKOU_TANI).Text)           'PF加工　単位
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_KIN, Text1(ptxPF_KAKOU_KIN).Text)            'PF加工　金額

    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_TANI, Text1(ptxPF_KAKOU_TANI).Text)           'PE加工　単位
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_KIN, Text1(ptxPE_KAKOU_KIN).Text)             'PE加工　金額

    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_TANI, Text1(ptxPE_SHIZAI_TANI).Text)         'PF資材　単位
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_KIN, Text1(ptxPE_SHIZAI_KIN).Text)           'PF資材　金額

    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_TANI, Text1(ptxHINBAN_LABEL_TANI).Text)   '品番表示ﾗﾍﾞﾙ　単位
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_KIN, Text1(ptxHINBAN_LABEL_KIN).Text)     '品番表示ﾗﾍﾞﾙ　金額

    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_TANI, Text1(ptxKOUJI_SETSU_TANI).Text)     '設置工事説明書　単位
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_KIN, Text1(ptxKOUJI_SETSU_KIN).Text)       '設置工事説明書　金額

    Call UniCode_Conv(ITEM_O_REC.KONPOU_TANI, Text1(ptxKONPOU_TANI).Text)               '梱包材　単位
    Call UniCode_Conv(ITEM_O_REC.KONPOU_KIN, Text1(ptxKONPOU_KIN).Text)                 '梱包材　金額

    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_TANI, Text1(ptxFUKU_SHIZAI_TANI).Text)     '副資材　単位
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_KIN, Text1(ptxFUKU_SHIZAI_KIN).Text)       '副資材　金額

    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_TANI, Text1(ptxKONPOU_ASSY_TANI).Text)     '梱包ASSY　単位
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_KIN, Text1(ptxKONPOU_ASSY_KIN).Text)       '梱包ASSY　金額

    Call UniCode_Conv(ITEM_O_REC.KANRI_TANI, Text1(ptxKANRI_TANI).Text)                 '管理費　単位
    Call UniCode_Conv(ITEM_O_REC.KANRI_KIN, Text1(ptxKANRI_KIN).Text)                   '管理費　金額
    
    Call UniCode_Conv(ITEM_O_REC.GOUKEI_KIN, CLng(lblGOUKEI_KIN.Caption))               '合計　金額
    Call UniCode_Conv(ITEM_O_REC.INPUT_TANTO_CODE, Text1(ptxTanto_Code).Text)           '入力担当者ｺｰﾄﾞ

    
    Call UniCode_Conv(ITEM_O_REC.BUZAI_TANTO_NAME, Text1(ptxBuzai_Tanto_Name).Text)     '部材担当者名
    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_T_KIN, Text1(ptxNAKANISHI_T_KIN).Text)       '中西工料　提出金額
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_T_KIN, Text1(ptxSHOHIN_T_KIN).Text)             '商品化工料　提出金額
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_T_KIN, Text1(ptxPF_KAKOU_T_KIN).Text)         'PF加工　提出金額
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_T_KIN, Text1(ptxPE_KAKOU_T_KIN).Text)         'PE加工　提出金額
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_T_KIN, Text1(ptxPE_SHIZAI_T_KIN).Text)       'PF資材　提出金額
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_T_KIN, Text1(ptxHINBAN_LABEL_T_KIN).Text) '品番表示ﾗﾍﾞﾙ　提出金額
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_T_KIN, Text1(ptxKOUJI_SETSU_T_KIN).Text)   '設置工事説明書　提出金額
    Call UniCode_Conv(ITEM_O_REC.KONPOU_T_KIN, Text1(ptxKONPOU_T_KIN).Text)             '梱包材　提出金額
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_T_KIN, Text1(ptxFUKU_SHIZAI_T_KIN).Text)   '副資材　提出金額
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_T_KIN, Text1(ptxKONPOU_ASSY_T_KIN).Text)   '梱包ASSY　提出金額
    Call UniCode_Conv(ITEM_O_REC.KANRI_T_KIN, Text1(ptxKANRI_T_KIN).Text)               '管理費　提出金額
    Call UniCode_Conv(ITEM_O_REC.GOUKEI_T_KIN, CLng(lblGOUKEI_T_KIN.Caption))           '提出合計金額

    
    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_F, Check1(chkNAKANISHI_F).Value)             '中西工料 見積書表示ﾌﾗｸ
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_F, Check1(chkSHOHIN_F).Value)                   '商品化工料 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_F, Check1(chkPF_KAKOU_F).Value)               'PF加工 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_F, Check1(chkPE_KAKOU_F).Value)               'PE加工 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_F, Check1(chkPE_SHIZAI_F).Value)             'PF資材 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_F, Check1(chkHINBAN_LABEL_F).Value)       '品番表示ﾗﾍﾞﾙ 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_F, Check1(chkKOUJI_SETSU_F).Value)         '設置工事説明書 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(ITEM_O_REC.KONPOU_F, Check1(chkKONPOU_F).Value)                   '梱包材 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_F, Check1(chkFUKU_SHIZAI_F).Value)         '副資材 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_F, Check1(chkKONPOU_ASSY_F).Value)         '梱包ASSY 見積書表示ﾌﾗｸﾞ
    Call UniCode_Conv(ITEM_O_REC.KANRI_F, Check1(chkKANRI_F).Value)                     '管理費 見積書表示ﾌﾗｸﾞ

    
    
    
    
'>>>>>>>>>>>>>>>>>> 2017.11.08
    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_QTY, Text1(ptxNAKANISHI_QTY).Text)           '中西工料　数量
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_QTY, Text1(ptxSHOHIN_QTY).Text)                 '商品化工料　数量
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_QTY, Text1(ptxPF_KAKOU_QTY).Text)             'PF加工　数量
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_QTY, Text1(ptxPE_KAKOU_QTY).Text)             'PE加工　数量
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_QTY, Text1(ptxPE_SHIZAI_QTY).Text)           'PE資材　数量
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_QTY, Text1(ptxHINBAN_LABEL_QTY).Text)     '品番表示ﾗﾍﾞﾙ　数量
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_QTY, Text1(ptxKOUJI_SETSU_QTY).Text)       '設置工事説明書　数量
    Call UniCode_Conv(ITEM_O_REC.KONPOU_QTY, Text1(ptxKONPOU_QTY).Text)                 '梱包材　数量
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_QTY, Text1(ptxFUKU_SHIZAI_QTY).Text)       '副資材　数量
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_QTY, Text1(ptxKONPOU_ASSY_QTY).Text)       '梱包ASSY　数量
    Call UniCode_Conv(ITEM_O_REC.KANRI_QTY, Text1(ptxKANRI_QTY).Text)                   '管理費　数量
    
    
    Call UniCode_Conv(ITEM_O_REC.NAKANISHI_T_TAN, Text1(ptxNAKANISHI_T_TAN).Text)       '中西工料　提出単価
    Call UniCode_Conv(ITEM_O_REC.SHOHIN_T_TAN, Text1(ptxSHOHIN_T_TAN).Text)             '商品化工料　提出単価
    Call UniCode_Conv(ITEM_O_REC.PF_KAKOU_T_TAN, Text1(ptxPF_KAKOU_T_TAN).Text)         'PF加工　提出単価
    Call UniCode_Conv(ITEM_O_REC.PE_KAKOU_T_TAN, Text1(ptxPE_KAKOU_T_TAN).Text)         'PE加工　提出単価
    Call UniCode_Conv(ITEM_O_REC.PE_SHIZAI_T_TAN, Text1(ptxPE_SHIZAI_T_TAN).Text)       'PE資材　提出単価
    Call UniCode_Conv(ITEM_O_REC.HINBAN_LABEL_T_TAN, Text1(ptxHINBAN_LABEL_T_TAN).Text) '品番表示ﾗﾍﾞﾙ　提出単価
    Call UniCode_Conv(ITEM_O_REC.KOUJI_SETSU_T_TAN, Text1(ptxKOUJI_SETSU_T_TAN).Text)   '設置工事説明書　提出単価
    Call UniCode_Conv(ITEM_O_REC.KONPOU_T_TAN, Text1(ptxKONPOU_T_TAN).Text)             '梱包材　提出単価
    Call UniCode_Conv(ITEM_O_REC.FUKU_SHIZAI_T_TAN, Text1(ptxFUKU_SHIZAI_T_TAN).Text)   '副資材　提出単価
    Call UniCode_Conv(ITEM_O_REC.KONPOU_ASSY_T_TAN, Text1(ptxKONPOU_ASSY_T_TAN).Text)   '梱包ASSY　提出単価
    Call UniCode_Conv(ITEM_O_REC.KANRI_T_TAN, Text1(ptxKANRI_T_TAN).Text)               '管理費　提出単価
'>>>>>>>>>>>>>>>>>> 2017.11.08
    
    
    
    
    
    
    
    
    
    
    
    Call UniCode_Conv(ITEM_O_REC.UPD_TANTO, "SEI16")
    Call UniCode_Conv(ITEM_O_REC.UPD_DATETIME, Format(Now, "yyyyymmddhhmmss"))
    
    
    Do
        
        
        sts = BTRV(BtOpInsert, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
                
        
        Select Case sts
            Case BtNoErr
            
                Exit Do
            
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Do
                End If
            
            
            Case Else
                Call File_Error(sts, upd_com, "大阪事　見積用品目マスタ")
                Exit Function
    
        End Select
    Loop

    '---------------------------------------------------    '大阪事　見積用品目マスタ ボディ出力
    
'   Set TDBGrid1(pGrdKOUSEI).Array = KOUSEI
    
    
'   TDBGrid1(pGrdKOUSEI).Update




    If TDBGrid1(pGrdKOUSEI).ApproxCount = 0 Then

    Else


        For i = 1 To KOUSEI.UpperBound(1)
    
    
            If Trim(KOUSEI(i, ColKO_S_HIN_GAI)) = "" And Trim(KOUSEI(i, ColKO_T_HIN_NAME)) = "" Then
            Else
                                                                                                            
                                                                                            
                                                                                            
                
                Call Rclr_ITEM_O_REC
                
                
                Call UniCode_Conv(ITEM_O_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
                Call UniCode_Conv(ITEM_O_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
                Call UniCode_Conv(ITEM_O_REC.HIN_GAI, Text1(ptxHin_Gai).Text)
                
                Call UniCode_Conv(ITEM_O_REC.SEQ_NO, Format(i, "000"))          '2017.11.08
                
                
                Call UniCode_Conv(ITEM_O_REC.KO_JGYOBU, Right(KOUSEI(i, ColKO_JGYOBU), 1))
                Call UniCode_Conv(ITEM_O_REC.KO_NAIGAI, NAIGAI_NAI)
                Call UniCode_Conv(ITEM_O_REC.KO_HIN_GAI, KOUSEI(i, ColKO_S_HIN_GAI))
                                                                                            
                Call UniCode_Conv(ITEM_O_REC.KO_QTY, KOUSEI(i, ColKO_QTY))      '2017.11.08
                                                                                
                                                                                
                                                                                
                                                                                
                Call UniCode_Conv(ITEM_O_REC.T_HIN_NAME, KOUSEI(i, ColKO_T_HIN_NAME))
                Call UniCode_Conv(ITEM_O_REC.TANI, KOUSEI(i, ColKO_TANI))
                Call UniCode_Conv(ITEM_O_REC.T_TANKA, KOUSEI(i, ColKO_T_TANKA))
                Call UniCode_Conv(ITEM_O_REC.T_KINGAKU, KOUSEI(i, ColKO_T_KINGAKU))
                                                                                
                                                                                
                                                                                '種別
                Call UniCode_Conv(ITEM_O_REC.KO_SYUBETSU, Right(KOUSEI(i, ColKO_SYUBETSU), 2))       '2017.11.16
                                                                                
                                                                                
                                                                                
                
                
                
            
                Call UniCode_Conv(ITEM_O_REC.UPD_TANTO, "SEI16")                                '更新担当者ｺｰﾄﾞ
                                                                                                '更新日時
                Call UniCode_Conv(ITEM_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
            
            
                Do
                    
                    DoEvents
                    
                    sts = BTRV(BtOpInsert, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
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
                            Call File_Error(sts, BtOpInsert, "大阪事　見積用品目マスタ")
                            GoTo Abort_Tran
                    End Select
                
                Loop
    
    
    
                
                
    
    
    
    
    
    
    
    
    
    
    
            End If
        Next i
    End If


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.09.28  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>




End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
    Call Input_UnLock
    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    
    Call Input_UnLock
                                        




End Function

Private Sub Text1_LostFocus(Index As Integer)
    
Dim i   As Integer
    
    
    Select Case Index
        Case ptxHin_Gai
            
            
            Text1(Index).Text = StrConv(RTrim(Text1(Index).Text), vbUpperCase)
            
            
            If Trim(svHin_Gai) = (Text1(ptxHin_Gai).Text) Then
            Else
                For i = 2 To 5
                    Command1(i).Enabled = False
                Next i
            
            
            End If
    End Select
End Sub
Private Function ITEM_O_RESET_PROC() As Integer
'----------------------------------------------------------------------------
'                   大阪事　見積用品目マスタ  再構築
'----------------------------------------------------------------------------
Dim com         As Integer
Dim sts         As Integer
Dim NEW_ITEM    As Integer

Dim ans         As Integer

Dim LAST_SEQ_NO As Integer

    ITEM_O_RESET_PROC = True

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2017.11.14  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    '---------------------------------------------------    '見積用品目マスタ(tmp)  該当データ全件削除し　見積用品目マスタ→見積用品目マスタ(tmp)


    com = BtOpGetFirst
       
    Do
        
        DoEvents
        
        Do
        
            sts = BTRV(com + BtSNoWait, tmpITEM_O_POS, tmpITEM_O_REC, Len(tmpITEM_O_REC), K0_tmpITEM_O, Len(K0_tmpITEM_O), 0)
                
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrEOF
                    Exit Do
                
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ITEM_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "大阪事　見積用品目マスタ")
                    
                    Exit Function
            End Select
    
        Loop
            
        If sts = BtErrEOF Then
            Exit Do
        End If


        Do
            sts = BTRV(BtOpDelete, tmpITEM_O_POS, tmpITEM_O_REC, Len(tmpITEM_O_REC), K0_tmpITEM_O, Len(K0_tmpITEM_O), 0)
            Select Case sts
                Case BtNoErr
                
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_CMOPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, tmpITEM_O_POS, tmpITEM_O_REC, Len(tmpITEM_O_REC), K0_tmpITEM_O, Len(K0_tmpITEM_O), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "大阪事　見積用品目マスタ")
                        End If
                        Exit Function
                    End If
            
                
                Case Else
                    Call File_Error(sts, BtOpDelete, "大阪事　見積用品目マスタ")
                    Exit Function
            End Select
        Loop
    
        com = BtOpGetNext
    
    Loop
'----------------------------------- 構成品の内容を構成マスタより再構築する  2017.11.15 ------------------------------
    '---------------------------------------------------    '大阪事　見積用品目マスタ  該当データ全件削除
    '該当データ全件削除
    Call UniCode_Conv(K0_ITEM_O.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM_O.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM_O.HIN_GAI, Text1(ptxHin_Gai).Text)
       
    Call UniCode_Conv(K0_ITEM_O.SEQ_NO, "")
       
    com = BtOpGetGreaterEqual
       
    Do
        
        DoEvents
        sts = BTRV(com, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
            
        Select Case sts
            Case BtNoErr
            
                    If StrConv(ITEM_O_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                       StrConv(ITEM_O_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                            Trim(StrConv(ITEM_O_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHin_Gai).Text) Then
                    sts = BTRV(BtOpUnlock, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
                    If sts Then
                        Call File_Error(sts, BtOpUnlock, "大阪事　見積用品目マスタ")
                        Exit Function
                    End If
                End If
                Exit Do
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com + BtSNoWait, "大阪事　見積用品目マスタ")
                
                Exit Function
        End Select


        Do
            DoEvents
            sts = BTRV(BtOpInsert, tmpITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_tmpITEM_O, Len(K0_tmpITEM_O), 0)
            Select Case sts
                Case BtNoErr
                
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_CMOPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "大阪事　見積用品目マスタ")
                        End If
                        Exit Function
                    End If
                
                
                Case Else
                    Call File_Error(sts, BtOpInsert, "大阪事　見積用品目マスタ")
                    Exit Function
            End Select
        Loop



        Do
        DoEvents
        sts = BTRV(BtOpDelete, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
            Select Case sts
                Case BtNoErr
                
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_CMOPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "大阪事　見積用品目マスタ")
                        End If
                        Exit Function
                    End If
                
                
                Case Else
                    Call File_Error(sts, BtOpDelete, "大阪事　見積用品目マスタ")
                    Exit Function
            End Select
        Loop
    
        com = BtOpGetNext
    
    Loop
    
    
    Call UniCode_Conv(K0_tmpITEM_O.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_tmpITEM_O.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_tmpITEM_O.HIN_GAI, Text1(ptxHin_Gai).Text)

    Call UniCode_Conv(K0_ITEM_O.SEQ_NO, "000")     '2017.11.08

    sts = BTRV(BtOpGetEqual, tmpITEM_O_POS, tmpITEM_O_REC, Len(tmpITEM_O_REC), K0_tmpITEM_O, Len(K0_tmpITEM_O), 0)
    Select Case sts
        Case BtNoErr
            NEW_ITEM = 0
        
        Case BtErrKeyNotFound

            NEW_ITEM = 1
        Case Else
            Call File_Error(sts, BtOpGetEqual, "大阪事 見積用品目マスタ")
            Exit Function

    End Select
    
    If NEW_ITEM = 1 Then
        ITEM_O_RESET_PROC = False
        Exit Function
    End If
    
    Do
        DoEvents
        sts = BTRV(BtOpInsert, ITEM_O_POS, tmpITEM_O_REC, Len(tmpITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
        Select Case sts
            Case BtNoErr
            
                Exit Do
            Case BtErrEOF
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_CMOPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
                    If sts Then
                        Call File_Error(sts, BtOpUnlock, "大阪事　見積用品目マスタ")
                    End If
                    Exit Function
                End If
            
            
            Case Else
                Call File_Error(sts, BtOpDelete, "大阪事　見積用品目マスタ")
                Exit Function
        End Select
    Loop

    Do
        DoEvents
        sts = BTRV(BtOpDelete, tmpITEM_O_POS, tmpITEM_O_REC, Len(tmpITEM_O_REC), K0_tmpITEM_O, Len(K0_tmpITEM_O), 0)
        Select Case sts
            Case BtNoErr
            
                Exit Do
            Case BtErrEOF
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_CMOPO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
                    If sts Then
                        Call File_Error(sts, BtOpUnlock, "大阪事　見積用品目マスタ")
                    End If
                    Exit Function
                End If
            
            
            Case Else
                Call File_Error(sts, BtOpDelete, "大阪事　見積用品目マスタ")
                Exit Function
        End Select
    Loop


    
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Text1(ptxHin_Gai).Text)

    Call UniCode_Conv(K0_P_COMPO.SEQNO, "001")
    
    com = BtOpGetGreater
    
    Do
        sts = BTRV(com, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_ITEM_O), 0)
            
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_COMPO_O_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                   StrConv(P_COMPO_O_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                    Trim(StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHin_Gai).Text) Then
                    
                    Exit Do
                
                End If
            Case BtErrEOF
                
                Exit Do
            
            Case Else
                Call File_Error(sts, com + BtSNoWait, "大阪事　見積用品目マスタ")
                
                Exit Function
        
        End Select
    
        
        
        Call UniCode_Conv(K1_tmpITEM_O.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K1_tmpITEM_O.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
        Call UniCode_Conv(K1_tmpITEM_O.HIN_GAI, Text1(ptxHin_Gai).Text)
        
        
        Call UniCode_Conv(K1_tmpITEM_O.KO_JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
        Call UniCode_Conv(K1_tmpITEM_O.KO_NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(K1_tmpITEM_O.KO_HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
        
        
        sts = BTRV(BtOpGetEqual, tmpITEM_O_POS, tmpITEM_O_REC, Len(tmpITEM_O_REC), K1_tmpITEM_O, Len(K1_tmpITEM_O), 1)
            
        Select Case sts
            Case BtNoErr
            
            
             '   Call UniCode_Conv(tmpITEM_O_REC.KO_SYUBETSU, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode)) '2017.11.29
            
                Call UniCode_Conv(tmpITEM_O_REC.SEQ_NO, StrConv(P_COMPO_K_REC.SEQNO, vbUnicode))
                
                LAST_SEQ_NO = Val(StrConv(P_COMPO_K_REC.SEQNO, vbUnicode))
                
                Call UniCode_Conv(tmpITEM_O_REC.KO_QTY, StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))
                
                Call UniCode_Conv(tmpITEM_O_REC.UPD_TANTO, "SEI16")                                '更新担当者ｺｰﾄﾞ
                                                                                                '更新日時
                Call UniCode_Conv(tmpITEM_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                
                
                sts = BTRV(BtOpInsert, ITEM_O_POS, tmpITEM_O_REC, Len(tmpITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
                    
                Select Case sts
                    Case BtNoErr
                    Case Else
                        Call File_Error(sts, com + BtSNoWait, "大阪事　見積用品目マスタ")
                        
                        Exit Function
                
                End Select
            
                sts = BTRV(BtOpDelete, tmpITEM_O_POS, tmpITEM_O_REC, Len(tmpITEM_O_REC), K0_tmpITEM_O, Len(K0_tmpITEM_O), 1)
                    
                Select Case sts
                    Case BtNoErr
                    Case Else
                        Call File_Error(sts, com + BtSNoWait, "大阪事　見積用品目マスタ")
                        
                        Exit Function
                
                End Select
            
            
            
            Case BtErrKeyNotFound
                
            
                Call Rclr_ITEM_O_REC
                
                
                Call UniCode_Conv(ITEM_O_REC.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 3, 1))
                Call UniCode_Conv(ITEM_O_REC.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 4, 1))
                Call UniCode_Conv(ITEM_O_REC.HIN_GAI, Text1(ptxHin_Gai).Text)
                
                Call UniCode_Conv(ITEM_O_REC.SEQ_NO, StrConv(P_COMPO_K_REC.SEQNO, vbUnicode))
                LAST_SEQ_NO = Val(StrConv(P_COMPO_K_REC.SEQNO, vbUnicode))
                
                
                Call UniCode_Conv(ITEM_O_REC.KO_JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                Call UniCode_Conv(ITEM_O_REC.KO_NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                Call UniCode_Conv(ITEM_O_REC.KO_HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                                                                                            
                Call UniCode_Conv(ITEM_O_REC.KO_QTY, StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))
                                                                                
                                                                                
                                                                                
            
                Call UniCode_Conv(ITEM_O_REC.INS_TANTO, "SEI16")                                '更新担当者ｺｰﾄﾞ
                                                                                                '更新日時
                Call UniCode_Conv(ITEM_O_REC.Ins_DateTime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                                                                                
                
                
            
                Call UniCode_Conv(ITEM_O_REC.INS_TANTO, "SEI16")                                '更新担当者ｺｰﾄﾞ
                                                                                                '更新日時
                Call UniCode_Conv(ITEM_O_REC.Ins_DateTime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
            
            
                
                sts = BTRV(BtOpInsert, ITEM_O_POS, ITEM_O_REC, Len(ITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
                Select Case sts
                    Case BtNoErr
                    Case Else
                        Call File_Error(sts, BtOpInsert, "大阪事　見積用品目マスタ")
                        Exit Function
                End Select
            
            Case Else
                Call File_Error(sts, com + BtSNoWait, "大阪事　見積用品目マスタ")
                
                Exit Function
        
        End Select
        
        
        
        
        
        
        
        
        com = BtOpGetNext
    
    
    Loop



    Call UniCode_Conv(K0_tmpITEM_O.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_tmpITEM_O.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_tmpITEM_O.HIN_GAI, Text1(ptxHin_Gai).Text)

    Call UniCode_Conv(K0_tmpITEM_O.SEQ_NO, "")
    
    com = BtOpGetGreater
    
    Do
        sts = BTRV(com, tmpITEM_O_POS, tmpITEM_O_REC, Len(tmpITEM_O_REC), K0_tmpITEM_O, Len(K0_tmpITEM_O), 0)
            
        Select Case sts
            Case BtNoErr
            
                If StrConv(tmpITEM_O_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                   StrConv(tmpITEM_O_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                    Trim(StrConv(tmpITEM_O_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHin_Gai).Text) Then
                    
                    Exit Do
                
                End If
            Case BtErrEOF
                
                Exit Do
            
            Case Else
                Call File_Error(sts, com + BtSNoWait, "大阪事　見積用品目マスタ")
                
                Exit Function
        
        End Select
    
        
        
        If Trim(StrConv(tmpITEM_O_REC.KO_HIN_GAI, vbUnicode)) = "" Then
            LAST_SEQ_NO = LAST_SEQ_NO + 10
        
            Call UniCode_Conv(tmpITEM_O_REC.SEQ_NO, Format(LAST_SEQ_NO, "000"))
        
            sts = BTRV(BtOpInsert, ITEM_O_POS, tmpITEM_O_REC, Len(tmpITEM_O_REC), K0_ITEM_O, Len(K0_ITEM_O), 0)
                
            Select Case sts
                Case BtNoErr
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "大阪事　見積用品目マスタ")
                    
                    Exit Function
            
            End Select
        
        End If
            
        sts = BTRV(BtOpDelete, tmpITEM_O_POS, tmpITEM_O_REC, Len(tmpITEM_O_REC), K0_tmpITEM_O, Len(K0_tmpITEM_O), 1)
            
        Select Case sts
            Case BtNoErr
            Case Else
                Call File_Error(sts, com + BtSNoWait, "大阪事　見積用品目マスタ")
                
                Exit Function
        
        End Select
        
        
        com = BtOpGetNext
    
    
    Loop

    ITEM_O_RESET_PROC = False

End Function
