VERSION 5.00
Begin VB.Form PM000103 
   Caption         =   "管理マスタメンテナンス"
   ClientHeight    =   6300
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   12045
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
   ScaleHeight     =   6300
   ScaleWidth      =   12045
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   30
      Left            =   9660
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4440
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   29
      Left            =   9660
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4080
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   28
      Left            =   9660
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3720
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   27
      Left            =   9660
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3360
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   26
      Left            =   9660
      MaxLength       =   3
      TabIndex        =   26
      Top             =   3000
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   25
      Left            =   9660
      MaxLength       =   3
      TabIndex        =   25
      Top             =   2640
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   24
      Left            =   9660
      MaxLength       =   3
      TabIndex        =   24
      Top             =   2280
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   23
      Left            =   9660
      MaxLength       =   3
      TabIndex        =   23
      Top             =   1920
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   22
      Left            =   9660
      MaxLength       =   3
      TabIndex        =   22
      Top             =   1560
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   21
      Left            =   9660
      MaxLength       =   3
      TabIndex        =   21
      Top             =   1200
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   20
      Left            =   5775
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4440
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   19
      Left            =   5775
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4080
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   18
      Left            =   5775
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3720
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   17
      Left            =   5775
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3360
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   16
      Left            =   5775
      MaxLength       =   3
      TabIndex        =   16
      Top             =   3000
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   15
      Left            =   5775
      MaxLength       =   3
      TabIndex        =   15
      Top             =   2640
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   14
      Left            =   5775
      MaxLength       =   3
      TabIndex        =   14
      Top             =   2280
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   13
      Left            =   5775
      MaxLength       =   3
      TabIndex        =   13
      Top             =   1920
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   12
      Left            =   5775
      MaxLength       =   3
      TabIndex        =   12
      Top             =   1560
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   11
      Left            =   5775
      MaxLength       =   3
      TabIndex        =   11
      Top             =   1200
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   10
      Left            =   3675
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4440
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   9
      Left            =   3675
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4080
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   3675
      MaxLength       =   3
      TabIndex        =   8
      Top             =   3720
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   3675
      MaxLength       =   3
      TabIndex        =   7
      Top             =   3360
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   3675
      MaxLength       =   3
      TabIndex        =   6
      Top             =   3000
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   3675
      MaxLength       =   3
      TabIndex        =   5
      Top             =   2640
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   3675
      MaxLength       =   3
      TabIndex        =   4
      Top             =   2280
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   3675
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1920
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   3675
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1560
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   360
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   3675
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1200
      Width           =   480
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   4  '全角ひらがな
      Index           =   0
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   375
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
      TabIndex        =   42
      Top             =   5880
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
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   3
      Left            =   2640
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   31
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "ラベル貼り"
      Height          =   300
      Index           =   14
      Left            =   4305
      TabIndex        =   57
      Top             =   1560
      Width           =   1440
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Height          =   300
      Index           =   33
      Left            =   6405
      TabIndex        =   76
      Top             =   4440
      Width           =   3285
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Height          =   300
      Index           =   32
      Left            =   6405
      TabIndex        =   75
      Top             =   4080
      Width           =   3285
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Height          =   300
      Index           =   31
      Left            =   6405
      TabIndex        =   74
      Top             =   3720
      Width           =   3285
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Height          =   300
      Index           =   30
      Left            =   6405
      TabIndex        =   73
      Top             =   3360
      Width           =   3285
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "後片付け"
      Height          =   300
      Index           =   29
      Left            =   6405
      TabIndex        =   72
      Top             =   3000
      Width           =   3285
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "不適合処理"
      Height          =   300
      Index           =   28
      Left            =   6405
      TabIndex        =   71
      Top             =   2640
      Width           =   3285
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "同梱部品引き落とし伝票入力"
      Height          =   300
      Index           =   27
      Left            =   6405
      TabIndex        =   70
      Top             =   2280
      Width           =   3285
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "同梱部品引き落とし伝票発行"
      Height          =   300
      Index           =   26
      Left            =   6405
      TabIndex        =   69
      Top             =   1920
      Width           =   3285
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "部品搬入"
      Height          =   300
      Index           =   25
      Left            =   6405
      TabIndex        =   68
      Top             =   1560
      Width           =   3285
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "検査表記入"
      Height          =   300
      Index           =   24
      Left            =   6405
      TabIndex        =   67
      Top             =   1200
      Width           =   3285
   End
   Begin VB.Label Label 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "後工程"
      Height          =   300
      Index           =   23
      Left            =   6405
      TabIndex        =   66
      Top             =   840
      Width           =   3720
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Height          =   300
      Index           =   22
      Left            =   4305
      TabIndex        =   65
      Top             =   4440
      Width           =   1440
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Height          =   300
      Index           =   21
      Left            =   4305
      TabIndex        =   64
      Top             =   4080
      Width           =   1440
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Height          =   300
      Index           =   20
      Left            =   4305
      TabIndex        =   63
      Top             =   3720
      Width           =   1440
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Height          =   300
      Index           =   19
      Left            =   4305
      TabIndex        =   62
      Top             =   3360
      Width           =   1440
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "集合梱包"
      Height          =   300
      Index           =   18
      Left            =   4305
      TabIndex        =   61
      Top             =   3000
      Width           =   1440
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "同梱"
      Height          =   300
      Index           =   17
      Left            =   4305
      TabIndex        =   60
      Top             =   2640
      Width           =   1440
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "加工"
      Height          =   300
      Index           =   16
      Left            =   4305
      TabIndex        =   59
      Top             =   2280
      Width           =   1440
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "個装作業"
      Height          =   300
      Index           =   15
      Left            =   4305
      TabIndex        =   58
      Top             =   1920
      Width           =   1440
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "見本確認"
      Height          =   300
      Index           =   13
      Left            =   4305
      TabIndex        =   56
      Top             =   1200
      Width           =   1440
   End
   Begin VB.Label Label 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "作業工程"
      Height          =   300
      Index           =   12
      Left            =   4305
      TabIndex        =   55
      Top             =   840
      Width           =   1980
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Height          =   300
      Index           =   11
      Left            =   735
      TabIndex        =   54
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Height          =   300
      Index           =   10
      Left            =   735
      TabIndex        =   53
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "指図票発行"
      Height          =   300
      Index           =   9
      Left            =   735
      TabIndex        =   52
      Top             =   3720
      Width           =   2940
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "ラベル発行"
      Height          =   300
      Index           =   8
      Left            =   735
      TabIndex        =   51
      Top             =   3360
      Width           =   2940
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "同梱部品準備"
      Height          =   300
      Index           =   7
      Left            =   735
      TabIndex        =   50
      Top             =   3000
      Width           =   2970
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "副資材準備"
      Height          =   300
      Index           =   6
      Left            =   735
      TabIndex        =   49
      Top             =   2640
      Width           =   2940
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "部品準備"
      Height          =   300
      Index           =   5
      Left            =   735
      TabIndex        =   48
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "同梱部品在庫有無確認"
      Height          =   300
      Index           =   4
      Left            =   735
      TabIndex        =   47
      Top             =   1920
      Width           =   2985
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "副資材在庫有無確認"
      Height          =   300
      Index           =   3
      Left            =   735
      TabIndex        =   46
      Top             =   1560
      Width           =   2955
   End
   Begin VB.Label Label 
      BorderStyle     =   1  '実線
      Caption         =   "事前商品化部品／数量選定"
      Height          =   345
      Index           =   2
      Left            =   735
      TabIndex        =   45
      Top             =   1200
      Width           =   2940
   End
   Begin VB.Label Label 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "前工程"
      Height          =   420
      Index           =   1
      Left            =   735
      TabIndex        =   44
      Top             =   840
      Width           =   3405
   End
   Begin VB.Label Label 
      Caption         =   "ﾚｺｰﾄﾞ№"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   43
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PM000103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxRec_No% = 0                'ﾚｺｰﾄﾞ№(入力不可)

Private Const ptxBEF_KOTEI01% = 1           '前工程01
Private Const ptxBEF_KOTEI02% = 2           '前工程02
Private Const ptxBEF_KOTEI03% = 3           '前工程03
Private Const ptxBEF_KOTEI04% = 4           '前工程04
Private Const ptxBEF_KOTEI05% = 5           '前工程05
Private Const ptxBEF_KOTEI06% = 6           '前工程06
Private Const ptxBEF_KOTEI07% = 7           '前工程07
Private Const ptxBEF_KOTEI08% = 8           '前工程08
Private Const ptxBEF_KOTEI09% = 9           '前工程09
Private Const ptxBEF_KOTEI10% = 10          '前工程10

Private Const ptxMAIN_KOTEI01% = 11         '作業工程01
Private Const ptxMAIN_KOTEI02% = 12         '作業工程02
Private Const ptxMAIN_KOTEI03% = 13         '作業工程03
Private Const ptxMAIN_KOTEI04% = 14         '作業工程04
Private Const ptxMAIN_KOTEI05% = 15         '作業工程05
Private Const ptxMAIN_KOTEI06% = 16         '作業工程06
Private Const ptxMAIN_KOTEI07% = 17         '作業工程07
Private Const ptxMAIN_KOTEI08% = 18         '作業工程08
Private Const ptxMAIN_KOTEI09% = 19         '作業工程09
Private Const ptxMAIN_KOTEI10% = 20         '作業工程10


Private Const ptxAFT_KOTEI01% = 21          '後工程01
Private Const ptxAFT_KOTEI02% = 22          '後工程02
Private Const ptxAFT_KOTEI03% = 23          '後工程03
Private Const ptxAFT_KOTEI04% = 24          '後工程04
Private Const ptxAFT_KOTEI05% = 25          '後工程05
Private Const ptxAFT_KOTEI06% = 26          '後工程06
Private Const ptxAFT_KOTEI07% = 27          '後工程07
Private Const ptxAFT_KOTEI08% = 28          '後工程08
Private Const ptxAFT_KOTEI09% = 29          '後工程09
Private Const ptxAFT_KOTEI10% = 30          '後工程10
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PM000103.MousePointer = vbHourglass

    Call Ctrl_Lock(PM000103)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM000103)


    PM000103.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts As Integer
    
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        Case ptxBEF_KOTEI01 To ptxAFT_KOTEI10             '前工程01～後工程10

            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function
'Private Function Item_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   画面表示
'----------------------------------------------------------------------------
'Dim sts As Integer
'Dim i   As Integer
'Dim j   As Integer
'
'    Item_Disp_Proc = True
'
'
'    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_DEF_No)
'
'    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC02, Len(P_KANRIREC02), K0_P_KANRI, Len(K0_P_KANRI), 0)
'    Select Case sts
'        Case BtNoErr
'        Case BtErrKeyNotFound
'
'            For i = 0 To 9
'                Call UniCode_Conv(P_KANRIREC02.BEF_KOTEI(i).KOTEI, "")
'            Next i
'
'            For i = 0 To 9
'                Call UniCode_Conv(P_KANRIREC02.MAIN_KOTEI(i).KOTEI, "")
'            Next i
'
'            For i = 0 To 9
'                Call UniCode_Conv(P_KANRIREC02.AFT_KOTEI(i).KOTEI, "")
'            Next i
'
'            For i = 0 To 4
'                Call UniCode_Conv(P_KANRIREC02.FUTAI_KOTEI(i).KOTEI, "")
'            Next i
'
'            For i = 0 To 4
'                Call UniCode_Conv(P_KANRIREC02.KEIHI(i).KOTEI, "")
'            Next i
'
'            Call UniCode_Conv(P_KANRIREC02.FILLER, "")
'
'        Case Else
'            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "管理マスタ")
'            Exit Function
'    End Select
'
'
'    Text1(ptxRec_No).Text = P_ST_KANRI_DEF_No
'
'
'
'
'    j = 0
'    For i = ptxBEF_KOTEI01 To ptxBEF_KOTEI10      '前工程
'
'        If IsNumeric(StrConv(P_KANRIREC02.BEF_KOTEI(j).KOTEI, vbUnicode)) Then
'            Text1(i).Text = Format(CInt(StrConv(P_KANRIREC02.BEF_KOTEI(j).KOTEI, vbUnicode)), "#")
'        Else
'            Text1(i).Text = ""
'        End If
'
'        j = j + 1
'
'    Next i
'
'    j = 0
'    For i = ptxMAIN_KOTEI01 To ptxMAIN_KOTEI10  '作業工程
'
'        If IsNumeric(StrConv(P_KANRIREC02.MAIN_KOTEI(j).KOTEI, vbUnicode)) Then
'            Text1(i).Text = Format(CInt(StrConv(P_KANRIREC02.MAIN_KOTEI(j).KOTEI, vbUnicode)), "#")
'        Else
'            Text1(i).Text = ""
'        End If
'
'        j = j + 1
'
'    Next i
'
'    j = 0
'    For i = ptxAFT_KOTEI01 To ptxAFT_KOTEI10    '後工程
'
'        If IsNumeric(StrConv(P_KANRIREC02.AFT_KOTEI(j).KOTEI, vbUnicode)) Then
'            Text1(i).Text = Format(CInt(StrConv(P_KANRIREC02.AFT_KOTEI(j).KOTEI, vbUnicode)), "#")
'        Else
'            Text1(i).Text = ""
'        End If
'
'        j = j + 1
'
'    Next i
'
'
'
'
'    Item_Disp_Proc = False
'
'End Function

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   管理マスタ出力
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim com     As Integer

Dim i       As Integer
Dim j       As Integer


    Update_Proc = True
    '管理ﾏｽﾀ（KEY=0）読み込み
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_DEF_No)
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_KANRI_POS, P_KANRIREC02, Len(P_KANRIREC02), K0_P_KANRI, Len(K0_P_KANRI), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                
                
                For i = 0 To 9
                    Call UniCode_Conv(P_KANRIREC02.BEF_KOTEI(i).KOTEI, "")
                Next i
            
                For i = 0 To 9
                    Call UniCode_Conv(P_KANRIREC02.MAIN_KOTEI(i).KOTEI, "")
                Next i
            
                For i = 0 To 9
                    Call UniCode_Conv(P_KANRIREC02.AFT_KOTEI(i).KOTEI, "")
                Next i
            
                For i = 0 To 4
                    Call UniCode_Conv(P_KANRIREC02.FUTAI_KOTEI(i).KOTEI, "")
                Next i
            
                For i = 0 To 4
                    Call UniCode_Conv(P_KANRIREC02.KEIHI(i).KOTEI, "")
                Next i
            
                Call UniCode_Conv(P_KANRIREC02.FILLER, "")
                
                
                
                
                
                
                
                
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Update_Proc = True
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "管理マスタ")
                Exit Function
        
        End Select
    
    
    Loop
    '--------------------------------------------------------レコード内容編集
    
    Call UniCode_Conv(P_KANRIREC02.REC_NO, P_ST_KANRI_DEF_No)
    
    j = 0
    For i = ptxBEF_KOTEI01 To ptxBEF_KOTEI10
    
        If IsNumeric(Text1(i).Text) Then
        
            Call UniCode_Conv(P_KANRIREC02.BEF_KOTEI(j).KOTEI, Format(CInt(Text1(i).Text), "000"))
        
        Else
            
            Call UniCode_Conv(P_KANRIREC02.BEF_KOTEI(j).KOTEI, "")
        
        End If
    
        j = j + 1
    
    Next i
    
    j = 0
    For i = ptxMAIN_KOTEI01 To ptxMAIN_KOTEI10
    
        If IsNumeric(Text1(i).Text) Then
        
            Call UniCode_Conv(P_KANRIREC02.MAIN_KOTEI(j).KOTEI, Format(CInt(Text1(i).Text), "000"))
        
        Else
            
            Call UniCode_Conv(P_KANRIREC02.MAIN_KOTEI(j).KOTEI, "")
        
        End If
    
        j = j + 1
    
    Next i
    
    j = 0
    For i = ptxAFT_KOTEI01 To ptxAFT_KOTEI10
    
        If IsNumeric(Text1(i).Text) Then
        
            Call UniCode_Conv(P_KANRIREC02.AFT_KOTEI(j).KOTEI, Format(CInt(Text1(i).Text), "000"))
        
        Else
            
            Call UniCode_Conv(P_KANRIREC02.AFT_KOTEI(j).KOTEI, "")
        
        End If
    
        j = j + 1
    
    Next i
    
    
    
    
    
    
    
    
    
    
    
    
    Do
        
        DoEvents
        
        sts = BTRV(com, P_KANRI_POS, P_KANRIREC02, Len(P_KANRIREC02), K0_P_KANRI, Len(K0_P_KANRI), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, com, "管理マスタ")
                Exit Function
        End Select
    Loop

    j = 2
    For i = 1 To 10
        If Trim(Label(j).Caption) = "" Then
        
            If WriteIni("KOUTEI", "BEF" & Format(i, "00"), "SEI_SYS", "*") Then
                MsgBox "INIﾌｧｲﾙ書き込み異常"
                Exit Function
            End If
        Else
        
        
            If WriteIni("KOUTEI", "BEF" & Format(i, "00"), "SEI_SYS", Trim(Label(j).Caption) & "," & Trim(Text1(j - 1).Text)) Then
                MsgBox "INIﾌｧｲﾙ書き込み異常"
                Exit Function
            End If
        
        End If
        j = j + 1
    Next i

    j = 13
    For i = 1 To 10
        
        If Trim(Trim(Label(j).Caption)) = "" Then
        
            If WriteIni("KOUTEI", "MAIN" & Format(i, "00"), "SEI_SYS", "*") Then
                MsgBox "INIﾌｧｲﾙ書き込み異常"
                Exit Function
            End If
        
        
        
        Else
        
            If WriteIni("KOUTEI", "MAIN" & Format(i, "00"), "SEI_SYS", Trim(Label(j).Caption) & "," & Trim(Text1(j - 2).Text)) Then
                MsgBox "INIﾌｧｲﾙ書き込み異常"
                Exit Function
            End If
        
        End If
        
        j = j + 1
    Next i


    j = 24
    For i = 1 To 10
        
        If Trim(Trim(Label(j).Caption)) = "" Then
        
            If WriteIni("KOUTEI", "AFT" & Format(i, "00"), "SEI_SYS", "*") Then
                MsgBox "INIﾌｧｲﾙ書き込み異常"
                Exit Function
            End If
        
        
        
        Else
        
            If WriteIni("KOUTEI", "AFT" & Format(i, "00"), "SEI_SYS", Trim(Label(j).Caption) & "," & Trim(Text1(j - 3).Text)) Then
                MsgBox "INIﾌｧｲﾙ書き込み異常"
                Exit Function
            End If
        
        End If
        
        j = j + 1
    Next i



    Update_Proc = False


End Function


Private Sub Command1_Click(Index As Integer)

Dim ans As Integer
Dim i   As Integer

    Select Case Index
        Case P_CMD_Upd        '更新
            
            
            For i = ptxRec_No To ptxAFT_KOTEI10
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
            
            Beep
            ans = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            End If
            Text1(ptxBEF_KOTEI01).SetFocus
        Case P_CMD_DEL                      '削除
        Case P_CMD_DSP                      '検索/表示
        
        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        Case P_CMD_PRT                      '印刷
        
        Case P_CMD_End                      '終了
            Me.Visible = False
    End Select

End Sub

Private Sub Form_Activate()
                                '画面初期設定
Dim c       As String * 128
Dim i       As Integer
Dim j       As Integer

Dim KOUTEI  As Variant


                                'ログファイル名取り込み
'前工程セット
    j = 1
    For i = 1 To 10
        j = j + 1
        Label(j).Caption = ""
        If GetIni("KOUTEI", "BEF" & Format(i, "00"), "SEI_SYS", c) Then
        Else
            If Trim(c) = "*" Then
            Else
                KOUTEI = Split(Trim(c), ",", -1)
                Label(j).Caption = Trim(KOUTEI(0))
                Text1(j - 1).Text = Trim(KOUTEI(1))
                Text1(j - 1).Locked = False
                Text1(j - 1).TabStop = True
                            
            End If
        End If
    Next i

'主工程セット
    j = 12
    For i = 1 To 10
        j = j + 1
        Label(j).Caption = ""
        If GetIni("KOUTEI", "MAIN" & Format(i, "00"), "SEI_SYS", c) Then
        Else
            If Trim(c) = "*" Then
            Else
                KOUTEI = Split(Trim(c), ",", -1)
                Label(j).Caption = Trim(KOUTEI(0))
                Text1(j - 2).Text = Trim(KOUTEI(1))
                Text1(j - 2).Locked = False
                Text1(j - 2).TabStop = True
            End If
        End If
    Next i

'後工程セット
    j = 23
    For i = 1 To 10
        j = j + 1
        Label(j).Caption = ""
        If GetIni("KOUTEI", "AFT" & Format(i, "00"), "SEI_SYS", c) Then
        Else
            If Trim(c) = "*" Then
            Else
                KOUTEI = Split(Trim(c), ",", -1)
                Label(j).Caption = Trim(KOUTEI(0))
                Text1(j - 3).Text = Trim(KOUTEI(1))
                Text1(j - 3).Locked = False
                Text1(j - 3).TabStop = True
            End If
        End If
    Next i
                                

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
    PM000103.Caption = PM000103.Caption & LAST_UPDATE_DAY

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    If Error_Check_Proc(Index) Then     'エラーチェック
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動
End Sub

