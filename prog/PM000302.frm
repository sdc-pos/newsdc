VERSION 5.00
Begin VB.Form PM000302 
   Caption         =   "資材マスタメンテナンス"
   ClientHeight    =   11205
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   16320
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
   ScaleHeight     =   11205
   ScaleWidth      =   16320
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   14880
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   169
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "画面印刷"
      Height          =   495
      Left            =   14760
      TabIndex        =   168
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   63
      Left            =   11280
      TabIndex        =   16
      Top             =   1680
      Width           =   4935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "在庫数非表示(棚卸表)"
      Height          =   255
      Index           =   3
      Left            =   7920
      TabIndex        =   15
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   62
      Left            =   5130
      MaxLength       =   1
      TabIndex        =   36
      Top             =   3900
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   61
      Left            =   12600
      MaxLength       =   4
      TabIndex        =   24
      Top             =   2520
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   60
      Left            =   3150
      MaxLength       =   2
      TabIndex        =   35
      Top             =   3900
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   59
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   34
      Top             =   3900
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   58
      Left            =   1890
      MaxLength       =   2
      TabIndex        =   33
      Top             =   3900
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   57
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   32
      Top             =   3900
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   56
      Left            =   7560
      MaxLength       =   2
      TabIndex        =   31
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   55
      Left            =   7560
      MaxLength       =   9
      TabIndex        =   27
      Top             =   3000
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   23
      Left            =   5175
      MaxLength       =   8
      TabIndex        =   30
      Top             =   3480
      Width           =   510
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   22
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   29
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   15
      Left            =   8505
      MaxLength       =   10
      TabIndex        =   19
      Top             =   2160
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   14
      Left            =   4635
      MaxLength       =   20
      TabIndex        =   18
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   13
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   17
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   21
      Left            =   11250
      MaxLength       =   2
      TabIndex        =   28
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   20
      Left            =   4635
      MaxLength       =   9
      TabIndex        =   26
      Top             =   3000
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   19
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   25
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   18
      Left            =   4680
      MaxLength       =   8
      TabIndex        =   22
      Top             =   2580
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   17
      Left            =   3285
      MaxLength       =   8
      TabIndex        =   21
      Top             =   2580
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   16
      Left            =   1935
      MaxLength       =   8
      TabIndex        =   20
      Top             =   2580
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   12
      Left            =   10620
      MaxLength       =   5
      TabIndex        =   23
      Top             =   2520
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   54
      Left            =   14025
      MaxLength       =   11
      TabIndex        =   71
      Top             =   8820
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   8
      ItemData        =   "PM000302.frx":0000
      Left            =   8190
      List            =   "PM000302.frx":0002
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   70
      Top             =   8820
      Width           =   4110
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   53
      Left            =   7470
      MaxLength       =   5
      TabIndex        =   69
      Top             =   8820
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   52
      Left            =   4815
      MaxLength       =   8
      TabIndex        =   68
      Top             =   8820
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   51
      Left            =   1800
      MaxLength       =   11
      TabIndex        =   67
      Top             =   8820
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   13185
      MaxLength       =   8
      TabIndex        =   11
      Top             =   720
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "在庫管理対象外"
      Height          =   255
      Index           =   2
      Left            =   5535
      TabIndex        =   14
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "ﾗﾍﾞﾙ貼り計上なし"
      Height          =   255
      Index           =   1
      Left            =   3015
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   4
      ItemData        =   "PM000302.frx":0004
      Left            =   1320
      List            =   "PM000302.frx":0006
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   12
      Top             =   1680
      Width           =   2835
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   0
      Left            =   1440
      Locked          =   -1  'True
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   32
      Left            =   12495
      MaxLength       =   8
      TabIndex        =   46
      Top             =   5340
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   31
      Left            =   12495
      MaxLength       =   11
      TabIndex        =   45
      Top             =   4860
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   40
      Left            =   12495
      MaxLength       =   11
      TabIndex        =   55
      Top             =   6300
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   41
      Left            =   12495
      MaxLength       =   8
      TabIndex        =   56
      Top             =   6780
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   50
      Left            =   12495
      MaxLength       =   8
      TabIndex        =   66
      Top             =   8280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   49
      Left            =   12495
      MaxLength       =   11
      TabIndex        =   65
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   10
      Left            =   9000
      MaxLength       =   11
      TabIndex        =   75
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "組立製品"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   9720
      TabIndex        =   9
      Top             =   840
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   7425
      MaxLength       =   3
      TabIndex        =   7
      Top             =   720
      Width           =   510
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   4320
      MaxLength       =   2
      TabIndex        =   5
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   1305
      MaxLength       =   2
      TabIndex        =   3
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Index           =   1
      Left            =   4080
      MaxLength       =   40
      TabIndex        =   2
      Top             =   120
      Width           =   4965
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   1
      ItemData        =   "PM000302.frx":0008
      Left            =   1665
      List            =   "PM000302.frx":000A
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   4
      Top             =   720
      Width           =   1635
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   2
      Left            =   4680
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   6
      Top             =   720
      Width           =   1635
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   3
      Left            =   7920
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   8
      Top             =   720
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   11520
      MaxLength       =   11
      TabIndex        =   10
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   72
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   3825
      MaxLength       =   11
      TabIndex        =   73
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   9
      Left            =   6615
      MaxLength       =   11
      TabIndex        =   74
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   11
      Left            =   11745
      MaxLength       =   8
      TabIndex        =   76
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   24
      Left            =   1665
      MaxLength       =   5
      TabIndex        =   37
      Top             =   4380
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   5
      ItemData        =   "PM000302.frx":000C
      Left            =   2385
      List            =   "PM000302.frx":000E
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   38
      Top             =   4380
      Width           =   4110
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   25
      Left            =   7815
      MaxLength       =   11
      TabIndex        =   39
      Top             =   4380
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   26
      Left            =   10335
      MaxLength       =   11
      TabIndex        =   40
      Top             =   4380
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   29
      Left            =   7815
      MaxLength       =   8
      TabIndex        =   43
      Top             =   4860
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   30
      Left            =   10335
      MaxLength       =   3
      TabIndex        =   44
      Top             =   4860
      Width           =   510
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   27
      Left            =   12495
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   28
      Left            =   13800
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   4380
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   33
      Left            =   1665
      MaxLength       =   5
      TabIndex        =   47
      Top             =   5820
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   6
      ItemData        =   "PM000302.frx":0010
      Left            =   2385
      List            =   "PM000302.frx":0012
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   48
      Top             =   5820
      Width           =   4110
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   34
      Left            =   7815
      MaxLength       =   11
      TabIndex        =   49
      Top             =   5820
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   35
      Left            =   10335
      MaxLength       =   11
      TabIndex        =   50
      Top             =   5820
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   38
      Left            =   7815
      MaxLength       =   8
      TabIndex        =   53
      Top             =   6300
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   39
      Left            =   10335
      MaxLength       =   3
      TabIndex        =   54
      Top             =   6300
      Width           =   510
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   36
      Left            =   12450
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   5820
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   37
      Left            =   13800
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   5820
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   42
      Left            =   1665
      MaxLength       =   5
      TabIndex        =   57
      Top             =   7320
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   7
      ItemData        =   "PM000302.frx":0014
      Left            =   2385
      List            =   "PM000302.frx":0016
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   58
      Top             =   7320
      Width           =   4110
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   43
      Left            =   7815
      MaxLength       =   11
      TabIndex        =   59
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   44
      Left            =   10335
      MaxLength       =   11
      TabIndex        =   60
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   47
      Left            =   7815
      MaxLength       =   8
      TabIndex        =   63
      Top             =   7800
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   48
      Left            =   10335
      MaxLength       =   3
      TabIndex        =   64
      Top             =   7800
      Width           =   510
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   45
      Left            =   12495
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   46
      Left            =   13800
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   7320
      Width           =   960
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
      Left            =   10305
      TabIndex        =   88
      Top             =   9960
      Width           =   870
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
      Left            =   9495
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   9960
      Width           =   870
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
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   9960
      Width           =   870
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
      Left            =   7785
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   9960
      Width           =   870
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
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   9960
      Width           =   870
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
      Left            =   5625
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   9960
      Width           =   870
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
      Left            =   4815
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   9960
      Width           =   870
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
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   9960
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "削 除"
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
      Left            =   2655
      TabIndex        =   80
      Top             =   9960
      Width           =   870
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
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   9960
      Width           =   870
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
      Left            =   945
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   9960
      Width           =   870
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
      Left            =   135
      TabIndex        =   77
      Top             =   9960
      Width           =   870
   End
   Begin VB.Label Label 
      Caption         =   "備考"
      Height          =   255
      Index           =   75
      Left            =   10680
      TabIndex        =   167
      Top             =   1800
      Width           =   525
   End
   Begin VB.Label lblUpd_DateTime 
      Caption         =   "99999 99999999-999999"
      Height          =   315
      Left            =   8145
      TabIndex        =   166
      Top             =   9420
      Width           =   2670
   End
   Begin VB.Label Label 
      Caption         =   "更新ＩＤ／日時："
      Height          =   255
      Index           =   74
      Left            =   6075
      TabIndex        =   165
      Top             =   9420
      Width           =   2040
   End
   Begin VB.Label lblIns_DateTime 
      Caption         =   "99999 99999999-999999"
      Height          =   315
      Left            =   2385
      TabIndex        =   164
      Top             =   9360
      Width           =   2670
   End
   Begin VB.Label Label 
      Caption         =   "登録ＩＤ／日時："
      Height          =   255
      Index           =   73
      Left            =   315
      TabIndex        =   163
      Top             =   9360
      Width           =   2040
   End
   Begin VB.Label Label 
      Caption         =   "空白 or 0：集合梱包／１：単体梱包"
      Height          =   255
      Index           =   72
      Left            =   5625
      TabIndex        =   162
      Top             =   4020
      Width           =   4155
   End
   Begin VB.Label Label 
      Caption         =   "梱包区分"
      Height          =   255
      Index           =   71
      Left            =   3960
      TabIndex        =   161
      Top             =   4020
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "口数"
      Height          =   255
      Index           =   70
      Left            =   11970
      TabIndex        =   160
      Top             =   2640
      Width           =   555
   End
   Begin VB.Label Label 
      Caption         =   "m3"
      Height          =   255
      Index           =   69
      Left            =   7245
      TabIndex        =   159
      Top             =   2700
      Width           =   240
   End
   Begin VB.Label lblSIZE 
      Alignment       =   1  '右揃え
      Height          =   315
      Left            =   6075
      TabIndex        =   158
      Top             =   2700
      Width           =   1140
   End
   Begin VB.Label Label 
      Caption         =   "="
      Height          =   255
      Index           =   68
      Left            =   5850
      TabIndex        =   157
      Top             =   2700
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "－"
      Height          =   255
      Index           =   67
      Left            =   2925
      TabIndex        =   156
      Top             =   4020
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "－"
      Height          =   255
      Index           =   66
      Left            =   2295
      TabIndex        =   155
      Top             =   4020
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "－"
      Height          =   255
      Index           =   65
      Left            =   1665
      TabIndex        =   154
      Top             =   4020
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "標準棚番"
      Height          =   255
      Index           =   64
      Left            =   240
      TabIndex        =   153
      Top             =   4020
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   $"PM000302.frx":0018
      Height          =   255
      Index           =   63
      Left            =   7965
      TabIndex        =   152
      Top             =   3600
      Width           =   4155
   End
   Begin VB.Label Label 
      Caption         =   "商品化箱代"
      Height          =   255
      Index           =   62
      Left            =   6210
      TabIndex        =   151
      Top             =   3600
      Width           =   1320
   End
   Begin VB.Label Label 
      Caption         =   "秒"
      Height          =   255
      Index           =   61
      Left            =   8820
      TabIndex        =   150
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label 
      Caption         =   "集合梱包"
      Height          =   255
      Index           =   60
      Left            =   6525
      TabIndex        =   149
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "秒"
      Height          =   255
      Index           =   59
      Left            =   5895
      TabIndex        =   148
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label 
      Caption         =   "（空白：出力しない）"
      Height          =   255
      Index           =   58
      Left            =   11655
      TabIndex        =   147
      Top             =   3120
      Width           =   2445
   End
   Begin VB.Label Label 
      Caption         =   "ﾃｰﾌﾟ長"
      Height          =   255
      Index           =   57
      Left            =   4005
      TabIndex        =   146
      Top             =   3600
      Width           =   780
   End
   Begin VB.Label Label 
      Caption         =   "ﾃｰﾌﾟ種類"
      Height          =   255
      Index           =   56
      Left            =   240
      TabIndex        =   145
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "(1:あり)"
      Height          =   255
      Index           =   55
      Left            =   1800
      TabIndex        =   144
      Top             =   3120
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "印刷"
      Height          =   255
      Index           =   54
      Left            =   720
      TabIndex        =   143
      Top             =   3120
      Width           =   555
   End
   Begin VB.Label Label 
      Caption         =   "材質"
      Height          =   255
      Index           =   53
      Left            =   4005
      TabIndex        =   142
      Top             =   2280
      Width           =   555
   End
   Begin VB.Label Label 
      Caption         =   "形式"
      Height          =   255
      Index           =   52
      Left            =   720
      TabIndex        =   141
      Top             =   2280
      Width           =   555
   End
   Begin VB.Label Label 
      Caption         =   "輸送箱Ｆ"
      Height          =   255
      Index           =   51
      Left            =   10080
      TabIndex        =   140
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "作業工数"
      Height          =   255
      Index           =   46
      Left            =   3465
      TabIndex        =   139
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "強度／厚み"
      Height          =   255
      Index           =   50
      Left            =   7245
      TabIndex        =   138
      Top             =   2280
      Width           =   1320
   End
   Begin VB.Label Label 
      Caption         =   "X"
      Height          =   255
      Index           =   49
      Left            =   4455
      TabIndex        =   137
      Top             =   2700
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "X"
      Height          =   255
      Index           =   48
      Left            =   3105
      TabIndex        =   136
      Top             =   2700
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "才数"
      Height          =   255
      Index           =   45
      Left            =   9990
      TabIndex        =   134
      Top             =   2640
      Width           =   555
   End
   Begin VB.Label Label 
      Caption         =   "最新仕入単価"
      Height          =   255
      Index           =   44
      Left            =   12450
      TabIndex        =   133
      Top             =   8940
      Width           =   1635
   End
   Begin VB.Label Label 
      Caption         =   "最新仕入先"
      Height          =   255
      Index           =   43
      Left            =   6120
      TabIndex        =   132
      Top             =   8940
      Width           =   1365
   End
   Begin VB.Label Label 
      Caption         =   "数量"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   42
      Left            =   13815
      TabIndex        =   131
      Top             =   480
      Width           =   510
   End
   Begin VB.Label Label 
      Caption         =   "最終出庫数"
      Height          =   255
      Index           =   41
      Left            =   3510
      TabIndex        =   130
      Top             =   8940
      Width           =   1365
   End
   Begin VB.Label Label 
      Caption         =   "最終出荷日"
      Height          =   255
      Index           =   40
      Left            =   360
      TabIndex        =   129
      Top             =   8940
      Width           =   1365
   End
   Begin VB.Label Label 
      Caption         =   "/"
      Height          =   255
      Index           =   39
      Left            =   13095
      TabIndex        =   128
      Top             =   840
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "資材区分"
      Height          =   255
      Index           =   38
      Left            =   240
      TabIndex        =   127
      Top             =   1800
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "前回注文日"
      Height          =   255
      Index           =   37
      Left            =   11145
      TabIndex        =   126
      Top             =   4980
      Width           =   1230
   End
   Begin VB.Label Label 
      Caption         =   "注文数"
      Height          =   255
      Index           =   36
      Left            =   11640
      TabIndex        =   125
      Top             =   5460
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "注文数"
      Height          =   255
      Index           =   35
      Left            =   11640
      TabIndex        =   124
      Top             =   6900
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "前回注文日"
      Height          =   255
      Index           =   34
      Left            =   11145
      TabIndex        =   123
      Top             =   6420
      Width           =   1230
   End
   Begin VB.Label Label 
      Caption         =   "注文数"
      Height          =   255
      Index           =   33
      Left            =   11640
      TabIndex        =   122
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "前回注文日"
      Height          =   255
      Index           =   32
      Left            =   11145
      TabIndex        =   121
      Top             =   7920
      Width           =   1230
   End
   Begin VB.Label Label 
      Caption         =   "設定日"
      Height          =   255
      Index           =   31
      Left            =   8145
      TabIndex        =   120
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "設定日"
      Height          =   255
      Index           =   7
      Left            =   3015
      TabIndex        =   119
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "%"
      Height          =   255
      Index           =   30
      Left            =   14835
      TabIndex        =   118
      Top             =   7680
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "%"
      Height          =   255
      Index           =   23
      Left            =   14835
      TabIndex        =   117
      Top             =   6120
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "%"
      Height          =   255
      Index           =   16
      Left            =   14835
      TabIndex        =   116
      Top             =   4560
      Width           =   150
   End
   Begin VB.Label Label 
      Caption         =   "資材品番"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   115
      Top             =   240
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "品名"
      Height          =   255
      Index           =   2
      Left            =   2745
      TabIndex        =   114
      Top             =   240
      Width           =   510
   End
   Begin VB.Label Label 
      Caption         =   "仕入区分"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   113
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "販売区分"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3375
      TabIndex        =   112
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "収支単位"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   6345
      TabIndex        =   111
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "前月在庫金額"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   11385
      TabIndex        =   110
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "標準売価"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   109
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "標準原価"
      Height          =   255
      Index           =   8
      Left            =   5400
      TabIndex        =   108
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "危険在庫"
      Height          =   255
      Index           =   9
      Left            =   10665
      TabIndex        =   107
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "仕入先(1)"
      Height          =   255
      Index           =   10
      Left            =   495
      TabIndex        =   106
      Top             =   4500
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "仕入単価"
      Height          =   255
      Index           =   11
      Left            =   6735
      TabIndex        =   105
      Top             =   4500
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "設定日"
      Height          =   255
      Index           =   12
      Left            =   9480
      TabIndex        =   104
      Top             =   4500
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "ﾛｯﾄ数"
      Height          =   255
      Index           =   13
      Left            =   7095
      TabIndex        =   103
      Top             =   4980
      Width           =   645
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   360
      X2              =   14280
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label 
      Caption         =   "ﾘｰﾄﾞﾀｲﾑ"
      Height          =   255
      Index           =   14
      Left            =   9345
      TabIndex        =   102
      Top             =   4980
      Width           =   870
   End
   Begin VB.Label Label 
      Caption         =   "粗利"
      Height          =   255
      Index           =   15
      Left            =   11865
      TabIndex        =   101
      Top             =   4500
      Width           =   510
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   360
      X2              =   14280
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label 
      Caption         =   "仕入先(2)"
      Height          =   255
      Index           =   17
      Left            =   495
      TabIndex        =   100
      Top             =   5940
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "仕入単価"
      Height          =   255
      Index           =   18
      Left            =   6735
      TabIndex        =   99
      Top             =   5940
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "設定日"
      Height          =   255
      Index           =   19
      Left            =   9480
      TabIndex        =   98
      Top             =   5940
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "ﾛｯﾄ数"
      Height          =   255
      Index           =   20
      Left            =   7095
      TabIndex        =   97
      Top             =   6420
      Width           =   645
   End
   Begin VB.Label Label 
      Caption         =   "ﾘｰﾄﾞﾀｲﾑ"
      Height          =   255
      Index           =   21
      Left            =   9345
      TabIndex        =   96
      Top             =   6420
      Width           =   870
   End
   Begin VB.Label Label 
      Caption         =   "粗利"
      Height          =   255
      Index           =   22
      Left            =   11865
      TabIndex        =   95
      Top             =   5940
      Width           =   510
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   360
      X2              =   14160
      Y1              =   7260
      Y2              =   7260
   End
   Begin VB.Label Label 
      Caption         =   "仕入先(3)"
      Height          =   255
      Index           =   24
      Left            =   495
      TabIndex        =   94
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "仕入単価"
      Height          =   255
      Index           =   25
      Left            =   6735
      TabIndex        =   93
      Top             =   7440
      Width           =   1005
   End
   Begin VB.Label Label 
      Caption         =   "設定日"
      Height          =   255
      Index           =   26
      Left            =   9480
      TabIndex        =   92
      Top             =   7440
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "ﾛｯﾄ数"
      Height          =   255
      Index           =   27
      Left            =   7095
      TabIndex        =   91
      Top             =   7920
      Width           =   645
   End
   Begin VB.Label Label 
      Caption         =   "ﾘｰﾄﾞﾀｲﾑ"
      Height          =   255
      Index           =   28
      Left            =   9345
      TabIndex        =   90
      Top             =   7920
      Width           =   870
   End
   Begin VB.Label Label 
      Caption         =   "粗利"
      Height          =   255
      Index           =   29
      Left            =   11865
      TabIndex        =   89
      Top             =   7440
      Width           =   510
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   630
      X2              =   14310
      Y1              =   8700
      Y2              =   8700
   End
   Begin VB.Label Label 
      Caption         =   "ｻｲｽﾞ(WxDxH)mm"
      Height          =   255
      Index           =   47
      Left            =   45
      TabIndex        =   135
      Top             =   2700
      Width           =   1635
   End
End
Attribute VB_Name = "PM000302"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'テキスト用添字
Private Const ptxHIN_GAI% = 0               '品番
Private Const ptxHIN_NAME% = 1              '品名
Private Const ptxG_SHIIRE_KBN% = 2          '仕入区分
Private Const ptxG_HANBAI_KBN% = 3          '販売区分
Private Const ptxG_SYUSHI% = 4              '収支単位
Private Const ptxG_ZEN_ZAIKO_KIN% = 5       '前月在庫金額
Private Const ptxG_ZEN_ZAIKO_QTY% = 6       '前月在庫数量
Private Const ptxG_ST_URITAN% = 7           '標準粗利売価単価
Private Const ptxG_ST_URITAN_DT% = 8        '標準粗利売価単価設定日
Private Const ptxG_ST_SHITAN% = 9           '標準粗利原価単価
Private Const ptxG_ST_SHITAN_DT% = 10       '標準粗利原価単価設定日
Private Const ptxHOJYU_P% = 11              '補充点（危険在庫）

Private Const ptxSAI_SU% = 12               '才数               2008.02.14


Private Const ptxD_KEISHIKI% = 13           '形式               2008.02.14
Private Const ptxD_MATERIAL% = 14           '材質               2008.02.14
Private Const ptxD_THICKNESS% = 15          'ﾀﾞﾝﾎﾞｰﾙ　厚さ      2008.02.14
    
    
Private Const ptxD_SIZE_W% = 16             'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（W）   2008.02.14
Private Const ptxD_SIZE_D% = 17             'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（D）   2008.02.14
Private Const ptxD_SIZE_H% = 18             'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（H）   2008.02.14
        
Private Const ptxD_PRINT% = 19              '印刷する／しない   2008.02.14
            
        
Private Const ptxS_KOUSU% = 20              '商品化　工数       2008.02.14


Private Const ptxSE_USOU_F% = 21            '輸送箱出力      2008.02.14


    
Private Const ptxUSE_TAPE_KIND% = 22        '使用テープ種類     2008.02.14
Private Const ptxUSE_TAPE_LNG% = 23         '使用テープ長     2008.02.14



Private Const ptxSHI_CODE1% = 24            '仕入先ｺｰﾄﾞ(1)
Private Const ptxSHI_TANKA1% = 25           '仕入単価(1)
Private Const ptxSHI_TANKA_DT1% = 26        '仕入単価設定日(1)
Private Const ptxSHI_ARARI1% = 27           '粗利額(1)
Private Const ptxSHI_ARARI_RITU1% = 28      '粗利率(1)
Private Const ptxSHI_LOT1% = 29             'ﾛｯﾄ数(1)
Private Const ptxSHI_LEAD_TIME1% = 30       'ﾘｰﾄﾞﾀｲﾑ(1)
Private Const ptxSHI_LAST_ORDER_DT1% = 31   '前回注文日(1)
Private Const ptxSHI_LAST_ORDER_QTY1% = 32  '前回注文数(1)

Private Const ptxSHI_CODE2% = 33            '仕入先ｺｰﾄﾞ(2)
Private Const ptxSHI_TANKA2% = 34           '仕入単価(2)
Private Const ptxSHI_TANKA_DT2% = 35        '仕入単価設定日(2)
Private Const ptxSHI_ARARI2% = 36           '粗利額(2)
Private Const ptxSHI_ARARI_RITU2% = 37      '粗利率(2)
Private Const ptxSHI_LOT2% = 38             'ﾛｯﾄ数(2)
Private Const ptxSHI_LEAD_TIME2% = 39       'ﾘｰﾄﾞﾀｲﾑ(2)
Private Const ptxSHI_LAST_ORDER_DT2% = 40   '前回注文日(2)
Private Const ptxSHI_LAST_ORDER_QTY2% = 41  '前回注文数(2)

Private Const ptxSHI_CODE3% = 42            '仕入先ｺｰﾄﾞ(3)
Private Const ptxSHI_TANKA3% = 43           '仕入単価(3)
Private Const ptxSHI_TANKA_DT3% = 44        '仕入単価設定日(3)
Private Const ptxSHI_ARARI3% = 45           '粗利額(3)
Private Const ptxSHI_ARARI_RITU3% = 46      '粗利率(3)
Private Const ptxSHI_LOT3% = 47             'ﾛｯﾄ数(3)
Private Const ptxSHI_LEAD_TIME3% = 48       'ﾘｰﾄﾞﾀｲﾑ(3)
Private Const ptxSHI_LAST_ORDER_DT3% = 49   '前回注文日(3)
Private Const ptxSHI_LAST_ORDER_QTY3% = 50  '前回注文数(3)

Private Const ptxLAST_SYU_DT = 51           '最終出庫日
Private Const ptxG_LAST_SYUKA_QTY = 52      '最終出庫数

Private Const ptxLAST_CODE = 53             '最新仕入先コード   2007.05.28
Private Const ptxLAST_TANKA = 54            '最新仕入単価       2007.05.28

Private Const ptxSEI_SYU_KON = 55           '集合梱包           2008.07.16
Private Const ptxSEI_KBN = 56               '請求区分           2008.07.16

Private Const ptxST_SOKO% = 57              '標準棚番　倉庫     2009.09.01
Private Const ptxST_RETU% = 58              '標準棚番　列       2009.09.01
Private Const ptxST_REN% = 59               '標準棚番　連       2009.09.01
Private Const ptxST_DAN% = 60               '標準棚番　段       2009.09.01

Private Const ptxKUTI_SU% = 61              '口数               2010.01.18
Private Const ptxKONPOU_F% = 62             '梱包区分           2010.01.18

Private Const ptxSHIIRE_BIKOU% = 63        '仕入備考           2018.04.19


'コンボ用添字
Private Const pcmbNAIGAI% = 0               '国内外
Private Const pcmbG_SHIIRE% = 1             '仕入区分
Private Const pcmbG_HANBAI% = 2             '販売区分
Private Const pcmbG_SYUSHI% = 3             '収支単位
Private Const pcmbG_SHIZAI_KBN% = 4         '資材区分
Private Const pcmbSHIIRE1% = 5              '仕入先(1)
Private Const pcmbSHIIRE2% = 6              '仕入先(2)
Private Const pcmbSHIIRE3% = 7              '仕入先(3)
Private Const pcmbLAST_CODE% = 8            '最新仕入先         2007.05.28
'チェック用添字
Private Const pchkG_KUMITATE% = 0           '組立製品
Private Const pchkG_LABEL_NON% = 1          'ﾗﾍﾞﾙ貼り計上なし
Private Const pchkZAIKO_F% = 2              '在庫管理対象外

Private Const pchkZAIKO_CLR_F% = 3          '棚卸表　在庫数非表示   2012.12.13

Private INIT_FLG    As Boolean

Private svTANKA     As String               '2018.04.09

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PM000302.MousePointer = vbHourglass

    Call Ctrl_Lock(PM000302)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM000302)


    PM000302.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
Dim com     As Integer
Dim ans     As Integer
Dim sts     As Integer

Dim i       As Integer
    
Dim w_CODE  As String * 22      '2018.04.19
    
    
    Error_Check_Proc = True
    
    
    
    
    Select Case Mode
        
        Case ptxHIN_GAI      '品番
            
            If Trim(Text1(ptxHIN_GAI).Text) = "" Then
                MsgBox "入力した項目はエラーです。(品番)"
                Text1(ptxHIN_GAI).SetFocus
                Exit Function
            End If
            
        
            If G_SCREEN_FLG = G_SCREEN_INS Then
                '新規時は重複チェック
                Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI), 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
            
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        ans = MsgBox("入力したコードは、登録済です。更新処理として継続しますか？", vbYesNo, "確認入力")
                        If ans = vbNo Then
                            Text1(ptxHIN_GAI).SetFocus
                            Exit Function
                        End If
                                    
                        w_CODE = Last_JGYOBU & Right(Combo1(pcmbNAIGAI), 1) & Text1(ptxHIN_GAI).Text                '2018.04.19
'                        Call Item_Disp_Proc(Last_JGYOBU & Right(Combo1(pcmbNAIGAI), 1) & Text1(ptxHIN_GAI).Text)   '2018.04.19
                        Call Item_Disp_Proc(w_CODE)                                                                 '2018.04.19
                    
                    Case BtErrKeyNotFound
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                End Select
            
            
                Text1(ptxHIN_GAI).BackColor = G_INPUT_NG
                Text1(ptxHIN_GAI).Locked = True
                Text1(ptxHIN_GAI).TabStop = False
            
'                Text1(ptxHIN_NAME).BackColor = G_INPUT_NG
'                Text1(ptxHIN_NAME).Locked = True
'                Text1(ptxHIN_NAME).TabStop = False
            
            End If
        
        Case ptxG_SHIIRE_KBN       '仕入区分
        
    
                If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                    Error_Check_Proc = False        '2016.05.18
                    Exit Function                   '2016.05.18
                End If                              '2016.05.18
            
'2015.09.16            If Last_JGYOBU = SHIZAI Then
                If Trim(Text1(ptxG_SHIIRE_KBN).Text) = "" Then
'2016.06.18                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(仕入区分)"
                    Text1(ptxG_SHIIRE_KBN).SetFocus
                    Exit Function
                End If
            
            
                For i = 0 To Combo1(pcmbG_SHIIRE).ListCount - 1
                    If Text1(ptxG_SHIIRE_KBN).Text = Left(Right(Combo1(pcmbG_SHIIRE).List(i), 3), 2) Then
                        Combo1(pcmbG_SHIIRE).ListIndex = i
                        Exit For
                    End If
                Next i
                
                If i > (Combo1(pcmbG_SHIIRE).ListCount - 1) Then
'2016.05.18                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(仕入区分)"
                    Text1(ptxG_SHIIRE_KBN).SetFocus
                    Exit Function
                End If
                        
'2015.09.16            End If
        Case ptxG_HANBAI_KBN       '販売区分
        

                If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                    Error_Check_Proc = False        '2016.05.18
                    Exit Function                   '2016.05.18
                End If                              '2016.05.18


'2015.09.16            If Last_JGYOBU = SHIZAI Then
                If Trim(Text1(ptxG_HANBAI_KBN).Text) = "" Then
'                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(販売区分)"
                    Text1(ptxG_HANBAI_KBN).SetFocus
                    Exit Function
                End If
            
            
                For i = 0 To Combo1(pcmbG_HANBAI).ListCount - 1
                    If Text1(ptxG_HANBAI_KBN).Text = Left(Right(Combo1(pcmbG_HANBAI).List(i), 3), 2) Then
                        Combo1(pcmbG_HANBAI).ListIndex = i
                        Exit For
                    End If
                Next i
                
                If i > (Combo1(pcmbG_HANBAI).ListCount - 1) Then
'2016.05.18                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(販売区分)"
                    Text1(ptxG_HANBAI_KBN).SetFocus
                    Exit Function
                End If
'2015.09.16            End If
        Case ptxG_SYUSHI           '収支単位
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
        
        
            If Trim(Text1(ptxG_SYUSHI).Text) = "" Then
''                MsgBox "入力した項目はエラーです。"
''                Text1(ptxG_SYUSHI).SetFocus
''                Exit Function
            Else
        
                For i = 0 To Combo1(pcmbG_SYUSHI).ListCount - 1
                    If Text1(ptxG_SYUSHI).Text = Right(Combo1(pcmbG_SYUSHI).List(i), 3) Then
                        Combo1(pcmbG_SYUSHI).ListIndex = i
                        Exit For
                    End If
                Next i
                
                If i > (Combo1(pcmbG_SYUSHI).ListCount - 1) Then
'2016.05.18                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(収支単位)"
                    Text1(ptxG_SYUSHI).SetFocus
                    Exit Function
                End If
            End If
        
        Case ptxG_ZEN_ZAIKO_KIN    '前月在庫金額
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            
            If Trim(Text1(ptxG_ZEN_ZAIKO_KIN).Text) = "" Then
                Text1(ptxG_ZEN_ZAIKO_KIN).Text = "0"
            End If
        
            If Not IsNumeric(Text1(ptxG_ZEN_ZAIKO_KIN).Text) Then
'2016.05.18                MsgBox "入力した項目はエラーです。"
                MsgBox "入力した項目はエラーです。(前月在庫金額)"
                Text1(ptxG_ZEN_ZAIKO_KIN).SetFocus
                Exit Function
            End If
        
            Text1(ptxG_ZEN_ZAIKO_KIN).Text = Format(CLng(Text1(ptxG_ZEN_ZAIKO_KIN).Text), "#0")
        
        Case ptxG_ZEN_ZAIKO_QTY    '前月在庫数量
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            
            If Trim(Text1(ptxG_ZEN_ZAIKO_QTY).Text) = "" Then
                Text1(ptxG_ZEN_ZAIKO_QTY).Text = "0"
            End If
        
            If Not IsNumeric(Text1(ptxG_ZEN_ZAIKO_QTY).Text) Then
'2016.05.18                MsgBox "入力した項目はエラーです。"
                MsgBox "入力した項目はエラーです。(前月在庫数量)"
                Text1(ptxG_ZEN_ZAIKO_QTY).SetFocus
                Exit Function
            End If
        
            Text1(ptxG_ZEN_ZAIKO_QTY).Text = Format(CLng(Text1(ptxG_ZEN_ZAIKO_QTY).Text), "#0")
        
        
        Case ptxG_ST_URITAN        '標準売価
            
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
        
            If Trim(Text1(ptxG_ST_URITAN).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxG_ST_URITAN).Text) Then
'2016.05.18                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(標準売価)"
                    Text1(ptxG_ST_URITAN).SetFocus
                    Exit Function
                End If
            
                Text1(ptxG_ST_URITAN).Text = Format(CDbl(Text1(ptxG_ST_URITAN).Text), "#0.00")
            End If
        
        Case ptxG_ST_URITAN_DT     '標準売価設定日
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            
            If Trim(Text1(ptxG_ST_URITAN_DT).Text) = "" Then
                If Trim(Text1(ptxG_ST_URITAN).Text) <> "" Then
                    Text1(ptxG_ST_URITAN_DT).Text = Format(Now, "YYYY/MM/DD")
                End If
            Else
                If Not IsDate(Text1(ptxG_ST_URITAN_DT).Text) Then
'2016.05.18                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(標準売価設定日)"
                    Text1(ptxG_ST_URITAN_DT).SetFocus
                    Exit Function
                End If
            
                Text1(ptxG_ST_URITAN_DT).Text = Format(Text1(ptxG_ST_URITAN_DT).Text, "YYYY/MM/DD")
            End If
        
        Case ptxG_ST_SHITAN       '標準原価
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            
            
            If Trim(Text1(ptxG_ST_SHITAN).Text) = "" Then
            Else
        
                If Not IsNumeric(Text1(ptxG_ST_SHITAN).Text) Then
'2016.05.18                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(標準原価)"
                    Text1(ptxG_ST_URITAN).SetFocus
                    Exit Function
                End If
            
                Text1(ptxG_ST_SHITAN).Text = Format(CDbl(Text1(ptxG_ST_SHITAN).Text), "#0.00")
            End If
        
        Case ptxG_ST_SHITAN_DT     '標準原価設定日
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            
            
            If Trim(Text1(ptxG_ST_SHITAN_DT).Text) = "" Then
                If Trim(Text1(ptxG_ST_SHITAN).Text) <> "" Then
                    Text1(ptxG_ST_SHITAN_DT).Text = Format(Now, "YYYY/MM/DD")
                End If
            Else
                If Not IsDate(Text1(ptxG_ST_SHITAN_DT).Text) Then
'2016.05.18                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(標準原価設定日)"
                    Text1(ptxG_ST_SHITAN_DT).SetFocus
                    Exit Function
                End If
            
                Text1(ptxG_ST_SHITAN_DT).Text = Format(Text1(ptxG_ST_SHITAN_DT).Text, "YYYY/MM/DD")
            End If
        
        
        Case ptxHOJYU_P            '危険在庫
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
        
            If Trim(Text1(ptxHOJYU_P).Text) = "" Then
                Text1(ptxHOJYU_P).Text = "0"
            End If
            
            If Not IsNumeric(Text1(ptxHOJYU_P).Text) Then
'2016.05.18                MsgBox "入力した項目はエラーです。"
                MsgBox "入力した項目はエラーです。(危険在庫)"
                Text1(ptxG_ST_URITAN).SetFocus
                Exit Function
            End If
        
            Text1(ptxHOJYU_P).Text = Format(CLng(Text1(ptxHOJYU_P).Text), "#0")
              
              
              
        Case ptxSAI_SU              '才数   2008.02.14
              
            
            If Trim(Text1(ptxSAI_SU).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSAI_SU).Text) Then
                    MsgBox "入力した項目はエラーです。(才数)"
                    Text1(ptxSAI_SU).SetFocus
                    Exit Function
                Else
                    Text1(ptxSAI_SU).Text = Format(CCur(Text1(ptxSAI_SU).Text), "#0.00")
                End If
            End If
              
              
              
              
              
              
        Case ptxD_KEISHIKI          '形式               2008.02.14
        Case ptxD_MATERIAL          '材質               2008.02.14
        Case ptxD_THICKNESS         'ﾀﾞﾝﾎﾞｰﾙ　厚さ      2008.02.14
    
    
        Case ptxD_SIZE_W            'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（W）   2008.02.14
        
            If Text1(ptxD_SIZE_W).Text = "" Then
            Else
                If Not IsNumeric(Text1(ptxD_SIZE_W).Text) Then
                    MsgBox "入力した項目はエラーです。(ｻｲｽﾞ（W）)"
                    Text1(ptxD_SIZE_W).SetFocus
                    Exit Function
                Else
                    
On Error GoTo Error_Proc
                    
                    Text1(ptxD_SIZE_W).Text = Format(Val(Text1(ptxD_SIZE_W).Text), "#")
                
                
                    If IsNumeric(Text1(ptxD_SIZE_W).Text) And IsNumeric(Text1(ptxD_SIZE_D).Text) And IsNumeric(Text1(ptxD_SIZE_H).Text) Then
                    
                        lblSIZE.Caption = Format(ToHalfAdjust(CCur(Val(Text1(ptxD_SIZE_W).Text) / 1000 * Val(Text1(ptxD_SIZE_D).Text) / 1000 * Val(Text1(ptxD_SIZE_H).Text) / 1000), 4), "#0.000")
                    
                        If Trim(Text1(ptxSAI_SU).Text) = "" Then
                        
'                            Text1(ptxSAI_SU).Text = Format(ToHalfAdjust(CCur(CCur(lblSIZE.Caption) / 0.028), 1), "0.00")       '2016.04.22
                            Text1(ptxSAI_SU).Text = Format(ToHalfAdjust(CCur(CCur(lblSIZE.Caption) / 0.028), 2), "0.00")        '2016.04.22
                        
                        
                        End If
                    
                    End If
                
                
On Error GoTo 0
                
                
                
                End If
            End If
        
        Case ptxD_SIZE_D            'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（D）   2008.02.14
        
            If Text1(ptxD_SIZE_D).Text = "" Then
            Else
                If Not IsNumeric(Text1(ptxD_SIZE_D).Text) Then
                    MsgBox "入力した項目はエラーです。(ｻｲｽﾞ（D）)"
                    Text1(ptxD_SIZE_D).SetFocus
                    Exit Function
                Else
                    
On Error GoTo Error_Proc
                    
                    Text1(ptxD_SIZE_D).Text = Format(Val(Text1(ptxD_SIZE_D).Text), "#")
                
                
                    If IsNumeric(Text1(ptxD_SIZE_W).Text) And IsNumeric(Text1(ptxD_SIZE_D).Text) And IsNumeric(Text1(ptxD_SIZE_H).Text) Then
                    
                        
                        lblSIZE.Caption = Format(ToHalfAdjust(CCur(Val(Text1(ptxD_SIZE_W).Text) / 1000 * Val(Text1(ptxD_SIZE_D).Text) / 1000 * Val(Text1(ptxD_SIZE_H).Text) / 1000), 4), "#0.000")
                    
                        If Trim(Text1(ptxSAI_SU).Text) = "" Then
                        
'                            Text1(ptxSAI_SU).Text = Format(ToHalfAdjust(CCur(CCur(lblSIZE.Caption) / 0.028), 1), "0.00")       '2016.04.22
                            Text1(ptxSAI_SU).Text = Format(ToHalfAdjust(CCur(CCur(lblSIZE.Caption) / 0.028), 2), "0.00")        '2016.04.22
                        
                        
                        End If
                    
                    End If
                
On Error GoTo 0
                
                End If
            
            End If
        
        
        
        
        Case ptxD_SIZE_H            'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（H）   2008.02.14
            
        
            If Text1(ptxD_SIZE_H).Text = "" Then
            Else
                If Not IsNumeric(Text1(ptxD_SIZE_H).Text) Then
                    MsgBox "入力した項目はエラーです。(ｻｲｽﾞ（H）)"
                    Text1(ptxD_SIZE_H).SetFocus
                    Exit Function
                Else
                    
On Error GoTo Error_Proc
                    
                    Text1(ptxD_SIZE_H).Text = Format(Val(Text1(ptxD_SIZE_H).Text), "#")
                
                    If IsNumeric(Text1(ptxD_SIZE_W).Text) And IsNumeric(Text1(ptxD_SIZE_D).Text) And IsNumeric(Text1(ptxD_SIZE_H).Text) Then
                    
                        lblSIZE.Caption = Format(ToHalfAdjust(CCur(Val(Text1(ptxD_SIZE_W).Text) / 1000 * Val(Text1(ptxD_SIZE_D).Text) / 1000 * Val(Text1(ptxD_SIZE_H).Text) / 1000), 4), "#0.000")
                    
                        If Trim(Text1(ptxSAI_SU).Text) = "" Then
                        
'                            Text1(ptxSAI_SU).Text = Format(ToHalfAdjust(CCur(CCur(lblSIZE.Caption) / 0.028), 1), "0.00")       '2016.04.22
                            Text1(ptxSAI_SU).Text = Format(ToHalfAdjust(CCur(CCur(lblSIZE.Caption) / 0.028), 2), "0.00")        '2016.04.22
                        
                        
                        End If
                    
                    End If
On Error GoTo 0
                
                End If
            End If
        
        
        
        Case ptxD_PRINT             '印刷する／しない   2008.02.14
        Case ptxS_KOUSU             '商品化　工数       2008.07.16
            
        
            If Trim(Text1(Mode).Text) = "" Then
            Else
                If IsNumeric(Text1(Mode).Text) Then
                Else
                    MsgBox "入力した項目はエラーです。(作業工数)"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If
        
        Case ptxSEI_SYU_KON         '集合梱包　工数     2008.07.16
            
        
            If Trim(Text1(Mode).Text) = "" Then
            Else
                If IsNumeric(Text1(Mode).Text) Then
                Else
                    MsgBox "入力した項目はエラーです。(集合梱包工数)"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If
        
        
        Case ptxSE_USOU_F           '輸送箱　出力ﾌﾗｸﾞ   2008.02.14
        Case ptxUSE_TAPE_KIND       '使用テープ種類     2008.02.14
        Case ptxUSE_TAPE_LNG        '使用テープ長       2008.02.14
              
        Case ptxSEI_KBN        '請求区分           2008.07.16
            
            
            
            If Trim(Text1(Mode).Text) = "" Or _
                Trim(Text1(Mode).Text) = "1" Or _
                Trim(Text1(Mode).Text) = "2" Then
            Else
                MsgBox "入力した項目はエラーです。(請求区分)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        
        Case ptxST_SOKO      '標準棚番    2009.09.01
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(Val(Text1(Mode).Text), "00")
            End If
        
        Case ptxST_RETU      '標準棚番    2009.09.01
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(Val(Text1(Mode).Text), "00")
            End If
        
        Case ptxST_REN      '標準棚番    2009.09.01
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(Val(Text1(Mode).Text), "00")
            End If
        
        Case ptxST_DAN      '標準棚番    2009.09.01
            If IsNumeric(Text1(Mode).Text) Then
                Text1(Mode).Text = Format(Val(Text1(Mode).Text), "00")
            End If
        
        
            If Trim(Text1(ptxST_SOKO).Text) = "" And Trim(Text1(ptxST_RETU).Text) = "" And Trim(Text1(ptxST_REN).Text) = "" And Trim(Text1(ptxST_DAN).Text) = "" Then
            Else
                If Trim(Text1(ptxST_SOKO).Text) = "**" And Trim(Text1(ptxST_RETU).Text) = "**" And Trim(Text1(ptxST_REN).Text) = "**" And Trim(Text1(ptxST_DAN).Text) = "**" Then
                Else
                    Call UniCode_Conv(K0_TANA.SOKO_NO, Text1(ptxST_SOKO).Text)
                    Call UniCode_Conv(K0_TANA.Retu, Text1(ptxST_RETU).Text)
                    Call UniCode_Conv(K0_TANA.Ren, Text1(ptxST_REN).Text)
                    Call UniCode_Conv(K0_TANA.Dan, Text1(ptxST_DAN).Text)
            
            
                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound
                            MsgBox "入力した項目はエラーです。(標準棚番)"
                            Text1(ptxST_SOKO).SetFocus
                            Exit Function
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "棚マスタ")
                            Exit Function
                    End Select
                End If
        
        
        
        
        
            End If
        
        
        
        Case ptxSHI_CODE1           '仕入先(1)
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            
            Text1(Mode).Text = StrConv(Text1(Mode).Text, vbUpperCase) '2016.01.26
            If Trim(Text1(ptxSHI_CODE1).Text) = "" Then
                For i = ptxSHI_CODE1 To ptxSHI_LAST_ORDER_QTY1
                    Text1(i).Text = ""
                Next i
                Combo1(pcmbSHIIRE1).ListIndex = -1
            Else
                For i = 0 To Combo1(pcmbSHIIRE1).ListCount - 1
                    If Trim(Text1(ptxSHI_CODE1).Text) = Trim(Right(Combo1(pcmbSHIIRE1).List(i), 5)) Then
                        Combo1(pcmbSHIIRE1).ListIndex = i
                        Exit For
                    End If
                Next i
                
                If i > (Combo1(pcmbSHIIRE1).ListCount - 1) Then
'2016.05.18                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(仕入先(1))"
                    Text1(ptxSHI_CODE1).SetFocus
                    Exit Function
                End If
                        
            
            
            
            End If
        
        Case ptxSHI_TANKA1         '仕入単価(1)
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_TANKA1).Text) = "" Then
            
                Text1(ptxSHI_ARARI1).Text = ""
                Text1(ptxSHI_ARARI_RITU1).Text = ""
            
            Else
                
                If Not IsNumeric(Text1(ptxSHI_TANKA1).Text) Then
'2016.05.18                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(仕入単価(1))"
                    Text1(ptxSHI_TANKA1).SetFocus
                    Exit Function
                End If
            
                Text1(ptxSHI_TANKA1).Text = Format(CDbl(Text1(ptxSHI_TANKA1).Text), "#0.00")
            
                If Not IsNumeric(Text1(ptxSHI_TANKA1).Text) Or _
                    Not IsNumeric(Text1(ptxG_ST_URITAN).Text) Then

                    Text1(ptxSHI_ARARI1).Text = ""
                    Text1(ptxSHI_ARARI_RITU1).Text = ""
                Else
                    Text1(ptxSHI_ARARI1).Text = Format(CDbl(Text1(ptxG_ST_URITAN).Text) - CDbl(Text1(ptxSHI_TANKA1).Text), "#0.00")
                    
                    If CDbl(Text1(ptxG_ST_URITAN).Text) = 0 Then
                        Text1(ptxSHI_ARARI_RITU1).Text = "0.00"
                    Else
                        Text1(ptxSHI_ARARI_RITU1).Text = Format(CDbl(Text1(ptxSHI_ARARI1).Text) / CDbl(Text1(ptxG_ST_URITAN).Text) * 100, "#0.00")
                    End If
                End If
            
            End If
                
        Case ptxSHI_TANKA_DT1     '仕入単価設定日(1)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_TANKA_DT1).Text) = "" Then
                If Trim(Text1(ptxSHI_TANKA1).Text) <> "" Then
                    Text1(ptxSHI_TANKA_DT1).Text = Format(Now, "YYYY/MM/DD")
                End If
            Else
                If Not IsDate(Text1(ptxSHI_TANKA_DT1).Text) Then
'2016.05.18                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(仕入単価設定日(1))"
                    Text1(ptxSHI_TANKA_DT1).SetFocus
                    Exit Function
                End If
            
                Text1(ptxSHI_TANKA_DT1).Text = Format(Text1(ptxSHI_TANKA_DT1).Text, "YYYY/MM/DD")
            End If
                
        
        Case ptxSHI_LOT1           'ﾛｯﾄ数(1)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LOT1).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSHI_LOT1).Text) Then
'2016.05.18                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(ﾛｯﾄ数(1))"
                    Text1(ptxSHI_LOT1).SetFocus
                    Exit Function
                End If
        
                Text1(ptxSHI_LOT1).Text = Format(CLng(Text1(ptxSHI_LOT1).Text), "#0")
        
            End If
        
        Case ptxSHI_LEAD_TIME1     'ﾘｰﾄﾞﾀｲﾑ(1)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LEAD_TIME1).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSHI_LEAD_TIME1).Text) Then
'2016.05.18                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(ﾘｰﾄﾞﾀｲﾑ(1))"
                    Text1(ptxSHI_LEAD_TIME1).SetFocus
                    Exit Function
                End If
        
                Text1(ptxSHI_LEAD_TIME1).Text = Format(CLng(Text1(ptxSHI_LEAD_TIME1).Text), "#0")
            
            End If
        
        Case ptxSHI_LAST_ORDER_DT1     '前回注文日(1)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LAST_ORDER_DT1).Text) = "" Then
            Else
                If Not IsDate(Text1(ptxSHI_LAST_ORDER_DT1).Text) Then
'2016.05.18                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(前回注文日(1))"
                    Text1(ptxSHI_LAST_ORDER_DT1).SetFocus
                    Exit Function
                End If
            
                Text1(ptxSHI_LAST_ORDER_DT1).Text = Format(Text1(ptxSHI_LAST_ORDER_DT1).Text, "YYYY/MM/DD")
            End If
        
        Case ptxSHI_LAST_ORDER_QTY1    '前回注文数(1)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LAST_ORDER_QTY1).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSHI_LAST_ORDER_QTY1).Text) Then
'2016.05.18                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。('前回注文数(1))"
                    Text1(ptxSHI_LAST_ORDER_QTY1).SetFocus
                    Exit Function
                End If
        
                Text1(ptxSHI_LAST_ORDER_QTY1).Text = Format(CLng(Text1(ptxSHI_LAST_ORDER_QTY1).Text), "#0")
        
            End If
        
        Case ptxSHI_CODE2          '仕入先(2)
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            Text1(Mode).Text = StrConv(Text1(Mode).Text, vbUpperCase) '2016.01.26
            If Trim(Text1(ptxSHI_CODE2).Text) = "" Then
                For i = ptxSHI_CODE2 To ptxSHI_LAST_ORDER_QTY2
                    Text1(i).Text = ""
                Next i
                Combo1(pcmbSHIIRE2).ListIndex = -1
            Else
                For i = 0 To Combo1(pcmbSHIIRE2).ListCount - 1
                    If Trim(Text1(ptxSHI_CODE2).Text) = Trim(Right(Combo1(pcmbSHIIRE2).List(i), 5)) Then
                        Combo1(pcmbSHIIRE2).ListIndex = i
                        Exit For
                    End If
                Next i
                
                If i > (Combo1(pcmbSHIIRE2).ListCount - 1) Then
'2016.05.18                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(仕入先(2))"
                    Text1(ptxSHI_CODE2).SetFocus
                    Exit Function
                End If
                        
            
            
            
            End If
        
        Case ptxSHI_TANKA2         '仕入単価(2)
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_TANKA2).Text) = "" Then
            
                Text1(ptxSHI_ARARI2).Text = ""
                Text1(ptxSHI_ARARI_RITU2).Text = ""
            
            Else
                
                If Not IsNumeric(Text1(ptxSHI_TANKA2).Text) Then
'2016.05.18                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。仕入単価(2)"
                    Text1(ptxSHI_TANKA2).SetFocus
                    Exit Function
                End If
            
                Text1(ptxSHI_TANKA2).Text = Format(CDbl(Text1(ptxSHI_TANKA2).Text), "#0.00")
            
                If Not IsNumeric(Text1(ptxG_ST_URITAN).Text) Or _
                    Not IsNumeric(Text1(ptxG_ST_URITAN).Text) Then
                                        
                    Text1(ptxSHI_ARARI2).Text = ""
                    Text1(ptxSHI_ARARI_RITU2).Text = ""
                Else
                    Text1(ptxSHI_ARARI2).Text = Format(CDbl(Text1(ptxG_ST_URITAN).Text) - CDbl(Text1(ptxSHI_TANKA2).Text), "#0.00")
                    If CDbl(Text1(ptxG_ST_URITAN).Text) = 0 Then
                        Text1(ptxSHI_ARARI_RITU2).Text = "0.00"
                    Else
                        Text1(ptxSHI_ARARI_RITU2).Text = Format(CDbl(Text1(ptxSHI_ARARI2).Text) / CDbl(Text1(ptxG_ST_URITAN).Text) * 100, "#0.00")
                    End If
                End If
            
            End If
        
        
        
        
        
                
        Case ptxSHI_TANKA_DT2     '仕入単価設定日(2)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_TANKA_DT2).Text) = "" Then
                If Trim(Text1(ptxSHI_TANKA2).Text) <> "" Then
                    Text1(ptxSHI_TANKA_DT2).Text = Format(Now, "YYYY/MM/DD")
                End If
            Else
                If Not IsDate(Text1(ptxSHI_TANKA_DT2).Text) Then
'                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。仕入単価設定日(2)"
                    Text1(ptxSHI_TANKA_DT2).SetFocus
                    Exit Function
                End If
            
                Text1(ptxSHI_TANKA_DT2).Text = Format(Text1(ptxSHI_TANKA_DT2).Text, "YYYY/MM/DD")
            End If
                
        
        Case ptxSHI_LOT2           'ﾛｯﾄ数(2)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LOT2).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSHI_LOT2).Text) Then
'                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。ﾛｯﾄ数(2)"
                    Text1(ptxSHI_LOT2).SetFocus
                    Exit Function
                End If
        
                Text1(ptxSHI_LOT2).Text = Format(CLng(Text1(ptxSHI_LOT2).Text), "#0")
        
            End If
        
        Case ptxSHI_LEAD_TIME2     'ﾘｰﾄﾞﾀｲﾑ(2)
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            
            If Trim(Text1(ptxSHI_LEAD_TIME2).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSHI_LEAD_TIME2).Text) Then
'                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(ﾘｰﾄﾞﾀｲﾑ(2))"
                    Text1(ptxSHI_LEAD_TIME2).SetFocus
                    Exit Function
                End If
        
                Text1(ptxSHI_LEAD_TIME2).Text = Format(CLng(Text1(ptxSHI_LEAD_TIME2).Text), "#0")
            
            End If
        
        Case ptxSHI_LAST_ORDER_DT2     '前回注文日(2)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LAST_ORDER_DT2).Text) = "" Then
            Else
                If Not IsDate(Text1(ptxSHI_LAST_ORDER_DT2).Text) Then
'                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(前回注文日(2))"
                    Text1(ptxSHI_LAST_ORDER_DT2).SetFocus
                    Exit Function
                End If
            
                Text1(ptxSHI_LAST_ORDER_DT2).Text = Format(Text1(ptxSHI_LAST_ORDER_DT2).Text, "YYYY/MM/DD")
            End If
        
        Case ptxSHI_LAST_ORDER_QTY2    '前回注文数(2)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LAST_ORDER_QTY2).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSHI_LAST_ORDER_QTY2).Text) Then
'                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(前回注文数(2))"
                    Text1(ptxSHI_LAST_ORDER_QTY2).SetFocus
                    Exit Function
                End If
        
                Text1(ptxSHI_LAST_ORDER_QTY2).Text = Format(CLng(Text1(ptxSHI_LAST_ORDER_QTY2).Text), "#0")
        
            End If
        Case ptxSHI_CODE3          '仕入先(3)
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            Text1(Mode).Text = StrConv(Text1(Mode).Text, vbUpperCase) '2016.01.26
            If Trim(Text1(ptxSHI_CODE3).Text) = "" Then
                For i = ptxSHI_CODE3 To ptxSHI_LAST_ORDER_QTY3
                    Text1(i).Text = ""
                Next i
                Combo1(pcmbSHIIRE3).ListIndex = -1
            Else
                For i = 0 To Combo1(pcmbSHIIRE3).ListCount - 1
                    If Trim(Text1(ptxSHI_CODE3).Text) = Trim(Right(Combo1(pcmbSHIIRE3).List(i), 5)) Then
                        Combo1(pcmbSHIIRE3).ListIndex = i
                        Exit For
                    End If
                Next i
                
                If i > (Combo1(pcmbSHIIRE3).ListCount - 1) Then
'                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(仕入先(3))"
                    Text1(ptxSHI_CODE3).SetFocus
                    Exit Function
                End If
                        
            
            
            
            End If
        
        Case ptxSHI_TANKA3         '仕入単価(3)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
        
            If Trim(Text1(ptxSHI_TANKA3).Text) = "" Then
            
                Text1(ptxSHI_ARARI3).Text = ""
                Text1(ptxSHI_ARARI_RITU3).Text = ""
            
            Else
                
                If Not IsNumeric(Text1(ptxSHI_TANKA3).Text) Then
'                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(仕入単価(3))"
                    Text1(ptxSHI_TANKA3).SetFocus
                    Exit Function
                End If
            
                Text1(ptxSHI_TANKA3).Text = Format(CDbl(Text1(ptxSHI_TANKA3).Text), "#0.00")
            
                If Not IsNumeric(Text1(ptxG_ST_URITAN).Text) Or _
                    Not IsNumeric(Text1(ptxG_ST_URITAN).Text) Then
                                        
                    Text1(ptxSHI_ARARI3).Text = ""
                    Text1(ptxSHI_ARARI_RITU3).Text = ""
                Else
                    Text1(ptxSHI_ARARI3).Text = Format(CDbl(Text1(ptxG_ST_URITAN).Text) - CDbl(Text1(ptxSHI_TANKA3).Text), "#0.00")
                                    
                    If CDbl(Text1(ptxG_ST_URITAN).Text) = 0 Then
                        Text1(ptxSHI_ARARI_RITU3).Text = "0.00"
                    Else
                        Text1(ptxSHI_ARARI_RITU3).Text = Format(CDbl(Text1(ptxSHI_ARARI3).Text) / CDbl(Text1(ptxG_ST_URITAN).Text) * 100, "#0.00")
                    End If
                End If
            
            End If
        
                
        Case ptxSHI_TANKA_DT3     '仕入単価設定日(3)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_TANKA_DT3).Text) = "" Then
                If Trim(Text1(ptxSHI_TANKA3).Text) <> "" Then
                    Text1(ptxSHI_TANKA_DT3).Text = Format(Now, "YYYY/MM/DD")
                End If
            Else
                If Not IsDate(Text1(ptxSHI_TANKA_DT3).Text) Then
'                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(仕入単価設定日(3))"
                    Text1(ptxSHI_TANKA_DT3).SetFocus
                    Exit Function
                End If
            
                Text1(ptxSHI_TANKA_DT3).Text = Format(Text1(ptxSHI_TANKA_DT3).Text, "YYYY/MM/DD")
            End If
                
        
        Case ptxSHI_LOT3           'ﾛｯﾄ数(3)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LOT3).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSHI_LOT3).Text) Then
'                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(ﾛｯﾄ数(3))"
                    Text1(ptxSHI_LOT3).SetFocus
                    Exit Function
                End If
        
                Text1(ptxSHI_LOT3).Text = Format(CLng(Text1(ptxSHI_LOT3).Text), "#0")
        
            End If
        
        Case ptxSHI_LEAD_TIME3     'ﾘｰﾄﾞﾀｲﾑ(3)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LEAD_TIME3).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSHI_LEAD_TIME3).Text) Then
'                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(ﾘｰﾄﾞﾀｲﾑ(3))"
                    Text1(ptxSHI_LEAD_TIME3).SetFocus
                    Exit Function
                End If
        
                Text1(ptxSHI_LEAD_TIME3).Text = Format(CLng(Text1(ptxSHI_LEAD_TIME3).Text), "#0")
            
            End If
        
        Case ptxSHI_TANKA_DT3     '前回注文日(3)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LAST_ORDER_DT3).Text) = "" Then
            Else
                If Not IsDate(Text1(ptxSHI_LAST_ORDER_DT3).Text) Then
'                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(前回注文日(3))"
                    Text1(ptxSHI_LAST_ORDER_DT3).SetFocus
                    Exit Function
                End If
            
                Text1(ptxSHI_LAST_ORDER_DT3).Text = Format(Text1(ptxSHI_LAST_ORDER_DT2).Text, "YYYY/MM/DD")
            End If
        
        Case ptxSHI_LAST_ORDER_QTY3    '前回注文数(3)
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxSHI_LAST_ORDER_QTY3).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxSHI_LAST_ORDER_QTY3).Text) Then
'                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(前回注文数(3))"
                    Text1(ptxSHI_LAST_ORDER_QTY3).SetFocus
                    Exit Function
                End If
        
                Text1(ptxSHI_LAST_ORDER_QTY3).Text = Format(CLng(Text1(ptxSHI_LAST_ORDER_QTY3).Text), "#0")
        
            End If
        
        
        Case ptxLAST_SYU_DT     '最終出庫日
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            If Trim(Text1(ptxLAST_SYU_DT).Text) = "" Then
            Else
                If Not IsDate(Text1(ptxLAST_SYU_DT).Text) Then
'                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(最終出庫日)"
                    Text1(ptxLAST_SYU_DT).SetFocus
                    Exit Function
                End If
            
                Text1(ptxLAST_SYU_DT).Text = Format(Text1(ptxLAST_SYU_DT).Text, "YYYY/MM/DD")
            End If
        
        Case ptxG_LAST_SYUKA_QTY    '最終出庫数
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
            If Trim(Text1(ptxG_LAST_SYUKA_QTY).Text) = "" Then
            Else
                If Not IsNumeric(Text1(ptxG_LAST_SYUKA_QTY).Text) Then
'                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(最終出庫数)"
                    Text1(ptxG_LAST_SYUKA_QTY).SetFocus
                    Exit Function
                End If
            
                Text1(ptxG_LAST_SYUKA_QTY).Text = Format(CLng(Text1(ptxG_LAST_SYUKA_QTY).Text), "#0")
            End If
        
        
        
        
        
        
        Case ptxLAST_CODE          '最新仕入先      2007.05.28
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            Text1(Mode).Text = StrConv(Text1(Mode).Text, vbUpperCase) '2016.01.26
            If Trim(Text1(ptxLAST_CODE).Text) = "" Then
                Combo1(pcmbLAST_CODE).ListIndex = -1
            Else
                For i = 0 To Combo1(pcmbLAST_CODE).ListCount - 1
                    If Trim(Text1(ptxLAST_CODE).Text) = Trim(Right(Combo1(pcmbLAST_CODE).List(i), 5)) Then
                        Combo1(pcmbLAST_CODE).ListIndex = i
                        Exit For
                    End If
                Next i
                
                If i > (Combo1(pcmbLAST_CODE).ListCount - 1) Then
'                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(最新仕入先)"
                    Text1(ptxLAST_CODE).SetFocus
                    Exit Function
                End If
                        
            
            
            
            End If
        
        Case ptxLAST_TANKA          '最新仕入単価    2007.05.28
        
            If Last_JGYOBU <> SHIZAI Then       '2016.05.18
                Error_Check_Proc = False        '2016.05.18
                Exit Function                   '2016.05.18
            End If                              '2016.05.18
        
            If Trim(Text1(ptxLAST_TANKA).Text) = "" Then
            
            
            Else
                
                If Not IsNumeric(Text1(ptxLAST_TANKA).Text) Then
'                    MsgBox "入力した項目はエラーです。"
                    MsgBox "入力した項目はエラーです。(最新仕入単価)"
                    Text1(ptxLAST_TANKA).SetFocus
                    Exit Function
                End If
            
                Text1(ptxLAST_TANKA).Text = Format(CDbl(Text1(ptxLAST_TANKA).Text), "#0.00")
            
            
            End If
        
        
        
        
        
        Case ptxKUTI_SU             '口数    2010.01.18
        
        
            If Trim(Text1(ptxKUTI_SU).Text) = "" Then
            
            
            Else
                
                If Not IsNumeric(Text1(ptxKUTI_SU).Text) Then
                    MsgBox "入力した項目はエラーです。(口数)"
                    Text1(ptxKUTI_SU).SetFocus
                    Exit Function
                End If
            
                Text1(ptxKUTI_SU).Text = Format(CCur(Text1(ptxKUTI_SU).Text), "#0.0")
            
            
            End If
        
        
        
        Case ptxKONPOU_F             '梱包区分    2010.01.18
        
        
            If Trim(Text1(ptxKONPOU_F).Text) = "" Or Trim(Text1(ptxKONPOU_F).Text) = "0" Or Trim(Text1(ptxKONPOU_F).Text) = "1" Then
            
            
            Else
                
                MsgBox "入力した項目はエラーです。(梱包区分)"
                Text1(ptxKONPOU_F).SetFocus
                Exit Function
            
            
            
            End If
        
        
        
    End Select
        
    Error_Check_Proc = False
    Exit Function


Error_Proc:
    
    If Err.Number = 6 Then
        MsgBox "桁数オーバーです。入力内容を確認してください。"
    
    
    
    Else
        MsgBox "入力異常です。入力内容を確認してください。"
    End If
    Text1(Mode).SetFocus

End Function
Private Function Item_Disp_Proc(CODE As String) As Integer
'----------------------------------------------------------------------------
'                   画面表示
'----------------------------------------------------------------------------
Dim sts As Integer
Dim i   As Integer

    Item_Disp_Proc = True
    
    '品目ﾏｽﾀ読み込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, Left(CODE, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(CODE, 2, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Right(CODE, 20))
    
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            'ﾚｺｰﾄﾞ内容の表示
                                            '国内外
            For i = 0 To Combo1(pcmbNAIGAI).ListCount - 1
                If Right(Combo1(pcmbNAIGAI).List(i), 1) = StrConv(ITEMREC.JGYOBU, vbUnicode) Then
                    Combo1(pcmbNAIGAI).ListIndex = i
                    Exit For
                End If
            Next i
                                            '品目ｺｰﾄﾞ
            Text1(ptxHIN_GAI).Text = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                            '品名
            Text1(ptxHIN_NAME).Text = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                                            '仕入区分
            Text1(ptxG_SHIIRE_KBN).Text = StrConv(ITEMREC.G_SHIIRE_KBN, vbUnicode)
                                            '仕入区分検索
            Combo1(pcmbG_SHIIRE).ListIndex = -1
            For i = 0 To Combo1(pcmbG_SHIIRE).ListCount - 1
                If Left(Right(Combo1(pcmbG_SHIIRE).List(i), 3), 2) = Text1(ptxG_SHIIRE_KBN).Text Then
                    Combo1(pcmbG_SHIIRE).ListIndex = i
                    Exit For
                End If
            Next i
                                            '販売区分
            Text1(ptxG_HANBAI_KBN).Text = StrConv(ITEMREC.G_HANBAI_KBN, vbUnicode)
                                            '販売区分検索
            Combo1(pcmbG_HANBAI).ListIndex = -1
            For i = 0 To Combo1(pcmbG_HANBAI).ListCount - 1
                If Left(Right(Combo1(pcmbG_HANBAI).List(i), 3), 2) = Text1(ptxG_HANBAI_KBN).Text Then
                    Combo1(pcmbG_HANBAI).ListIndex = i
                    Exit For
                End If
            Next i
                                            '収支単位
            Text1(ptxG_SYUSHI).Text = StrConv(ITEMREC.G_SYUSHI, vbUnicode)
                                            '収支単位検索
            Combo1(pcmbG_SYUSHI).ListIndex = -1
            For i = 0 To Combo1(pcmbG_SYUSHI).ListCount - 1
                If Right(Combo1(pcmbG_SYUSHI).List(i), 3) = Text1(ptxG_SYUSHI).Text Then
                    Combo1(pcmbG_SYUSHI).ListIndex = i
                    Exit For
                End If
            Next i
                                            '組立製品
            If StrConv(ITEMREC.G_KUMITATE, vbUnicode) = P_ASSEMBLY_ON Then
                Check1(pchkG_KUMITATE).Value = vbChecked
            Else
                Check1(pchkG_KUMITATE).Value = vbUnchecked
            End If
                                            '前月在庫金額
            If Trim(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode)) = "" Then
                Text1(ptxG_ZEN_ZAIKO_KIN).Text = ""
            Else
                Text1(ptxG_ZEN_ZAIKO_KIN).Text = Format(CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode)), "#0")
            End If
                                            '前月在庫数量
            If Not IsNumeric(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode)) Then
                Text1(ptxG_ZEN_ZAIKO_QTY).Text = ""
            Else
                Text1(ptxG_ZEN_ZAIKO_QTY).Text = Format(CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode)), "#0")
            End If
                                            '標準売価
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_ST_URITAN, vbUnicode))) Then
                Text1(ptxG_ST_URITAN).Text = ""
            Else
                Text1(ptxG_ST_URITAN).Text = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
            End If
                                            '標準売価設定日
            If Trim(StrConv(ITEMREC.G_ST_URITAN_DT, vbUnicode)) = "" Then
                Text1(ptxG_ST_URITAN_DT).Text = ""
            Else
                Text1(ptxG_ST_URITAN_DT).Text = Left(StrConv(ITEMREC.G_ST_URITAN_DT, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(ITEMREC.G_ST_URITAN_DT, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(ITEMREC.G_ST_URITAN_DT, vbUnicode), 2)
            End If
                                            '標準原価
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode))) Then
                Text1(ptxG_ST_SHITAN).Text = ""
            Else
                Text1(ptxG_ST_SHITAN).Text = Format(CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)), "#0.00")
            End If
                                            '標準原価設定日
            If Trim(StrConv(ITEMREC.G_ST_SHITAN_DT, vbUnicode)) = "" Then
                Text1(ptxG_ST_SHITAN_DT).Text = ""
            Else
                Text1(ptxG_ST_SHITAN_DT).Text = Left(StrConv(ITEMREC.G_ST_SHITAN_DT, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(ITEMREC.G_ST_SHITAN_DT, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(ITEMREC.G_ST_SHITAN_DT, vbUnicode), 2)
            End If
                                            
                                            '危険在庫
            If Not IsNumeric(Trim(StrConv(ITEMREC.HOJYU_P, vbUnicode))) Then
                Text1(ptxHOJYU_P).Text = ""
            Else
                Text1(ptxHOJYU_P).Text = Format(CDbl(StrConv(ITEMREC.HOJYU_P, vbUnicode)), "#0")
            End If
                                            '資材区分
            Combo1(pcmbG_SHIZAI_KBN).ListIndex = -1
            For i = 0 To Combo1(pcmbG_SHIZAI_KBN).ListCount - 1
                If Right(Combo1(pcmbG_SHIZAI_KBN).List(i), 1) = StrConv(ITEMREC.G_SHIZAI_KBN, vbUnicode) Then
                    Combo1(pcmbG_SHIZAI_KBN).ListIndex = i
                    Exit For
                End If
            Next i
                                            'ﾗﾍﾞﾙ貼り計上なし
            If StrConv(ITEMREC.G_LABEL_NON, vbUnicode) = P_G_LABEL_ON Then
                Check1(pchkG_LABEL_NON).Value = vbUnchecked
            Else
                Check1(pchkG_LABEL_NON).Value = vbChecked
            End If
                                            
                                            '在庫管理対象外
            If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_ON Then
                Check1(pchkZAIKO_F).Value = vbUnchecked
            Else
                Check1(pchkZAIKO_F).Value = vbChecked
            End If
                                            '棚卸表　在庫数非表示   2012.12.13
            If StrConv(ITEMREC.ZAIKO_CLR_F, vbUnicode) = "1" Then
                Check1(pchkZAIKO_CLR_F).Value = vbChecked
            Else
                Check1(pchkZAIKO_CLR_F).Value = vbUnchecked
            End If
                                            '才数   2008.02.14
            If IsNumeric(StrConv(ITEMREC.SAI_SU, vbUnicode)) Then
                Text1(ptxSAI_SU).Text = Format(CDbl(StrConv(ITEMREC.SAI_SU, vbUnicode)), "#0.00")
            Else
                Text1(ptxSAI_SU).Text = ""
            End If
                                            
                                            '形式 2008.02.14
            Text1(ptxD_KEISHIKI).Text = Trim(StrConv(ITEMREC.D_KEISHIKI, vbUnicode))
                                            'ﾀﾞﾝﾎﾞｰﾙ厚さ 2008.02.14
            Text1(ptxD_THICKNESS).Text = Trim(StrConv(ITEMREC.D_THICKNESS, vbUnicode))
                                            'ﾀﾞﾝﾎﾞｰﾙ材質 2008.02.14
            Text1(ptxD_MATERIAL).Text = Trim(StrConv(ITEMREC.D_MATERIAL, vbUnicode))
                                            
                                            
                                            'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ(W) 2008.02.14
            If IsNumeric(StrConv(ITEMREC.D_SIZE_W, vbUnicode)) Then
                Text1(ptxD_SIZE_W).Text = Format(Val(Trim(StrConv(ITEMREC.D_SIZE_W, vbUnicode))), "#0")
            Else
                Text1(ptxD_SIZE_W).Text = Trim(StrConv(ITEMREC.D_SIZE_W, vbUnicode))
            End If
                                            
                                            
                                            'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ(D) 2008.02.14
                                            
            If IsNumeric(StrConv(ITEMREC.D_SIZE_D, vbUnicode)) Then
                Text1(ptxD_SIZE_D).Text = Format(Val(Trim(StrConv(ITEMREC.D_SIZE_D, vbUnicode))), "#0")
            Else
                Text1(ptxD_SIZE_D).Text = Trim(StrConv(ITEMREC.D_SIZE_D, vbUnicode))
            End If
                                            
                                            
                                            'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ(H) 2008.02.14
            If IsNumeric(StrConv(ITEMREC.D_SIZE_H, vbUnicode)) Then
                Text1(ptxD_SIZE_H).Text = Format(Val(Trim(StrConv(ITEMREC.D_SIZE_H, vbUnicode))), "#0")
            Else
                Text1(ptxD_SIZE_H).Text = Trim(StrConv(ITEMREC.D_SIZE_H, vbUnicode))
            End If
                                            
                                            
            If IsNumeric(Text1(ptxD_SIZE_W).Text) And IsNumeric(Text1(ptxD_SIZE_D).Text) And IsNumeric(Text1(ptxD_SIZE_H).Text) Then
            
                
            
                lblSIZE.Caption = Format(ToHalfAdjust(CCur(Val(Text1(ptxD_SIZE_W).Text) / 1000 * Val(Text1(ptxD_SIZE_D).Text) / 1000 * Val(Text1(ptxD_SIZE_H).Text) / 1000), 4), "#0.000")
                                                
            End If
'            lblSIZE.Caption = (Val(Text1(ptxD_SIZE_W).Text) * Val(Text1(ptxD_SIZE_D).Text) * Val(Text1(ptxD_SIZE_H).Text)) / 1000000000
                                            
                                            
                                            
                                            '印刷する／しない  2008.02.14
            Text1(ptxD_PRINT).Text = Trim(StrConv(ITEMREC.D_PRINT, vbUnicode))
                                            
                                            '商品化　工数   2008.02.14
            If IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
                Text1(ptxS_KOUSU).Text = Format(CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode)), "#")
            Else
                Text1(ptxS_KOUSU).Text = ""
            End If
                                            '集合梱包       2008.07.16
            If IsNumeric(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)) Then
                Text1(ptxSEI_SYU_KON).Text = Format(CDbl(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)), "#")
            Else
                Text1(ptxSEI_SYU_KON).Text = ""
            End If
                                            
                                            
                                            '輸送箱出力ﾌﾗｸﾞ 2008.02.14
            Text1(ptxSE_USOU_F).Text = Trim(StrConv(ITEMREC.SE_USOU_F, vbUnicode))
                                            '使用テープ種類 2008.02.14
            Text1(ptxUSE_TAPE_KIND).Text = Trim(StrConv(ITEMREC.USE_TAPE_KIND, vbUnicode))
                                            '使用テープ長さ 2008.02.14
            Text1(ptxUSE_TAPE_LNG).Text = Trim(StrConv(ITEMREC.USE_TAPE_LNG, vbUnicode))
                                            '請求区分       2008.07.16
            Text1(ptxSEI_KBN).Text = Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode))
                                            
                                            
                                            '標準棚番       2009.09.01
            Text1(ptxST_SOKO).Text = Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                            '標準棚番       2009.09.01
            Text1(ptxST_RETU).Text = Trim(StrConv(ITEMREC.ST_RETU, vbUnicode))
                                            '標準棚番       2009.09.01
            Text1(ptxST_REN).Text = Trim(StrConv(ITEMREC.ST_REN, vbUnicode))
                                            '標準棚番       2009.09.01
            Text1(ptxST_DAN).Text = Trim(StrConv(ITEMREC.ST_DAN, vbUnicode))
                                            
                                            
                                            
                                            '口数           2010.01.18
            Text1(ptxKUTI_SU).Text = Trim(StrConv(ITEMREC.KUTI_SU, vbUnicode))
                                            '梱包区分           2010.01.18
            Text1(ptxKONPOU_F).Text = Trim(StrConv(ITEMREC.KONPOU_F, vbUnicode))
                                            
                                            
                                            
                                            
                                            '仕入先ｺ-ﾄﾞ(1)
            Text1(ptxSHI_CODE1).Text = Trim(StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
                                            '仕入先名称(1)
            Combo1(pcmbSHIIRE1).ListIndex = -1
            For i = 0 To Combo1(pcmbSHIIRE1).ListCount - 1
                If Trim(Right(Combo1(pcmbSHIIRE1).List(i), 5)) = Text1(ptxSHI_CODE1).Text Then
                    Combo1(pcmbSHIIRE1).ListIndex = i
                    Exit For
                End If
            Next i
                                            '仕入単価(1)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))) Then
                Text1(ptxSHI_TANKA1).Text = ""
            Else
                Text1(ptxSHI_TANKA1).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)), "#0.00")
            End If
                                            '仕入単価設定日(1)
            If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA_DT, vbUnicode)) = "" Then
                Text1(ptxSHI_TANKA_DT1).Text = ""
            Else
                Text1(ptxSHI_TANKA_DT1).Text = Left(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA_DT, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA_DT, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA_DT, vbUnicode), 2)
            End If
                                            '粗利／粗利率(1)
            If Not IsNumeric(Text1(ptxG_ST_URITAN).Text) Or _
                Not IsNumeric(Text1(ptxSHI_TANKA1).Text) Then
                                    
                Text1(ptxSHI_ARARI1).Text = ""
                Text1(ptxSHI_ARARI_RITU1).Text = ""
            Else
                Text1(ptxSHI_ARARI1).Text = Format(CDbl(Text1(ptxG_ST_URITAN).Text) - CDbl(Text1(ptxSHI_TANKA1).Text), "#0.00")
                If CDbl(Text1(ptxG_ST_URITAN).Text) = 0 Then
                    Text1(ptxSHI_ARARI_RITU1).Text = "0.00"
                Else
                    Text1(ptxSHI_ARARI_RITU1).Text = Format(CDbl(Text1(ptxSHI_ARARI1).Text) / CDbl(Text1(ptxG_ST_URITAN).Text) * 100, "#0.00")
                End If
            End If
                                            'ﾛｯﾄ数(1)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode))) Then
                Text1(ptxSHI_LOT1).Text = ""
            Else
                Text1(ptxSHI_LOT1).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode)), "#0")
            End If
                                            'ﾘｰﾄﾞﾀｲﾑ(1)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(0).LEAD_TIME, vbUnicode))) Then
                Text1(ptxSHI_LEAD_TIME1).Text = ""
            Else
                Text1(ptxSHI_LEAD_TIME1).Text = Format(CLng(StrConv(ITEMREC.G_SHIIRE_TBL(0).LEAD_TIME, vbUnicode)), "#0")
            End If
                                            '前回注文日(1)
            If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_DT, vbUnicode)) = "" Then
                Text1(ptxSHI_LAST_ORDER_DT1).Text = ""
            Else
                Text1(ptxSHI_LAST_ORDER_DT1).Text = Left(StrConv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_DT, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_DT, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_DT, vbUnicode), 2)
            End If
                                            '前回注文数(1)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_QTY, vbUnicode))) Then
                Text1(ptxSHI_LAST_ORDER_QTY1).Text = ""
            Else
                Text1(ptxSHI_LAST_ORDER_QTY1).Text = Format(CLng(StrConv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_QTY, vbUnicode)), "#0")
            End If
            
            
            
            
            
            
                                            '仕入先ｺ-ﾄﾞ(2)
            Text1(ptxSHI_CODE2).Text = Trim(StrConv(ITEMREC.G_SHIIRE_TBL(1).CODE, vbUnicode))
                                            '仕入先名称(2)
            Combo1(pcmbSHIIRE2).ListIndex = -1
            For i = 0 To Combo1(pcmbSHIIRE2).ListCount - 1
                If Trim(Right(Combo1(pcmbSHIIRE2).List(i), 5)) = Text1(ptxSHI_CODE2).Text Then
                    Combo1(pcmbSHIIRE2).ListIndex = i
                    Exit For
                End If
            Next i
                                            '仕入単価(2)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(1).TANKA, vbUnicode))) Then
                Text1(ptxSHI_TANKA2).Text = ""
            Else
                Text1(ptxSHI_TANKA2).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(1).TANKA, vbUnicode)), "#0.00")
            End If
                                            '仕入単価設定日(2)
            If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(1).TANKA_DT, vbUnicode)) = "" Then
                Text1(ptxSHI_TANKA_DT2).Text = ""
            Else
                Text1(ptxSHI_TANKA_DT2).Text = Left(StrConv(ITEMREC.G_SHIIRE_TBL(1).TANKA_DT, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(ITEMREC.G_SHIIRE_TBL(1).TANKA_DT, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(ITEMREC.G_SHIIRE_TBL(1).TANKA_DT, vbUnicode), 2)
            End If
                                            '粗利／粗利率(2)
            If Not IsNumeric(Text1(ptxG_ST_URITAN).Text) Or _
                Not IsNumeric(Text1(ptxSHI_TANKA2).Text) Then
                Text1(ptxSHI_ARARI2).Text = ""
                Text1(ptxSHI_ARARI_RITU2).Text = ""
            Else
                Text1(ptxSHI_ARARI2).Text = Format(CDbl(Text1(ptxG_ST_URITAN).Text) - CDbl(Text1(ptxSHI_TANKA2).Text), "#0.00")
                If CDbl(Text1(ptxG_ST_URITAN).Text) = 0 Then
                    Text1(ptxSHI_ARARI_RITU2).Text = "0.00"
                Else
                    Text1(ptxSHI_ARARI_RITU2).Text = Format(CDbl(Text1(ptxSHI_ARARI2).Text) / CDbl(Text1(ptxG_ST_URITAN).Text) * 100, "#0.00")
                End If
            End If
                                            'ﾛｯﾄ数(2)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(1).LOT, vbUnicode))) Then
                Text1(ptxSHI_LOT2).Text = ""
            Else
                Text1(ptxSHI_LOT2).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(1).LOT, vbUnicode)), "#0")
            End If
                                            'ﾘｰﾄﾞﾀｲﾑ(2)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(1).LEAD_TIME, vbUnicode))) Then
                Text1(ptxSHI_LEAD_TIME2).Text = ""
            Else
                Text1(ptxSHI_LEAD_TIME2).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(1).LEAD_TIME, vbUnicode)), "#0")
            End If
                                            '前回注文日(2)
            If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_DT, vbUnicode)) = "" Then
                Text1(ptxSHI_LAST_ORDER_DT2).Text = ""
            Else
                Text1(ptxSHI_LAST_ORDER_DT2).Text = Left(StrConv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_DT, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_DT, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_DT, vbUnicode), 2)
            End If
                                            '前回注文数(2)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_QTY, vbUnicode))) Then
                Text1(ptxSHI_LAST_ORDER_QTY2).Text = ""
            Else
                Text1(ptxSHI_LAST_ORDER_QTY2).Text = Format(CLng(StrConv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_QTY, vbUnicode)), "#0")
            End If
            
            
            
            
                                            '仕入先ｺ-ﾄﾞ(3)
            Text1(ptxSHI_CODE3).Text = Trim(StrConv(ITEMREC.G_SHIIRE_TBL(2).CODE, vbUnicode))
                                            '仕入先名称(3)
            Combo1(pcmbSHIIRE3).ListIndex = -1
            For i = 0 To Combo1(pcmbSHIIRE3).ListCount - 1
                If Trim(Right(Combo1(pcmbSHIIRE3).List(i), 5)) = Text1(ptxSHI_CODE3).Text Then
                    Combo1(pcmbSHIIRE3).ListIndex = i
                    Exit For
                End If
            Next i
                                            '仕入単価(3)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(2).TANKA, vbUnicode))) Then
                Text1(ptxSHI_TANKA3).Text = ""
            Else
                Text1(ptxSHI_TANKA3).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(2).TANKA, vbUnicode)), "#0.00")
            End If
                                            '仕入単価設定日(3)
            If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(2).TANKA_DT, vbUnicode)) = "" Then
                Text1(ptxSHI_TANKA_DT3).Text = ""
            Else
                Text1(ptxSHI_TANKA_DT3).Text = Left(StrConv(ITEMREC.G_SHIIRE_TBL(2).TANKA_DT, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(ITEMREC.G_SHIIRE_TBL(2).TANKA_DT, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(ITEMREC.G_SHIIRE_TBL(2).TANKA_DT, vbUnicode), 2)
            End If
                                            '粗利／粗利率(3)
            If Not IsNumeric(Text1(ptxG_ST_URITAN).Text) Or _
                Not IsNumeric(Text1(ptxSHI_TANKA3).Text) Then
                Text1(ptxSHI_ARARI3).Text = ""
                Text1(ptxSHI_ARARI_RITU3).Text = ""
            Else
                Text1(ptxSHI_ARARI3).Text = Format(CDbl(Text1(ptxG_ST_URITAN).Text) - CDbl(Text1(ptxSHI_TANKA3).Text), "#0.00")
                If CDbl(Text1(ptxG_ST_URITAN).Text) = 0 Then
                    Text1(ptxSHI_ARARI_RITU3).Text = "0.00"
                Else
                    Text1(ptxSHI_ARARI_RITU3).Text = Format(CDbl(Text1(ptxSHI_ARARI3).Text) / CDbl(Text1(ptxG_ST_URITAN).Text) * 100, "#0.00")
                End If
            End If
                                            'ﾛｯﾄ数(3)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(2).LOT, vbUnicode))) Then
                Text1(ptxSHI_LOT3).Text = ""
            Else
                Text1(ptxSHI_LOT3).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(2).LOT, vbUnicode)), "#0")
            End If
                                            'ﾘｰﾄﾞﾀｲﾑ(3)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(2).LEAD_TIME, vbUnicode))) Then
                Text1(ptxSHI_LEAD_TIME3).Text = ""
            Else
                Text1(ptxSHI_LEAD_TIME3).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(2).LEAD_TIME, vbUnicode)), "#0")
            End If
                                            '前回注文日(3)
            If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_DT, vbUnicode)) = "" Then
                Text1(ptxSHI_LAST_ORDER_DT3).Text = ""
            Else
                Text1(ptxSHI_LAST_ORDER_DT3).Text = Left(StrConv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_DT, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_DT, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_DT, vbUnicode), 2)
            End If
                                            '前回注文数(3)
            If Not IsNumeric(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_QTY, vbUnicode))) Then
                Text1(ptxSHI_LAST_ORDER_QTY3).Text = ""
            Else
                Text1(ptxSHI_LAST_ORDER_QTY3).Text = Format(CLng(StrConv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_QTY, vbUnicode)), "#0")
            End If
            



                                            '最終出庫日
            If Trim(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode)) = "" Then
                Text1(ptxLAST_SYU_DT).Text = ""
            Else
                Text1(ptxLAST_SYU_DT).Text = Left(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode), 2)
            End If
                                            
                                            '最終出庫数
            If Not IsNumeric(StrConv(ITEMREC.G_LAST_SYUKA_QTY, vbUnicode)) Then
                Text1(ptxG_LAST_SYUKA_QTY).Text = ""
            Else
                Text1(ptxG_LAST_SYUKA_QTY).Text = Format(CDbl(StrConv(ITEMREC.G_LAST_SYUKA_QTY, vbUnicode)), "#0")
            End If


                                            '最新仕入先 2007.05.28
            Text1(ptxLAST_CODE).Text = Trim(StrConv(ITEMREC.LAST_CODE, vbUnicode))
            Combo1(pcmbLAST_CODE).ListIndex = -1
            
            For i = 0 To Combo1(pcmbLAST_CODE).ListCount - 1
                If Trim(Right(Combo1(pcmbLAST_CODE).List(i), 5)) = Text1(ptxLAST_CODE).Text Then
                    Combo1(pcmbLAST_CODE).ListIndex = i
                    Exit For
                End If
            Next i
                                            '最新単価   2007.05.28
            If Not IsNumeric(Trim(StrConv(ITEMREC.LAST_TANKA, vbUnicode))) Then
                Text1(ptxLAST_TANKA).Text = ""
            Else
                Text1(ptxLAST_TANKA).Text = Format(CDbl(StrConv(ITEMREC.LAST_TANKA, vbUnicode)), "#0.00")
            End If

        
        
        
        
        
        
                                            '仕入備考   2018.04.19
            Text1(ptxSHIIRE_BIKOU).Text = StrConv(ITEMREC.SHIIRE_BIKOU, vbUnicode)
        
        
        
                                            '追加担当者/日時    2010.01.18
            lblIns_DateTime = StrConv(ITEMREC.INS_TANTO, vbUnicode) & " " & Mid(StrConv(ITEMREC.Ins_DateTime, vbUnicode), 1, 8) & "-" & Mid(StrConv(ITEMREC.Ins_DateTime, vbUnicode), 9, 4)
                                            
                                            
                                            
                                            '更新担当者/日時    2010.01.18
            lblUpd_DateTime = StrConv(ITEMREC.UPD_TANTO, vbUnicode) & " " & Mid(StrConv(ITEMREC.UPD_DATETIME, vbUnicode), 1, 8) & "-" & Mid(StrConv(ITEMREC.UPD_DATETIME, vbUnicode), 9, 4)
        
        
        
        
        
        
        
        
        
        
        Case BtErrKeyNotFound
        
            MsgBox "他端末で変更されています。前画面に戻ります。"
            PM000302.Visible = False
            INIT_FLG = False
            
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            PM000302.Visible = False
            INIT_FLG = False
    End Select

    Item_Disp_Proc = False

End Function

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   品目マスタ出力
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim com     As Integer
Dim i       As Integer

    Update_Proc = True
    
    '品目マスタ　読み込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
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
    '--------------------------------------------------------レコード内容編集
    
    If com = BtOpInsert Then
        
        
        Rclr_ITEMREC    '2018.04.19
        
        Call UniCode_Conv(ITEMREC.JGYOBU, Last_JGYOBU)              '事業部=資材
                                                                    '国内外
        Call UniCode_Conv(ITEMREC.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
        Call UniCode_Conv(ITEMREC.HIN_GAI, Text1(ptxHIN_GAI).Text)  '品目ｺｰﾄﾞ
        Call UniCode_Conv(ITEMREC.HIN_NAME, Text1(ptxHIN_NAME))     '品名
        Call UniCode_Conv(ITEMREC.ST_SET_DT, "")                    '標準棚番設定日付
        Call UniCode_Conv(ITEMREC.ST_SOKO, "")                      '標準入庫　倉庫
        Call UniCode_Conv(ITEMREC.ST_RETU, "")                      '標準入庫　列
        Call UniCode_Conv(ITEMREC.ST_REN, "")                       '標準入庫　連
        Call UniCode_Conv(ITEMREC.ST_DAN, "")                       '標準入庫　段
        Call UniCode_Conv(ITEMREC.BEF_SOKO, "")                     '前回入庫　倉庫
        Call UniCode_Conv(ITEMREC.BEF_RETU, "")                     '前回入庫　列
        Call UniCode_Conv(ITEMREC.BEF_REN, "")                      '前回入庫　連
        Call UniCode_Conv(ITEMREC.BEF_DAN, "")                      '前回入庫　段
        Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")                  '最終入庫日
        Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")                  '最終出庫日
        Call UniCode_Conv(ITEMREC.HIN_NAI, "")                      '品番（内）
        Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")                   'ﾎｽﾄ倉庫
        Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")                   'ﾎｽﾄ棚番
        Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")              '補充点
        Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")            '月平均出荷数
        Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                  'ｻﾝﾌﾟﾙ数
        Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                  'ｻﾝﾌﾟﾙ数
        Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")                  '最終入荷日付
        Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")                  '最終照合日付
        Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "00000000")         '照合時在庫数
        Call UniCode_Conv(ITEMREC.BIKOU, "")                        '印刷備考
        Call UniCode_Conv(ITEMREC.IRI_QTY, "")                      '印刷入り数
        Call UniCode_Conv(ITEMREC.JAN_CODE, "")                     'JANｺｰﾄﾞ
        Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")                   '品番読み替えｺｰﾄﾞ
        Call UniCode_Conv(ITEMREC.GOODS_KBN, "1")                   '商品化有無
        Call UniCode_Conv(ITEMREC.PACKING_NO, "")                   '個装箱№
        Call UniCode_Conv(ITEMREC.RANK, "")                         '現在ﾗﾝｸ
        Call UniCode_Conv(ITEMREC.NEW_RANK, "")                     '新ﾗﾝｸ
        Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")                  'ｸﾞﾘｯｸｽ棚番1
        Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")                  'ｸﾞﾘｯｸｽ棚番2
        Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")                  'ｸﾞﾘｯｸｽ棚番3
    
        Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")                 '品名E
        Call UniCode_Conv(ITEMREC.L_BIKOU, "")                      '備考
        Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")                '会社名
        Call UniCode_Conv(ITEMREC.L_KISHU1, "")                     '機種(1)
        Call UniCode_Conv(ITEMREC.L_KISHU2, "")                     '機種(2)
        Call UniCode_Conv(ITEMREC.L_KISHU3, "")                     '機種(3)
        Call UniCode_Conv(ITEMREC.L_PAPER, "")                      '紙
        Call UniCode_Conv(ITEMREC.L_PLASTIC, "")                    'ﾌﾟﾗｽﾁｯｸ
        Call UniCode_Conv(ITEMREC.L_URIKIN1, "")                    '価格(1)
        Call UniCode_Conv(ITEMREC.L_URIKIN2, "")                    '価格(2)
        Call UniCode_Conv(ITEMREC.L_URIKIN3, "")                    '価格(3)
        Call UniCode_Conv(ITEMREC.L_LABEL, "")                      '適用機種ﾗﾍﾞﾙ
        Call UniCode_Conv(ITEMREC.L_MAISU, "")                      'ﾗﾍﾞﾙ枚数
        Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")                '適用機種備考
        Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")                '作業指示
        Call UniCode_Conv(ITEMREC.L_BIKOU3, "")                     '備考(3)
        Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")                '事業部名
        Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")                    '入り数
        Call UniCode_Conv(ITEMREC.L_TANA1, "")                      '棚番(1)
        Call UniCode_Conv(ITEMREC.L_TANA2, "")                      '棚番(2)
        
        Call UniCode_Conv(ITEMREC.S_TANTO, "")                      '収単／担当者
        Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_ON)        'ﾗﾍﾞﾙ貼り付け
        
        Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, "")                '工数原価   2008.02.14
        Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "")                '工数売価   2008.02.14
        Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, "")             '工数設定日 2008.02.14
        
        Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, "")               '資材原価   2008.02.14
        Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "")               '資材売価   2008.02.14
        Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, "")            '資材設定日 2008.02.14
        
        
        
        Call UniCode_Conv(ITEMREC.SE_USOU_F, "")                    '輸送箱　出力ﾌﾗｸﾞ   2008.02.14

        Call UniCode_Conv(ITEMREC.USE_TAPE_KIND, "")                '使用テープ種類     2008.02.14
        Call UniCode_Conv(ITEMREC.USE_TAPE_LNG, "")                 '使用テープ長       2008.02.14

        Call UniCode_Conv(ITEMREC.H_TANA_MAKE, "")                  '棚番マーク         2008.04.02

        Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, "")                '請求単価　メモ     2008.04.15

        Call UniCode_Conv(ITEMREC.xGENSANKOKU, "")                  '原産国             2008.06.11-->2009.07.16 未使用

        Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, "")                '外装単価 9(8)V99   2008.06.12
        Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, "")             'PPSC加工単価9(8)   2008.06.12
        Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, "")               'BU加工単価9(8)     2008.06.12

        Call UniCode_Conv(ITEMREC.SEI_LOT, "")                      '生産ロット         2008.07.07
        Call UniCode_Conv(ITEMREC.SEI_RATE, "")                     '分レート           2008.07.07
        Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "")                  '集合梱包           2008.07.07

        Call UniCode_Conv(ITEMREC.SEI_TANKA_TANTO, "")              '単価設定担当者     2008.07.09

        Call UniCode_Conv(ITEMREC.SHIMUKE_CODE, "")                 '仕向け先           2008.07.09

        Call UniCode_Conv(ITEMREC.SEI_KBN, "")                      '請求区分           2008.07.16

        Call UniCode_Conv(ITEMREC.SEI_LABEL_QTY, "")                'ラベル貼り枚数     2008.07.19

        Call UniCode_Conv(ITEMREC.SEI_SZI_CNT, "")                  '資材件数     　    2008.08.20追加
        Call UniCode_Conv(ITEMREC.SEI_DKN_CNT, "")                  '同梱件数           2008.08.20追加

        For i = 0 To 9                                              '2008.09.19
            Call UniCode_Conv(ITEMREC.BEF_KOUTEI(i).BEF_KOUTEI, "")
            Call UniCode_Conv(ITEMREC.MAIN_KOUTEI(i).MAIN_KOUTEI, "")
            Call UniCode_Conv(ITEMREC.AFT_KOUTEI(i).AFT_KOUTEI, "")
        Next i
    


        Call UniCode_Conv(ITEMREC.SE_IO_TANKA_No, "")               '棚区分             200.09.19

        Call UniCode_Conv(ITEMREC.STAT, "")                         '状態区分           2009.01.21

        Call UniCode_Conv(ITEMREC.INSP_MESSAGE, "")                 '出荷検品ﾒｯｾｰｼﾞ     2009.04.17

        Call UniCode_Conv(ITEMREC.S_SEIKYU_F, "")                   '商品化請求ﾌﾗｸﾞ     2009.04.28
    
        Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "")            '商品化　工数売価   2009.06.02
        Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "")           '商品化　資材売価   2009.06.02
        Call UniCode_Conv(ITEMREC.BEF_S_GAISO_TANKA, "")            '外装単価 9(8)V99   2009.06.02
        Call UniCode_Conv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, "")         'PPSC加工単価9(8)   2009.06.02
        Call UniCode_Conv(ITEMREC.BEF_S_BU_KAKO_KOSU, "")           'BU加工単価9(8)     2009.06.02
    
        Call UniCode_Conv(ITEMREC.M_BIKOU, "")                      '見積書備考         2009.06.02
        Call UniCode_Conv(ITEMREC.SHIYOU_NO, "")                    '仕様書№           2009.06.02
        Call UniCode_Conv(ITEMREC.MITSUMORI_KBN, "")                '見積り区分         2009.06.02
        Call UniCode_Conv(ITEMREC.TANKA_KIRIKAE_DT, "")             '単価切替日付       2009.06.02
        Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, "")                  '切替区分           2009.06.02
    
        Call UniCode_Conv(ITEMREC.GENSANKOKU, "")                   '原産国             '2009.07.16
        
        
        
        
        
        
'-------    2010.10.04
        Call UniCode_Conv(ITEMREC.PLUS_KOUSU, "")                   'プラス分工数       2009.09.17
        Call UniCode_Conv(ITEMREC.KUTI_SU, "")                      '口数               2010.01.18
        Call UniCode_Conv(ITEMREC.KONPOU_F, "")                     '梱包区分           2010.01.18
    
        Call UniCode_Conv(ITEMREC.SAI_SU, "")                       '才数               2010.01.18
    
    
    
        Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, "")              '取込み時原産国     2010.07.20
        Call UniCode_Conv(ITEMREC.TORI_GEN_GENSANKOKU, "")          '取込み時原産国表示 2010.07.20
        Call UniCode_Conv(ITEMREC.TORI_SHIIRE_WORK_CENTER, "")      '仕入ﾜｰｸセンター    2010.07.20
        
    
    
        Call UniCode_Conv(ITEMREC.KANKYO_KBN, "")                   '環境種類区分       2010.07.27
        Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, "")                '環境種類区分適用開始 2010.07.27
        Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, "")             '環境種類区分数量   2010.07.27
    
        Call UniCode_Conv(ITEMREC.BEF_L_LABEL, "")                       '''''
    
        Call UniCode_Conv(ITEMREC.BEF_1_L_PAPER, "")                '           紙
        Call UniCode_Conv(ITEMREC.BEF_1_L_PLASTIC, "")              '           プラスチック
        Call UniCode_Conv(ITEMREC.BEF_2_L_PAPER, "")                '           紙
        Call UniCode_Conv(ITEMREC.BEF_2_L_PLASTIC, "")              '           プラスチック
        Call UniCode_Conv(ITEMREC.BEF_3_L_PAPER, "")                '           紙
        Call UniCode_Conv(ITEMREC.BEF_3_L_PLASTIC, "")              '           プラスチック
        Call UniCode_Conv(ITEMREC.BEF_4_L_PAPER, "")                '           紙
        Call UniCode_Conv(ITEMREC.BEF_4_L_PLASTIC, "")              '           プラスチック
        Call UniCode_Conv(ITEMREC.BEF_LAST_L_PAPER, "")             '           紙
        Call UniCode_Conv(ITEMREC.BEF_LAST_L_PLASTIC, "")           '           プラスチック


        Call UniCode_Conv(ITEMREC.BIKOU20, "")
'-------    2010.09.04
        
        Call UniCode_Conv(ITEMREC.SHIIRE_BIKOU, "")                 '仕入備考   2018.04.19
        
        
        
        Call UniCode_Conv(ITEMREC.INS_TANTO, "PM030")               '追加　担当者　     2009.01.21
        Call UniCode_Conv(ITEMREC.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))  '追加　日時         2009.01.21

        Call UniCode_Conv(ITEMREC.UPD_TANTO, "")                    '更新　担当者　     2005.11.15
        Call UniCode_Conv(ITEMREC.UPD_DATETIME, "")                 '更新　日時         2005.11.15
        
        
        
        Call UniCode_Conv(ITEMREC.FILLER, "")                       'Filler
    
    End If
    
    
    Call UniCode_Conv(ITEMREC.HIN_NAME, Text1(ptxHIN_NAME).Text)
    
    
    Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, Text1(ptxG_SHIIRE_KBN).Text)                '仕入区分
    Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, Text1(ptxG_HANBAI_KBN).Text)                '販売区分
    Call UniCode_Conv(ITEMREC.G_SYUSHI, Text1(ptxG_SYUSHI).Text)                        '収支単位
    If Check1(pchkG_KUMITATE).Value = vbChecked Then                                    '組立製品
        Call UniCode_Conv(ITEMREC.G_KUMITATE, P_ASSEMBLY_ON)
    Else
        Call UniCode_Conv(ITEMREC.G_KUMITATE, P_ASSEMBLY_OFF)
    End If
        
    If Trim(Text1(ptxG_ZEN_ZAIKO_KIN).Text) = "" Then                                   '前月在庫金額
        Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")
    Else
        If CLng(Text1(ptxG_ZEN_ZAIKO_KIN).Text) < 0 Then
            Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, Format(CLng(Text1(ptxG_ZEN_ZAIKO_KIN).Text), "0000000"))
        Else
            Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, Format(CLng(Text1(ptxG_ZEN_ZAIKO_KIN).Text), "00000000"))
        End If
    End If
    
    
    If Trim(Text1(ptxG_ZEN_ZAIKO_QTY).Text) = "" Then                                   '前月在庫数量
        Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "")
    Else
        If CLng(Text1(ptxG_ZEN_ZAIKO_QTY).Text) < 0 Then
        
            Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, Format(CLng(Text1(ptxG_ZEN_ZAIKO_QTY).Text), "0000000"))
        Else
            Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, Format(CLng(Text1(ptxG_ZEN_ZAIKO_QTY).Text), "00000000"))
        End If
    End If
    
    
    
    If Trim(Text1(ptxG_ST_URITAN).Text) = "" Then                                       '標準売価
        Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")
    Else
        Call UniCode_Conv(ITEMREC.G_ST_URITAN, Format(CDbl(Text1(ptxG_ST_URITAN).Text), "00000000.00"))
    End If
    If Trim(Text1(ptxG_ST_URITAN_DT).Text) = "" Then                                   '標準売価設定日
        Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, Format(Text1(ptxG_ST_URITAN_DT).Text, "YYYYMMDD"))
    End If
    
    If Trim(Text1(ptxG_ST_SHITAN).Text) = "" Then                                       '標準原価
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")
    Else
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN, Format(CDbl(Text1(ptxG_ST_SHITAN).Text), "00000000.00"))
    End If
    If Trim(Text1(ptxG_ST_SHITAN_DT).Text) = "" Then                                   '標準原価設定日
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, Format(Text1(ptxG_ST_SHITAN_DT).Text, "YYYYMMDD"))
    End If
    
    If Trim(Text1(ptxHOJYU_P).Text) = "" Then                                           '危険在庫
        Call UniCode_Conv(ITEMREC.HOJYU_P, "")
    Else
        Call UniCode_Conv(ITEMREC.HOJYU_P, Format(CLng(Text1(ptxHOJYU_P).Text), "00000000"))
    End If
        
        
    Call UniCode_Conv(ITEMREC.G_SHIZAI_KBN, Right(Combo1(pcmbG_SHIZAI_KBN).Text, 1))    '資材区分
        
    If Check1(pchkG_LABEL_NON).Value = vbChecked Then                                   'ラベル貼り
        Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_OFF)
    Else
        Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_ON)
    End If
        
    If Check1(pchkZAIKO_F).Value = vbUnchecked Then                                     '在庫管理対象
        Call UniCode_Conv(ITEMREC.ZAIKO_F, P_ZAIKO_F_ON)
    Else
        Call UniCode_Conv(ITEMREC.ZAIKO_F, P_ZAIKO_F_OFF)
    End If
        
        
        
    If Check1(pchkZAIKO_CLR_F).Value = vbChecked Then                                   '棚卸表　在庫数非表示   2012.12.13
        Call UniCode_Conv(ITEMREC.ZAIKO_CLR_F, "1")
    Else
        Call UniCode_Conv(ITEMREC.ZAIKO_CLR_F, "")
    End If
        
        
        
        
                                    '才数   2008.02.14
    If IsNumeric(Text1(ptxSAI_SU).Text) Then
'        Call UniCode_Conv(ITEMREC.SAI_SU, Format(CDbl(Text1(ptxSAI_SU).Text), "0.0"))
        Call UniCode_Conv(ITEMREC.SAI_SU, Format(CDbl(Text1(ptxSAI_SU).Text), "0.00"))
    Else
        Call UniCode_Conv(ITEMREC.SAI_SU, "")
    End If
                                    '形式    2008.02.14
    Call UniCode_Conv(ITEMREC.D_KEISHIKI, Text1(ptxD_KEISHIKI).Text)
                                    '材質   2008.02.14
    Call UniCode_Conv(ITEMREC.D_MATERIAL, Text1(ptxD_MATERIAL).Text)
                                    'ﾀﾞﾝﾎﾞｰﾙ厚さ    2008.02.14
    Call UniCode_Conv(ITEMREC.D_THICKNESS, Text1(ptxD_THICKNESS).Text)
                                    
                                    'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ(W) 2008.02.14
    Call UniCode_Conv(ITEMREC.D_SIZE_W, Text1(ptxD_SIZE_W).Text)
                                    'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ(D) 2008.02.14
    Call UniCode_Conv(ITEMREC.D_SIZE_D, Text1(ptxD_SIZE_D).Text)
                                    'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ(H) 2008.02.14
    Call UniCode_Conv(ITEMREC.D_SIZE_H, Text1(ptxD_SIZE_H).Text)
                                    
                                    '印刷する/しない 2008.02.14
    Call UniCode_Conv(ITEMREC.D_PRINT, Text1(ptxD_PRINT).Text)
                                    '商品化工数 2008.02.14
    If Trim(Text1(ptxS_KOUSU).Text) <> "" Then
        Call UniCode_Conv(ITEMREC.S_KOUSU, Format(CDbl(Text1(ptxS_KOUSU).Text), "00000000"))
    Else
        Call UniCode_Conv(ITEMREC.S_KOUSU, "")
    End If
                                    
                                    '集合梱包 2008.07.16
    If Trim(Text1(ptxSEI_SYU_KON).Text) <> "" Then
        Call UniCode_Conv(ITEMREC.SEI_SYU_KON, Format(CDbl(Text1(ptxSEI_SYU_KON).Text), "000000"))
    Else
        Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "")
    End If
                                    '輸送箱 2008.02.14
    Call UniCode_Conv(ITEMREC.SE_USOU_F, Text1(ptxSE_USOU_F).Text)
        
                                    '使用ﾃｰﾌﾟ種類 2008.02.14
    Call UniCode_Conv(ITEMREC.USE_TAPE_KIND, Text1(ptxUSE_TAPE_KIND).Text)
                                    '使用ﾃｰﾌﾟ長 2008.02.14
    Call UniCode_Conv(ITEMREC.USE_TAPE_LNG, Text1(ptxUSE_TAPE_LNG).Text)
                                    '請求区分 2008.07.16
    Call UniCode_Conv(ITEMREC.SEI_KBN, Text1(ptxSEI_KBN).Text)
                                    
                                    
    If Trim(Text1(ptxST_SOKO).Text) = "" And Trim(Text1(ptxST_RETU).Text) = "" And Trim(Text1(ptxST_REN).Text) = "" And Trim(Text1(ptxST_DAN).Text) = "" Then
    
    
        Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
    
    
    Else
        If Trim(Text1(ptxST_SOKO).Text) = Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) And Trim(Text1(ptxST_RETU).Text) = Trim(StrConv(ITEMREC.ST_RETU, vbUnicode)) And Trim(Text1(ptxST_REN).Text) = Trim(StrConv(ITEMREC.ST_REN, vbUnicode)) And Trim(Text1(ptxST_DAN).Text) = Trim(StrConv(ITEMREC.ST_DAN, vbUnicode)) Then
        Else
            Call UniCode_Conv(ITEMREC.ST_SET_DT, Format(Now, "YYYYMMDD"))
        End If
    End If
                                    
                                    
                                    
                                    
                                    
                                    
                                    
                                    '標準棚番 2009.09.01
    Call UniCode_Conv(ITEMREC.ST_SOKO, Text1(ptxST_SOKO).Text)
                                    '標準棚番 2009.09.01
    Call UniCode_Conv(ITEMREC.ST_RETU, Text1(ptxST_RETU).Text)
                                    '標準棚番 2009.09.01
    Call UniCode_Conv(ITEMREC.ST_REN, Text1(ptxST_REN).Text)
                                    '標準棚番 2009.09.01
    Call UniCode_Conv(ITEMREC.ST_DAN, Text1(ptxST_DAN).Text)
        
        
        
        
        
                                    '口数 2010.01.18
    Call UniCode_Conv(ITEMREC.KUTI_SU, Text1(ptxKUTI_SU).Text)
                                    '梱包F 2010.01.18
    Call UniCode_Conv(ITEMREC.KONPOU_F, Text1(ptxKONPOU_F).Text)
        
        
        
        
        
        
    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).CODE, Text1(ptxSHI_CODE1).Text)           '仕入先(1)
    If Trim(Text1(ptxSHI_TANKA1).Text) = "" Then                                        '仕入単価(1)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA, Format(CDbl(Text1(ptxSHI_TANKA1).Text), "00000000.00"))
    End If
    If Trim(Text1(ptxSHI_TANKA_DT1).Text) = "" Then                                     '仕入単価設定日(1)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA_DT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA_DT, Format(Text1(ptxSHI_TANKA_DT1).Text, "YYYYMMDD"))
    End If
    If Trim(Text1(ptxSHI_LOT1).Text) = "" Then                                          'ﾛｯﾄ(1)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LOT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LOT, Format(CLng(Text1(ptxSHI_LOT1).Text), "00000000"))
    End If
    If Trim(Text1(ptxSHI_LEAD_TIME1).Text) = "" Then                                    'ﾘｰﾄﾞﾀｲﾑ(1)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LEAD_TIME, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LEAD_TIME, Format(CLng(Text1(ptxSHI_LEAD_TIME1).Text), "000"))
    End If
    If Trim(Text1(ptxSHI_LAST_ORDER_DT1).Text) = "" Then                                '前回注文日(1)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_DT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_DT, Format(Text1(ptxSHI_LAST_ORDER_DT1).Text, "YYYYMMDD"))
    End If
    If Trim(Text1(ptxSHI_LAST_ORDER_QTY1).Text) = "" Then                               '前回注文数(1)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_QTY, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).LAST_ORDER_QTY, Format(CLng(Text1(ptxSHI_LAST_ORDER_QTY1).Text), "00000000"))
    End If
                
    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).CODE, Text1(ptxSHI_CODE2).Text)           '仕入先(2)
    If Trim(Text1(ptxSHI_TANKA2).Text) = "" Then                                        '仕入単価(2)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).TANKA, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).TANKA, Format(CDbl(Text1(ptxSHI_TANKA2).Text), "00000000.00"))
    End If
    If Trim(Text1(ptxSHI_TANKA_DT2).Text) = "" Then                                     '仕入単価設定日(2)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).TANKA_DT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).TANKA_DT, Format(Text1(ptxSHI_TANKA_DT2).Text, "YYYYMMDD"))
    End If
    If Trim(Text1(ptxSHI_LOT2).Text) = "" Then                                          'ﾛｯﾄ(2)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).LOT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).LOT, Format(CLng(Text1(ptxSHI_LOT2).Text), "00000000"))
    End If
    If Trim(Text1(ptxSHI_LEAD_TIME2).Text) = "" Then                                    'ﾘｰﾄﾞﾀｲﾑ(2)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).LEAD_TIME, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).LEAD_TIME, Format(CLng(Text1(ptxSHI_LEAD_TIME2).Text), "000"))
    End If
    If Trim(Text1(ptxSHI_LAST_ORDER_DT2).Text) = "" Then                                '前回注文日(2)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_DT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_DT, Format(Text1(ptxSHI_LAST_ORDER_DT2).Text, "YYYYMMDD"))
    End If
    If Trim(Text1(ptxSHI_LAST_ORDER_QTY2).Text) = "" Then                               '前回注文数(2)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_QTY, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(1).LAST_ORDER_QTY, Format(CLng(Text1(ptxSHI_LAST_ORDER_QTY2).Text), "00000000"))
    End If
                
    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).CODE, Text1(ptxSHI_CODE3).Text)           '仕入先(3)
    If Trim(Text1(ptxSHI_TANKA3).Text) = "" Then                                        '仕入単価(3)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).TANKA, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).TANKA, Format(CDbl(Text1(ptxSHI_TANKA3).Text), "00000000.00"))
    End If
    If Trim(Text1(ptxSHI_TANKA_DT3).Text) = "" Then                                     '仕入単価設定日(3)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).TANKA_DT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).TANKA_DT, Format(Text1(ptxSHI_TANKA_DT3).Text, "YYYYMMDD"))
    End If
    If Trim(Text1(ptxSHI_LOT3).Text) = "" Then                                           'ﾛｯﾄ(3)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).LOT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).LOT, Format(CLng(Text1(ptxSHI_LOT3).Text), "00000000"))
    End If
    If Trim(Text1(ptxSHI_LEAD_TIME3).Text) = "" Then                                    'ﾘｰﾄﾞﾀｲﾑ(3)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).LEAD_TIME, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).LEAD_TIME, Format(CLng(Text1(ptxSHI_LEAD_TIME3).Text), "000"))
    End If
    If Trim(Text1(ptxSHI_LAST_ORDER_DT3).Text) = "" Then                                '前回注文日(3)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_DT, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_DT, Format(Text1(ptxSHI_LAST_ORDER_DT3).Text, "YYYYMMDD"))
    End If
    If Trim(Text1(ptxSHI_LAST_ORDER_QTY3).Text) = "" Then                               '前回注文数(3)
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_QTY, "")
    Else
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(2).LAST_ORDER_QTY, Format(CLng(Text1(ptxSHI_LAST_ORDER_QTY3).Text), "00000000"))
    End If
    
        
    If Trim(Text1(ptxLAST_SYU_DT).Text) = "" Then                                       '最終出荷日
        Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")
    Else
        Call UniCode_Conv(ITEMREC.LAST_SYU_DT, Format(Text1(ptxLAST_SYU_DT).Text, "YYYYMMDD"))
    End If
    
    If Trim(Text1(ptxG_LAST_SYUKA_QTY).Text) = "" Then                                  '最終出荷数
        Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, "00000000")
    Else
        Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, Format(CLng(Text1(ptxG_LAST_SYUKA_QTY).Text), "00000000"))
    End If
    
    
                                                                                        '最新仕入先     2007.05.29
    Call UniCode_Conv(ITEMREC.LAST_CODE, Text1(ptxLAST_CODE).Text)
    
    If Trim(Text1(ptxLAST_TANKA).Text) = "" Then                                        '最新仕入単価   2007.05.29
        Call UniCode_Conv(ITEMREC.LAST_TANKA, "00000000.00")
    Else
        Call UniCode_Conv(ITEMREC.LAST_TANKA, Format(CDbl(Text1(ptxLAST_TANKA).Text), "00000000.00"))
    End If
    
    Call UniCode_Conv(ITEMREC.SHIIRE_BIKOU, Text1(ptxSHIIRE_BIKOU).Text)                '仕入備考　2018.04.19
    
    
    
    
    Call UniCode_Conv(ITEMREC.UPD_TANTO, App.EXEName)                                   '更新担当者ｺｰﾄﾞ
                                                                                        '更新日時
    Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    Do
        
        DoEvents
        
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Update_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Exit Function
        End Select
    
    Loop
    
    
    Update_Proc = False


End Function
Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   品目マスタ削除
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer

    Delete_Proc = True
    
    '品目マスタ　読み込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                Delete_Proc = False
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Delete_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                Exit Function
        
        End Select
    
    
    Loop
    
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpDelete, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Delete_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, BtOpDelete, "品目マスタ")
                Exit Function
        End Select
    Loop


    Delete_Proc = False


End Function


Private Sub Check1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    Select Case Index
        Case pcmbG_SHIIRE       '仕入区分
            Text1(ptxG_SHIIRE_KBN).Text = Left(Right(Combo1(pcmbG_SHIIRE).Text, 3), 2)
        Case pcmbG_HANBAI       '販売区分
            Text1(ptxG_HANBAI_KBN).Text = Left(Right(Combo1(pcmbG_HANBAI).Text, 3), 2)
        Case pcmbG_SYUSHI       '収支単位
            Text1(ptxG_SYUSHI).Text = Right(Combo1(pcmbG_SYUSHI).Text, 3)
        Case pcmbSHIIRE1        '仕入先(1)
            Text1(ptxSHI_CODE1).Text = Right(Combo1(pcmbSHIIRE1).Text, 5)
        Case pcmbSHIIRE2        '仕入先(2)
            Text1(ptxSHI_CODE2).Text = Right(Combo1(pcmbSHIIRE2).Text, 5)
        Case pcmbSHIIRE3        '仕入先(3)
            Text1(ptxSHI_CODE3).Text = Right(Combo1(pcmbSHIIRE3).Text, 5)
        Case pcmbLAST_CODE      '最新仕入先     2007.05.28
            Text1(ptxLAST_CODE).Text = Right(Combo1(pcmbLAST_CODE).Text, 5)
    
    End Select
    
    Call Tab_Ctrl(Shift)        '移動

End Sub


Private Sub Command1_Click(Index As Integer)

Dim ans As Integer
Dim i   As Integer


    Select Case Index
        Case P_CMD_Upd                      '更新
            
'2010.01.18            For i = ptxHIN_GAI To ptxST_DAN
            For i = ptxHIN_GAI To ptxKONPOU_F
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
            
            Beep
            ans = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Update_Proc() Then
                    PM000302.Visible = False
                    INIT_FLG = False
                End If
            Else
                Exit Sub
            End If
            PM000302.Visible = False
            INIT_FLG = False
                    
        
        
        Case P_CMD_DEL                      '削除
            ans = MsgBox("削除しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Delete_Proc() Then
                    PM000302.Visible = False
                    INIT_FLG = False
                End If
            Else
                Exit Sub
            End If
            PM000302.Visible = False
            INIT_FLG = False
        Case P_CMD_DSP                      '検索/表示
        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        Case P_CMD_PRT                      '印刷
        
        Case P_CMD_End                      '終了
            PM000302.Visible = False
            INIT_FLG = False
    End Select

End Sub

Private Sub Command2_Click()
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "資材マスタメンテナンス 画面印刷を開始しました ", Me.hwnd, 0)


Call Form_HCopy_Win7_NEW(Picture1, vbPRPSA4, vbPRORLandscape)       '2018.11.21

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "資材マスタメンテナンス 画面印刷を終了しました ", Me.hwnd, 0)


End Sub

Private Sub Form_Activate()
    
Dim i       As Integer
Dim CODE    As String
    
    If INIT_FLG Then
        Exit Sub
    End If


    Select Case G_SCREEN_FLG
        Case G_SCREEN_INS       '新規
                
            Text1(ptxHIN_GAI).BackColor = G_INPUT_OK
            Text1(ptxHIN_GAI).TabStop = True
            Text1(ptxHIN_GAI).Locked = False
                
'            Text1(ptxHIN_NAME).BackColor = G_INPUT_OK
'            Text1(ptxHIN_NAME).TabStop = True
'            Text1(ptxHIN_NAME).Locked = False
                
                
            For i = ptxHIN_GAI To ptxKONPOU_F
                Text1(i).Text = ""
            Next i
                
            Text1(ptxSHIIRE_BIKOU).Text = ""    '2018.04.19
                
            For i = pcmbG_SHIIRE To pcmbLAST_CODE
            
                Combo1(i).ListIndex = -1
            Next i
                
'2012.12.13            For i = pchkG_KUMITATE To pchkZAIKO_F
            For i = pchkG_KUMITATE To pchkZAIKO_CLR_F   '2012.12.13
                Check1(i).Value = vbUnchecked
            Next i
                
                
            lblSIZE.Caption = ""
            lblIns_DateTime.Caption = ""
            lblUpd_DateTime.Caption = ""
                
                
                
                
            Text1(ptxHIN_GAI).SetFocus
                
                
                
        
        Case G_SCREEN_UPD       '更新
    
                
    
    
            Text1(ptxHIN_GAI).BackColor = G_INPUT_NG
            Text1(ptxHIN_GAI).TabStop = False
            Text1(ptxHIN_GAI).Locked = True
    
'            Text1(ptxHIN_NAME).BackColor = G_INPUT_NG
'            Text1(ptxHIN_NAME).TabStop = False
'            Text1(ptxHIN_NAME).Locked = True
    
            
            CODE = PM000301.txSEL_KEY
            
            If Item_Disp_Proc(CODE) Then
                Unload Me
            End If
    
            Text1(ptxHIN_NAME).SetFocus
    
    End Select


    INIT_FLG = True

End Sub

Private Sub Form_DblClick()
'    PrintForm
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

Dim com     As Integer
Dim sts     As Integer


    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "資材マスタメンテナンス", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)


    Combo1(pcmbNAIGAI).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo1(pcmbNAIGAI).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    Combo1(pcmbNAIGAI).ListIndex = 0
    
    '仕入区分のセット
    If Code_Set_Proc(pcmbG_SHIIRE, P_KBN01_CD, 0) Then
        Unload Me
    End If
    
    
    '販売区分のセット
    If Code_Set_Proc(pcmbG_HANBAI, P_KBN02_CD, 0) Then
        Unload Me
    End If
    
    
    
    '収支内容のセット
    If Code_Set_Proc(pcmbG_SYUSHI, P_KBN03_CD, 1) Then
        Unload Me
    End If
    
    '資材区分のセット
    If Code_Set_Proc(pcmbG_SHIZAI_KBN, P_KBN08_CD, 1) Then
        Unload Me
    End If
    
    
    
    '仕入先のセット
    Combo1(pcmbSHIIRE1).Clear
    Combo1(pcmbSHIIRE2).Clear
    Combo1(pcmbSHIIRE3).Clear
    
    Combo1(pcmbLAST_CODE).Clear
    
    
    com = BtOpGetFirst
    
    Do
        DoEvents
    
        sts = BTRV(com, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
            
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "受払先マスタ")
                Unload Me
        End Select
        
        Combo1(pcmbSHIIRE1).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode) & " " & _
                                    StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        Combo1(pcmbSHIIRE2).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode) & " " & _
                                    StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        Combo1(pcmbSHIIRE3).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode) & " " & _
                                    StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        
        '最新仕入先 2007.05.28
        Combo1(pcmbLAST_CODE).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode) & " " & _
                                    StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        
        
        com = BtOpGetNext
    
    Loop
    
    
    INIT_FLG = False
    
    
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  前画面に戻す    2016.01.27
'Dim sts As Integer
'
'                                            '品目マスタＣＬＯＳＥ
'    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'    If sts Then
'        If sts <> BtErrNoOpen Then
'            Call File_Error(sts, BtOpClose, "品目マスタ")
'        End If
'    End If
'
'
'                                            '受払先マスタＣＬＯＳＥ
'    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
'    If sts Then
'        If sts <> BtErrNoOpen Then
'            Call File_Error(sts, BtOpClose, "受払先マスタ")
'        End If
'    End If
'
'                                            'コードマスタＣＬＯＳＥ
'    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
'    If sts Then
'        If sts <> BtErrNoOpen Then
'            Call File_Error(sts, BtOpClose, "コードマスタ")
'        End If
'    End If
'    sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
'    If sts Then
'        Call File_Error(sts, BtOpReset, "")
'    End If
'    Set PM000301 = Nothing
'    Set PM000302 = Nothing
'
'    End
'

    PM000302.Visible = False
    INIT_FLG = False



'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  前画面に戻す    2016.01.27

End Sub


Private Sub Text1_GotFocus(Index As Integer)
    
    
'>>>>>  2018.04.09
    Select Case Index
    
        Case ptxG_ST_URITAN, ptxG_ST_SHITAN, ptxSHI_TANKA1, ptxSHI_TANKA2, ptxSHI_TANKA3
            svTANKA = Text1(Index).Text
    
    End Select
'>>>>>  2018.04.09
    
    
    
    
    
    
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    If Error_Check_Proc(Index) Then     'エラーチェック
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動

End Sub

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

Private Sub Text1_LostFocus(Index As Integer)

'>>>>>>>>>>>>>>>>>>>>>>>>>  2016.01.26
    Select Case Index
        Case ptxST_SOKO
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)
                    
        Case ptxSHI_CODE1
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)
        Case ptxSHI_CODE2
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)
        Case ptxSHI_CODE3
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)
    
        Case ptxLAST_CODE
            Text1(Index).Text = StrConv(Text1(Index).Text, vbUpperCase)
    
    
'>>>>>  2018.04.09
        Case ptxG_ST_URITAN, ptxG_ST_SHITAN, ptxSHI_TANKA1, ptxSHI_TANKA2, ptxSHI_TANKA3
            If svTANKA <> Text1(Index).Text Then
                If Trim(Text1(Index).Text) = "" Then
                    Text1(Index + 1).Text = ""
                Else
                    Text1(Index + 1).Text = Format(Now, "YYYY/MM/DD")
                End If
            End If
'>>>>>  2018.04.09
    
    
    
    End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>  2016.01.26


End Sub
