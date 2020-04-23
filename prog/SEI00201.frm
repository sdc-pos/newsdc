VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form SEI00201 
   Caption         =   "[請求システム]請求書作成処理"
   ClientHeight    =   10755
   ClientLeft      =   2025
   ClientTop       =   1455
   ClientWidth     =   17865
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
   ScaleHeight     =   10755
   ScaleWidth      =   17865
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   7
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   22
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   3720
      Width           =   1845
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   21
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   3360
      Width           =   1845
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   20
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   3000
      Width           =   1845
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   19
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   2640
      Width           =   1845
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   18
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   2280
      Width           =   1845
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   17
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   3720
      Width           =   1845
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   14
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   2640
      Width           =   1845
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   15
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   3000
      Width           =   1845
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   16
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   3360
      Width           =   1845
   End
   Begin VB.CommandButton Command1 
      Caption         =   "科目振替明細表"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   6405
      TabIndex        =   3
      Top             =   120
      Width           =   2430
   End
   Begin VB.Frame Frame1 
      Caption         =   "表示条件"
      Height          =   735
      Left            =   4515
      TabIndex        =   39
      Top             =   960
      Width           =   4425
      Begin VB.CheckBox Check1 
         Caption         =   "良返明細"
         Height          =   255
         Index           =   2
         Left            =   2940
         TabIndex        =   42
         Top             =   360
         Width           =   1380
      End
      Begin VB.CheckBox Check1 
         Caption         =   "入庫明細"
         Height          =   255
         Index           =   1
         Left            =   1575
         TabIndex        =   41
         Top             =   360
         Width           =   1380
      End
      Begin VB.CheckBox Check1 
         Caption         =   "出荷明細"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   40
         Top             =   360
         Value           =   1  'ﾁｪｯｸ
         Width           =   1380
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終  了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   13125
      TabIndex        =   6
      Top             =   120
      Width           =   2010
   End
   Begin VB.CommandButton Command1 
      Caption         =   "良返明細表"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   11025
      TabIndex        =   5
      Top             =   120
      Width           =   2010
   End
   Begin VB.CommandButton Command1 
      Caption         =   "出荷明細表"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4305
      TabIndex        =   2
      Top             =   120
      Width           =   2010
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   12
      Left            =   13695
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   2640
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   11
      Left            =   13695
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   2280
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   10
      Left            =   13695
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   1920
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   8
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   6
      Left            =   1995
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3720
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   5
      Left            =   1995
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3360
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   4
      Left            =   1995
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3000
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   3
      Left            =   1995
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2640
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   2
      Left            =   1995
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2280
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   1
      Left            =   2940
      TabIndex        =   14
      Top             =   1320
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   0
      Left            =   1365
      TabIndex        =   12
      Top             =   1320
      Width           =   1380
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   1365
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   9
      Top             =   840
      Width           =   2220
   End
   Begin VB.CommandButton Command1 
      Caption         =   "入庫明細表"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   8925
      TabIndex        =   4
      Top             =   120
      Width           =   2010
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   7
      Top             =   10440
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton Command1 
      Caption         =   "表  紙"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2205
      TabIndex        =   1
      Top             =   120
      Width           =   2010
   End
   Begin VB.CommandButton Command1 
      Caption         =   "集　計"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2010
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   6255
      Left            =   105
      TabIndex        =   8
      Top             =   4200
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   11033
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "伝票日付"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "伝票№"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "出荷先/相手先"
      Columns(2).DataField=   ""
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
      Columns(5).Caption=   "数量"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "出荷工料 出庫"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "出荷工料 出荷"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "出荷工料 入庫"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "出荷工料 良返"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "商品化 工料"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "商品化 箱代"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   12
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=12"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2328"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2196"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1561"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1429"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3149"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3016"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2858"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2725"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2778"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=1561"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1429"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=2963"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2831"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=2963"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2831"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=514"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=2963"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2831"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=514"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(9).Width=2963"
      Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2831"
      Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=514"
      Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(51)=   "Column(10).Width=2963"
      Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=2831"
      Splits(0)._ColumnProps(54)=   "Column(10)._ColStyle=514"
      Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(56)=   "Column(11).Width=2963"
      Splits(0)._ColumnProps(57)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(11)._WidthInPix=2831"
      Splits(0)._ColumnProps(59)=   "Column(11)._ColStyle=514"
      Splits(0)._ColumnProps(60)=   "Column(11).Order=12"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageFooterFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "請求明細"
      AllowArrows     =   0   'False
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=ＭＳ Ｐゴシック"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=0,.fontsize=1125"
      _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ＭＳ ゴシック"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFFF80&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFFFF00&"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=87,.parent=1,.bgcolor=&H80FF80&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=96,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=89,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=91,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=94,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=95,.parent=10,.bgcolor=&HFFFF00&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=102,.parent=87"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=99,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=100,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=101,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=106,.parent=87"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=103,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=104,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=105,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=110,.parent=87"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=107,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=108,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=109,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=87"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=87"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=114,.parent=87,.alignment=1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=111,.parent=88"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=112,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=113,.parent=91"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=118,.parent=87,.alignment=1"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=115,.parent=88"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=116,.parent=89"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=117,.parent=91"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=16,.parent=87,.alignment=1"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=13,.parent=88"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=14,.parent=89"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=15,.parent=91"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=20,.parent=87,.alignment=1"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=17,.parent=88"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=18,.parent=89"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=19,.parent=91"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=24,.parent=87,.alignment=1"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=21,.parent=88"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=22,.parent=89"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=23,.parent=91"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=28,.parent=87,.alignment=1"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=25,.parent=88"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=26,.parent=89"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=27,.parent=91"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=32,.parent=87,.alignment=1"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=29,.parent=88"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=30,.parent=89"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=31,.parent=91"
      _StyleDefs(84)  =   "Named:id=33:Normal"
      _StyleDefs(85)  =   ":id=33,.parent=0"
      _StyleDefs(86)  =   "Named:id=34:Heading"
      _StyleDefs(87)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(88)  =   ":id=34,.wraptext=-1"
      _StyleDefs(89)  =   "Named:id=35:Footing"
      _StyleDefs(90)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(91)  =   "Named:id=36:Selected"
      _StyleDefs(92)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(93)  =   "Named:id=37:Caption"
      _StyleDefs(94)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(95)  =   "Named:id=38:HighlightRow"
      _StyleDefs(96)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(97)  =   "Named:id=39:EvenRow"
      _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(99)  =   "Named:id=40:OddRow"
      _StyleDefs(100) =   ":id=40,.parent=33"
      _StyleDefs(101) =   "Named:id=41:RecordSelector"
      _StyleDefs(102) =   ":id=41,.parent=34"
      _StyleDefs(103) =   "Named:id=42:FilterBar"
      _StyleDefs(104) =   ":id=42,.parent=33"
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   13
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   2280
      Width           =   1845
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   9
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '実線
      Caption         =   "工料"
      Height          =   375
      Index           =   7
      Left            =   8160
      TabIndex        =   27
      Top             =   2280
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '実線
      Caption         =   "商品化"
      Height          =   375
      Index           =   6
      Left            =   8160
      TabIndex        =   26
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '実線
      Caption         =   "数量"
      Height          =   375
      Index           =   13
      Left            =   6120
      TabIndex        =   49
      Top             =   1920
      Width           =   1845
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '実線
      Caption         =   "件数"
      Height          =   375
      Index           =   17
      Left            =   4320
      TabIndex        =   47
      Top             =   1920
      Width           =   1845
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '実線
      Caption         =   "請求金額"
      Height          =   375
      Index           =   12
      Left            =   12120
      TabIndex        =   37
      Top             =   2640
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '実線
      Caption         =   "消費税額"
      Height          =   375
      Index           =   11
      Left            =   12120
      TabIndex        =   35
      Top             =   2280
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '実線
      Caption         =   "請求合計"
      Height          =   375
      Index           =   10
      Left            =   12120
      TabIndex        =   33
      Top             =   1920
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '実線
      Caption         =   "小計"
      Height          =   375
      Index           =   9
      Left            =   8160
      TabIndex        =   31
      Top             =   3000
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '実線
      Caption         =   "箱代"
      Height          =   375
      Index           =   8
      Left            =   8160
      TabIndex        =   29
      Top             =   2640
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '実線
      Caption         =   "小計"
      Height          =   375
      Index           =   5
      Left            =   630
      TabIndex        =   20
      Top             =   3720
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '実線
      Caption         =   "良品返品"
      Height          =   375
      Index           =   4
      Left            =   630
      TabIndex        =   19
      Top             =   3360
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '実線
      Caption         =   "入庫"
      Height          =   375
      Index           =   3
      Left            =   630
      TabIndex        =   18
      Top             =   3000
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '実線
      Caption         =   "出荷"
      Height          =   375
      Index           =   2
      Left            =   630
      TabIndex        =   17
      Top             =   2640
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '実線
      Caption         =   "出庫"
      Height          =   375
      Index           =   1
      Left            =   630
      TabIndex        =   16
      Top             =   2280
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '実線
      Caption         =   "出荷工料"
      Height          =   375
      Index           =   0
      Left            =   630
      TabIndex        =   15
      Top             =   1920
      Width           =   3690
   End
   Begin VB.Label Label1 
      Caption         =   "～"
      Height          =   255
      Index           =   2
      Left            =   2730
      TabIndex        =   13
      Top             =   1440
      Width           =   225
   End
   Begin VB.Label Label1 
      Caption         =   "日付範囲"
      Height          =   255
      Index           =   1
      Left            =   315
      TabIndex        =   11
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "仕向先"
      Height          =   255
      Index           =   0
      Left            =   525
      TabIndex        =   10
      Top             =   960
      Width           =   750
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Index           =   0
      Begin VB.Menu SHORI 
         Caption         =   "集計"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "EXCEL(表紙)"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "EXCEL(出荷明細)"
         Index           =   2
      End
      Begin VB.Menu SHORI 
         Caption         =   "EXCEL(科目振替明細表)"
         Index           =   3
      End
      Begin VB.Menu SHORI 
         Caption         =   "EXCEL(入庫明細)"
         Index           =   4
      End
      Begin VB.Menu SHORI 
         Caption         =   "EXCEL(良品返品明細)"
         Index           =   5
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   6
      End
   End
End
Attribute VB_Name = "SEI00201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const pcmbSHIMUKE% = 0          '仕向け先

Private Const ptxS_Date% = 0            '日付範囲　開始
Private Const ptxE_Date% = 1            '日付範囲　終了

Private Const ptxGK_SYUKO_KOURYO% = 2   '出荷工料　出庫
Private Const ptxGK_SYUKA_KOURYO% = 3   '出荷工料　出荷
Private Const ptxGK_NYUKA_KOURYO% = 4   '出荷工料　入荷
Private Const ptxGK_RYOHEN_KOURYO% = 5  '出荷工料　良品
Private Const ptxGK_KOURYO% = 6         '出荷工料　小計

Private Const ptxGK_SYOHIN_KOURYO% = 7  '商品化　工料
Private Const ptxGK_SYOHIN_SHIZAI% = 8  '商品化　箱代
Private Const ptxGK_SYOHIN% = 9         '商品化　小計

Private Const ptxGK_SEIKYU% = 10        '請求合計
Private Const ptxGK_ZEI_KIN% = 11       '消費税額
Private Const ptxGK_SEIKYU_KIN% = 12    '請求金額

Private Const ptxGK_SYUKA_CNT% = 14     '出荷件数   2017.12.27
Private Const ptxGK_NYUKA_CNT% = 15     '入荷件数　 2017.12.27

Private Const ptxGK_SYUKA_QTY% = 19     '出荷数量   2017.12.27
Private Const ptxGK_NYUKA_QTY% = 20     '入荷数量   2017.12.27





Private Const pchkSYUKA% = 0            '出荷明細ＯＮ
Private Const pchkNYUKO% = 1            '入庫明細ＯＮ
Private Const pchkRYOHEN% = 2           '良返明細ＯＮ
    







Dim SEIKYU  As New XArrayDB

Private Const Min_Row% = 1              '最小行数

Dim Max_Row    As Integer               'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 11             '最大列数

Private Const ColSYUKA_YMD% = 0         '伝票日付
Private Const ColDEN_NO% = 1            '伝票№
Private Const ColMUKE_CODE% = 2         '出荷先

Private Const ColHIN_GAI% = 3           '品番
Private Const ColHIN_NAME% = 4          '品名



Private Const ColSURYO% = 5             '数量
Private Const ColSYUKO_KOURYO% = 6      '出荷工料　出庫分
Private Const ColSYUKA_KOURYO% = 7      '出荷工料　出荷分
Private Const ColNYUKA_KOURYO% = 8      '出荷工料　入庫分
Private Const ColRYOHEN_KOURYO% = 9     '出荷工料　良品返品
Private Const ColSYOHIN_KOURYO% = 10    '商品化　工料
Private Const ColSYOHIN_SHIZAI% = 11    '商品化　資材



Private GK_SYUKO_KOURYO     As Double   '出荷工料　出庫
Private GK_SYUKA_KOURYO     As Double   '出荷工料　出荷
Private GK_NYUKA_KOURYO     As Double   '出荷工料　入荷
Private GK_RYOHEN_KOURYO    As Double   '出荷工料　良品

Private GK_SYOHIN_KOURYO    As Double   '商品化　工料
Private GK_SYOHIN_SHIZAI    As Double   '商品化　箱代


Private GK_SYUKA_CNT        As Double   '出荷件数   2017.12.27
Private GK_NYUKA_CNT        As Double   '入荷件数　 2017.12.27

Private GK_SYUKA_QTY        As Double   '出荷数量   2017.12.27
Private GK_NYUKA_QTY        As Double   '入荷数量   2017.12.27


Dim Name1               As String
Dim Name2               As String
    
Dim ITEM                As String

Dim ADDR1               As String
Dim ADDR2               As String

Dim SYAMEI              As String

Dim BIKOU1              As String
Dim BIKOU2              As String
Dim BIKOU3              As String

Dim SHIMEBI             As String


Private Type MEISAI_TBL_tag
    HIN_NAME    As String               '請求書（表紙） 品名
    TEKIYO      As String               '請求書（表紙） 摘要
    KINGAKU     As Double               '請求書（表紙） 金額
End Type
Private MEISAI_TBL()    As MEISAI_TBL_tag



Private Type YUKO_SOKO_TBL                      '有効ﾎｽﾄ倉庫取り込みテーブル
    HS_SOKO             As String * 8
    NAIGAI              As String * 1
End Type

Dim Soko_T()            As YUKO_SOKO_TBL        '倉庫情報

Dim MyCenter            As String * 1           '洗濯機 小野OR袋井

Dim INV_IO_TANKA_No     As String * 2           '未登録時の入出庫区分
Dim INV_SYUKA_KBN       As String * 2           '未登録時の出荷区分

Dim INV_KBN11           As String * 2           '直送
Dim INV_KBN12           As String * 2           '海外
Dim INV_KBN71           As String * 2           '科目振替

Dim RYOHEN              As String * 2           '良品返品コード

Dim SYUKA_SHEET_TITLE   As String               '出荷明細シートタイトル 2009.06.17
Dim NYUKO_SHEET_TITLE   As String               '入庫明細シートタイトル 2009.06.17

'--------------------------------------- EXCEL用定数    2015.07.06
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
Private Const xlUnderlineStyleSingle% = 2
Private Const xlManual% = -4135
Private Const xlAscending% = 1
Private Const xlNo% = 2
Private Const xlTopToBottom% = 1
Private Const xlPinYin% = 1
'--------------------------------------- EXCEL用定数

'Private Const LAST_UPDATE_DAY$ = "([SEI0020] 2017.12.28 11:15)"
Private Const LAST_UPDATE_DAY$ = "([SEI0020] 2017.12.28 14:00)"



Private Sub Command1_Click(Index As Integer)
Dim ans As Integer

    Select Case Index
        Case 0                          '集計
        
            If Update_Proc() Then
                Unload Me
            End If
        
        Case 1                          'EXCEL出力(表紙)
        
            If COVER_Proc() Then
                Unload Me
            End If
        
        Case 2                          'EXCEL出力(出荷明細)
        
            If SYU_DETAIL_Proc() Then
                Unload Me
            End If
        
        Case 3                          'EXCEL出力(科目振替明細)
        
            If KAMOKU_DETAIL_Proc() Then
                Unload Me
            End If
        
        
        

        Case 4                          'EXCEL出力(入庫明細)
        
        
            If NYU_DETAIL_Proc() Then
                Unload Me
            End If
        
        
        
        
        
        Case 5                          'EXCEL出力(良品返品明細)
        
        
            If RYOHEN_DETAIL_Proc() Then
                Unload Me
            End If
        
        
        Case 6                          '終了
            Unload Me
        Case Else
            Beep
    End Select

End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
Private Sub Form_Load()
Dim i           As Integer
Dim j           As Integer
Dim c           As String * 128
Dim sts         As Integer

Dim S_DATE      As String
Dim E_DATE      As String
Dim S_YY        As String * 4
Dim S_MM        As String * 2
Dim S_DD        As String * 2
    
Dim Max_Soko    As Integer
    
    
    
'    If App.PrevInstance Then                       2017.12.27 DELETE
'        Beep
'        MsgBox "同一プログラム実行中です。"
'        End
'    End If


    
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]請求書作成処理", Me.hwnd, 0)
    'ペイン複数作る
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SETPARTS, 0, -1)


    SEI00201.Caption = SEI00201.Caption & LAST_UPDATE_DAY           '2017.12.27

    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                


    Max_Row = 9999
                                
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '向け先マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '管理マスタＯＰＥＮ
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '出荷予定ＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '過日分出荷予定ＯＰＥＮ
    If DEL_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'Y_GLICSＯＰＥＮ
    If Y_GLICS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'ロケーション別単価設定マスタＯＰＥＮ
    If SE_LOC_TANKA_M_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷先別単価設定マスタＯＰＥＮ
    If SE_SHIP_TANKA_M_Open(BtOpenNomal) Then
        Unload Me
    End If


    '管理マスタ読み込み
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
            
        Case Else
            Unload Me
    End Select

    If JGYOB_TB_Set(1) Then      '事業部の獲得
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

                                
                                '倉庫最大数を取り込み
                                
    If GetIni(App.EXEName, "MAX_SOKO", App.EXEName, c) Then
        Max_Soko = 1
    Else
        If Not IsNumeric(RTrim(c)) Then
            Max_Soko = 1
        Else
            Max_Soko = CInt(RTrim(c))
        End If
    End If
                                
                                
                                '在庫取り込み用テーブル作成
    ReDim Soko_T(0 To UBound(JGYOBU_T), 0 To Max_Soko - 1)
                                '倉庫情報取り込み
    For i = 0 To UBound(JGYOBU_T)
        j = 0
        Do
                                '有効倉庫獲得
            If GetIni(App.EXEName, "SOKO" & JGYOBU_T(i).CODE & Format(j + 1, "0"), App.EXEName, c) Then
'                Beep                                                                                       '2015.07.06
'                MsgBox "倉庫情報の獲得に失敗しました。処理を中止して下さい。"                              '2015.07.06
'                End                                                                                        '2015.07.06
                                                                                                            '2015.07.06
                Call LOG_OUT(LOG_F, "倉庫情報の獲得エラー　SEI0020.INI [SEI0020.INI] " & "SOKO" & JGYOBU_T(i).CODE & Format(j + 1, "0") & "=")
                
                c = "**"
            End If                                                                                          '2015.07.06
    
    
            
            If Trim(c) = "**" Then  '倉庫指定終了
                Exit Do
            End If
    
    
            Soko_T(i, j).HS_SOKO = Trim(c)
                            '国内外情報獲得
            If GetIni(App.EXEName, "NAIG" & JGYOBU_T(i).CODE & Format(j + 1, "0"), App.EXEName, c) Then
                Beep                                                                                    '2015.07.06
                MsgBox "国内外情報の獲得に失敗しました。処理を中止して下さい。"                         '2015.07.06
                End                                                                                     '2015.07.06
            End If                                                                                      '2015.07.06
                
            Soko_T(i, j).NAIGAI = Trim(c)
            j = j + 1
        Loop
    
    Next i

    If GetIni(App.EXEName, "CENTER", App.EXEName, c) Then
        MyCenter = "O"
    Else
        MyCenter = Trim(c)
    End If


    '未登録時の入出庫区分
    If GetIni(App.EXEName, "INV_IO_TANKA_No", App.EXEName, c) Then
        INV_IO_TANKA_No = ""
    Else
        INV_IO_TANKA_No = Trim(c)
    End If

    '未登録時の出荷区分
    If GetIni(App.EXEName, "INV_SYUKA_KBN", App.EXEName, c) Then
        INV_SYUKA_KBN = ""
    Else
        INV_SYUKA_KBN = Trim(c)
    End If

    '直送の出荷区分
    If GetIni(App.EXEName, "KBN11", App.EXEName, c) Then
        INV_KBN11 = ""
    Else
        INV_KBN11 = Trim(c)
    End If
    
    '海外の出荷区分
    If GetIni(App.EXEName, "KBN12", App.EXEName, c) Then
        INV_KBN12 = ""
    Else
        INV_KBN12 = Trim(c)
    End If

    '科目振替の出荷区分
    If GetIni(App.EXEName, "KBN71", App.EXEName, c) Then
        INV_KBN71 = ""
    Else
        INV_KBN71 = Trim(c)
    End If

    '良品返品の入出庫区分
    If GetIni(App.EXEName, "RYOHEN", App.EXEName, c) Then
        RYOHEN = ""
    Else
        RYOHEN = Trim(c)
    End If



    '2009.06.17
    If GetIni(App.EXEName, "SYUKA_SHEET_TITLE", App.EXEName, c) Then
        SYUKA_SHEET_TITLE = "③④⑤⑥出荷明細"
    Else
        SYUKA_SHEET_TITLE = Trim(c)
    End If

    '2009.06.17
    If GetIni(App.EXEName, "NYUKO_SHEET_TITLE", App.EXEName, c) Then
        NYUKO_SHEET_TITLE = "⑦入庫"
    Else
        NYUKO_SHEET_TITLE = Trim(c)
    End If





    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc
    '仕向け先のセット
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD, 0) Then
        Unload Me
    End If
    
    
    
    
    
    
    
    
    
    
    
    '請求書（表紙）の設定項目獲得
    If GetIni(App.EXEName, "Name1", App.EXEName, c) Then
        Name1 = ""
    Else
        Name1 = Trim(c)
    End If
    
    If GetIni(App.EXEName, "Name2", App.EXEName, c) Then
        Name2 = ""
    Else
        Name2 = Trim(c)
    End If
    If GetIni(App.EXEName, "Item", App.EXEName, c) Then
        ITEM = ""
    Else
        ITEM = Trim(c)
    End If
    If GetIni(App.EXEName, "ADDR1", App.EXEName, c) Then
        ADDR1 = ""
    Else
        ADDR1 = Trim(c)
    End If
    If GetIni(App.EXEName, "ADDR2", App.EXEName, c) Then
        ADDR2 = ""
    Else
        ADDR2 = Trim(c)
    End If
    If GetIni(App.EXEName, "SYAMEI", App.EXEName, c) Then
        SYAMEI = ""
    Else
        SYAMEI = Trim(c)
    End If
    If GetIni(App.EXEName, "BIKOU1", App.EXEName, c) Then
        BIKOU1 = ""
    Else
        BIKOU1 = Trim(c)
    End If
    If GetIni(App.EXEName, "BIKOU2", App.EXEName, c) Then
        BIKOU2 = ""
    Else
        BIKOU2 = Trim(c)
    End If
    If GetIni(App.EXEName, "BIKOU3", App.EXEName, c) Then
        BIKOU3 = ""
    Else
        BIKOU3 = Trim(c)
    End If
    If GetIni(App.EXEName, "SHIMEBI", App.EXEName, c) Then
        SHIMEBI = ""
    Else
        SHIMEBI = Trim(c)
    End If
    
        
    i = -1
    j = 1
    Do
        If GetIni(App.EXEName, "HIN_NAME" & Format(j, "00"), App.EXEName, c) Then
            Exit Do
        End If
              
        i = i + 1
        ReDim Preserve MEISAI_TBL(0 To i)
        MEISAI_TBL(i).HIN_NAME = Trim(c)
        
        If GetIni(App.EXEName, "TEKIYO" & Format(j, "00"), App.EXEName, c) Then
            
                        
            
            MEISAI_TBL(i).TEKIYO = ""
        Else
            MEISAI_TBL(i).TEKIYO = Trim(c)
        End If
    
        j = j + 1
    Loop
    
    
    Combo1(pcmbSHIMUKE).ListIndex = 0

    
    E_DATE = Format(Now, "YYYY/MM/DD")
    S_DATE = DateAdd("m", -1, Left(E_DATE, 8) & SHIMEBI)
    S_DD = Right(S_DATE, 2)
    S_DD = Format(CInt(S_DD) + 1, "00")
    
    S_DATE = Left(S_DATE, 7) & "/" & S_DD
    If IsDate(S_DATE) Then
    Else
        S_MM = Mid(S_DATE, 6, 2)
        S_MM = Format(S_MM + 1, "00")

        S_DATE = Right(S_DATE, 5) & S_MM & "/01"


        If IsDate(S_DATE) Then
        Else
            S_YY = Right(S_DATE, 4)
            S_YY = Format(CInt(S_YY) + 1, "0000")

            S_DATE = S_YY & "/01/01"
        End If
    End If


    Text1(ptxS_Date).Text = S_DATE
    Text1(ptxE_Date).Text = E_DATE
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
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "コードマスタ")
        End If
    End If
                                            '向け先マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先マスタ")
        End If
    End If
                                            '管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "管理マスタ")
        End If
    End If
                                            '出荷予定ＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定")
        End If
    End If
                                            '過日分出荷予定ＣＬＯＳＥ
    sts = BTRV(BtOpClose, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "過日分出荷予定")
        End If
    End If
    
                                'ロケーション別単価設定マスタＯＰＥＮ
    sts = BTRV(BtOpClose, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ロケーション別単価設定マスタ")
        End If
    End If
                                '出荷先別単価設定マスタＯＰＥＮ
    sts = BTRV(BtOpClose, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷先別単価設定マスタ")
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

    SEI00201.MousePointer = vbHourglass


    TDBGrid1.Enabled = False


    Call Ctrl_Lock(SEI00201)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(SEI00201)

    TDBGrid1.Enabled = True

    SEI00201.MousePointer = vbDefault

End Sub


Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   請求集計処理
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim ans         As Integer
    
Dim Row         As Long
    
Dim com         As Integer
    
Dim GK_ZEI_KIN  As Long
    
    
Dim Skip_Flg    As Boolean
    
Dim i           As Integer
Dim j           As Integer
    
    
    
    
    Update_Proc = True
    
    Call Input_Lock
                                    
                        '集計値　クリアー
    GK_SYUKO_KOURYO = 0
    GK_SYUKA_KOURYO = 0
    GK_NYUKA_KOURYO = 0
    GK_RYOHEN_KOURYO = 0

    GK_SYOHIN_KOURYO = 0
    GK_SYOHIN_SHIZAI = 0
                                    
                                    
    GK_SYUKA_CNT = 0    '出荷件数   2017.12.27
    GK_NYUKA_CNT = 0    '入荷件数　 2017.12.27
    
    GK_SYUKA_QTY = 0    '出荷数量   2017.12.27
    GK_NYUKA_QTY = 0    '入荷数量   2017.12.27
                                    
                                    
                                    
                                    
                                    
                        'テーブルリセット
    Set SEIKYU = Nothing
    Row = Min_Row - 1
    
    
    
    If Check1(pchkSYUKA).Value = vbChecked Then
    
    
        
                                            
        '------------------------------------------------------------------------   '過日分出荷予定の読み込み
        Call UniCode_Conv(K1_DEL_SYU.KEY_SYUKA_YMD, Format(Text1(ptxS_Date).Text, "YYYYMMDD"))
        
        com = BtOpGetGreaterEqual
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "過日分出荷予定処理開始", Me.hwnd, 0)
        
        Do
            
            DoEvents
            
            sts = BTRV(com, DEL_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K1_DEL_SYU, Len(K1_DEL_SYU), 1)
            Select Case sts
                Case BtNoErr
                
                
                    If Format(Text1(ptxE_Date), "YYYYMMDD") < StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                
                Case Else
                    Call File_Error(sts, com, "出荷予定")
                    Exit Function
            End Select
        
            
            Skip_Flg = False
            
'2008.05.16            If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
'2008.05.16
'2008.05.16                If (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Or _
'2008.05.16                    (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2") Or _
'2008.05.16                    (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Then
'2008.05.16
'2008.05.16                Else
'2008.05.16                    Skip_Flg = True
'2008.05.16
'2008.05.16                End If
'2008.05.16
'2008.05.16            End If
            
            
            '2008.05.16
            If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
            
                If (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Or _
                    (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2") Then
                
                Else
                    Skip_Flg = True
                
                End If
            
            End If
            '2008.05.16
            
            
            
            If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
                If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
                
                    If Not IsNumeric(StrConv(Y_SYUREC.LK_SEQ_NO, vbUnicode)) Then
                        Skip_Flg = True
                    End If
                End If
            
                If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2" Then
                
                    If Trim(StrConv(Y_SYUREC.KAN_YMD, vbUnicode)) = "" Then
                        Skip_Flg = True
                    End If
                End If
            
                If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
                
                    If Trim(StrConv(Y_SYUREC.KAN_YMD, vbUnicode)) = "" Then
                        Skip_Flg = True
                    End If
                End If
            End If
            
    
            If Not Skip_Flg Then
        
                If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" And _
                    (StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "25" Or StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "29") Then
                    If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                        StrConv(Y_SYUREC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
                    Else
                
                
                        Row = Row + 1
                
                        If SYU_Grid_Set_Proc(Row) Then
                            Exit Function
                        End If
                
                    End If
                End If
            
            End If
            
            com = BtOpGetNext
        Loop
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "過日分出荷予定処理終了", Me.hwnd, 0)

    '------------------------------------------------------------------------   '出荷予定の読み込み
        Call UniCode_Conv(K0_Y_SYU.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, "")
        
        com = BtOpGetGreater
        
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "出荷予定処理開始", Me.hwnd, 0)
        
        
        Do
            
            DoEvents
            
            sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            Select Case sts
                Case BtNoErr
                    If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                
                Case Else
                    Call File_Error(sts, com, "出荷予定")
                    Exit Function
            End Select
    
        
            
            Skip_Flg = False
            
            
'2008.05.16            If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
'2008.05.16
'2008.05.16                If (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Or _
'2008.05.16                    (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2") Or _
'2008.05.16                    (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Then
'2008.05.16
'2008.05.16                Else
'2008.05.16                    Skip_Flg = True
'2008.05.16
'2008.05.16                End If
'2008.05.16
'2008.05.16            End If
            
            
            '2008.05.16
            If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
            
                If (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Or _
                    (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2") Then
                
                Else
                    Skip_Flg = True
                
                End If
            
            End If
            '2008.05.16
            
            
            
            If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
                If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
                
                    If Not IsNumeric(StrConv(Y_SYUREC.LK_SEQ_NO, vbUnicode)) Then
                        Skip_Flg = True
                    End If
                End If
            
                If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2" Then
                
                    If Trim(StrConv(Y_SYUREC.KAN_YMD, vbUnicode)) = "" Then
                        Skip_Flg = True
                    End If
                End If
            
                If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
                
                    If Trim(StrConv(Y_SYUREC.KAN_YMD, vbUnicode)) = "" Then
                        Skip_Flg = True
                    End If
                End If
            End If
            
            If Not Skip_Flg Then
            
                If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" And _
                    (StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "25" Or StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "29") Then
                
                
                    If Format(Text1(ptxS_Date).Text, "YYYYMMDD") > StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) Or _
                        Format(Text1(ptxE_Date).Text, "YYYYMMDD") < StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) Then
                    Else
                
                        If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                            StrConv(Y_SYUREC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
                        Else
                            Row = Row + 1
                        
                            If SYU_Grid_Set_Proc(Row) Then
                                Exit Function
                            End If
                        
                        End If
                
                    End If
                End If
            End If
        
        
        
            com = BtOpGetNext
        Loop
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "出荷予定処理終了", Me.hwnd, 0)
    
    End If
    '------------------------------------------------------------------------   'Y_GLICSの読み込み（入庫）
    If Check1(pchkNYUKO).Value = vbChecked Then
    
        Call UniCode_Conv(K0_Y_GLICS.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_Y_GLICS.SYUKA_YMD, Format(Text1(ptxS_Date).Text, "YYYYMMDD"))
        Call UniCode_Conv(K0_Y_GLICS.TEXT_NO, "")
        
        com = BtOpGetGreater
        
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "入庫処理開始", Me.hwnd, 0)
        
        Do
            
            DoEvents
            
            
            sts = BTRV(com, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
            Select Case sts
                Case BtNoErr
                    
                    If StrConv(Y_GLICSREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Then
                        Exit Do
                    End If
                    
                    If StrConv(Y_GLICSREC.SYUKA_YMD, vbUnicode) > Format(Text1(ptxE_Date).Text, "YYYYMMDD") Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                
                Case Else
                    Call File_Error(sts, com, "Y_GLICS")
                    Exit Function
            End Select
    
    
            Skip_Flg = True
            For i = 0 To UBound(JGYOBU_T)               '収支区分のチェック
                If StrConv(Y_GLICSREC.JGYOBU, vbUnicode) = JGYOBU_T(i).CODE Then
                    For j = 0 To UBound(Soko_T, 2)
                        If Trim(StrConv(Y_GLICSREC.H_SOKO, vbUnicode)) = Trim(Soko_T(i, j).HS_SOKO) Then
                            Skip_Flg = False
                            Exit For
                        End If
                    Next j
                    Exit For
                End If
            Next i
        
            
            '2008.11.27 "4"追加
            If StrConv(Y_GLICSREC.IO_KBN, vbUnicode) <> "1" And StrConv(Y_GLICSREC.IO_KBN, vbUnicode) <> "4" Then
                Skip_Flg = True
            End If
        
        
            If StrConv(Y_GLICSREC.PM_KBN, vbUnicode) = "-" Then
                Skip_Flg = True
            End If
        
        
'            If Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) = "36003" Then
'                Skip_Flg = True
'            End If
'
'            If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "PP" Then
'                Skip_Flg = True
'            End If
'
'
'
'
'            Select Case StrConv(Y_GLICSREC.JGYOBU, vbUnicode)
'                Case SOJIKI                         '掃除機
'
'
'                    If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "KM" Then
'                        Skip_Flg = True
'                    End If
'
'                    If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "KK" Then
'                        Skip_Flg = True
'                    End If
'
'                    If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "GG" Then
'                        Skip_Flg = True
'                    End If
'
'                    If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "SS" Then
'                        Skip_Flg = True
'                    End If
'
'                    If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 5) = "0090K" Then
'                        Skip_Flg = True
'                    End If
'
'                    If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 5) = "0092H" Then
'                        Skip_Flg = True
'                    End If
'
'                    If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "AA" Then
'                        Skip_Flg = True
'                    End If
'
'
'
'                Case DENKA, SUIHAN, SENTAKU         '電化、炊飯、洗濯機（アイロン）
'
'
'                    Select Case MyCenter
'
'                        Case "O"
'
'                            If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "01" Then
'                                Skip_Flg = True
'                            End If
'
'                            If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 3) = "H33" Then
'                                Skip_Flg = True
'                            End If
'                            If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 3) = "H22" Then
'                                Skip_Flg = True
'                            End If
'
'                            If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "05" Then
'                                Skip_Flg = True
'                            End If
'
'                            If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "08" Then
'                                Skip_Flg = True
'                            End If
'
'                            If StrConv(Y_GLICSREC.JGYOBU, vbUnicode) = DENKA Then
'
'                                If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) <> "02" And _
'                                    Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) <> "G11" And _
'                                    Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) <> "G22" Then
'                                    Skip_Flg = True
'                                End If
'                            End If
'
'                            If (StrConv(Y_GLICSREC.JGYOBU, vbUnicode) = SUIHAN Or _
'                                StrConv(Y_GLICSREC.JGYOBU, vbUnicode) = SENTAKU) Then
'                                If (Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "P3" Or _
'                                    Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "S3") Then
'                                    Skip_Flg = True
'                                End If
'                            End If
'
'
'
'                            If (StrConv(Y_GLICSREC.JGYOBU, vbUnicode) = SUIHAN Or _
'                                StrConv(Y_GLICSREC.JGYOBU, vbUnicode) = SENTAKU) Then
'                                If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "RO" Then
'                                    Skip_Flg = True
'                                End If
'                            End If
'
'                            If (StrConv(Y_GLICSREC.JGYOBU, vbUnicode) = SUIHAN Or _
'                                StrConv(Y_GLICSREC.JGYOBU, vbUnicode) = SENTAKU) Then
'                                If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "07" Then
'                                    Skip_Flg = True
'                                End If
'                            End If
'
'
'
'
'
'
'                        Case "F"
'
'                            If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "P2" Then
'                                Skip_Flg = True
'                            End If
'
'                            If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "U2" Then
'                                Skip_Flg = True
'                            End If
'
'
'                            If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 3) <> "904" Then
'                                If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 1) = "9" Then
'                                  Skip_Flg = True
'                                End If
'                            End If
'
'                    End Select
'            End Select
        
        
        
            If Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) = "01B11" Or _
                Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) = "01C11" Then
            Else
                Skip_Flg = True
            End If
        
        
        
        
        
        
        
    
            If Not Skip_Flg Then
                If StrConv(Y_GLICSREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                    StrConv(Y_GLICSREC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
                Else
                    Row = Row + 1
            
            
            
                    If NYU_Grid_Set_Proc(Row) Then
                        Exit Function
                    End If
    
    
                End If
        
            End If
        
        
            com = BtOpGetNext
        Loop
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "入庫処理終了", Me.hwnd, 0)
    
    
    End If







    '------------------------------------------------------------------------   'Y_GLICSの読み込み(良品返品)
    If Check1(pchkRYOHEN).Value = vbChecked Then
    
        Call UniCode_Conv(K0_Y_GLICS.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_Y_GLICS.SYUKA_YMD, Format(Text1(ptxS_Date).Text, "YYYYMMDD"))
        Call UniCode_Conv(K0_Y_GLICS.TEXT_NO, "")
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "良品返品処理開始", Me.hwnd, 0)
        
        com = BtOpGetGreater
        
        Do
            
            DoEvents
            
            
            sts = BTRV(com, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
            Select Case sts
                Case BtNoErr
                    
                    If StrConv(Y_GLICSREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Then
                        Exit Do
                    End If
                    
                    If StrConv(Y_GLICSREC.SYUKA_YMD, vbUnicode) > Format(Text1(ptxE_Date).Text, "YYYYMMDD") Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                
                Case Else
                    Call File_Error(sts, com, "Y_GLICS")
                    Exit Function
            End Select
    
    
            Skip_Flg = True
            For i = 0 To UBound(JGYOBU_T)               '収支区分のチェック
                If StrConv(Y_GLICSREC.JGYOBA, vbUnicode) = JGYOBU_T(i).CODE Then
                    For j = 0 To UBound(Soko_T, 2)
                        If Trim(Y_GLICSREC.H_SOKO) = Soko_T(i, j).HS_SOKO Then
                            Skip_Flg = False
                            Exit For
                        End If
                    Next j
                    Exit For
                End If
            Next i
        
            
            If StrConv(Y_GLICSREC.IO_KBN, vbUnicode) <> "1" Then
                Skip_Flg = True
            End If
        
        
            If StrConv(Y_GLICSREC.PM_KBN, vbUnicode) = "-" Then
                Skip_Flg = True
            End If
        
        
'            If Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) = "36003" Then
'                Skip_Flg = True
'            End If
        
'            If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "PP" Then
'                Skip_Flg = True
'            End If
            
            
            
            
            If Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) = "0221B" Or _
                Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) = "0221C" Then
            Else
                Skip_Flg = True
            End If
        
        
        
        
        
        
    
            If Not Skip_Flg Then
    
    
                If StrConv(Y_GLICSREC.NAIGAI, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
    
                    Row = Row + 1
            
                    If RYOHEN_Grid_Set_Proc(Row) Then
                        Exit Function
                    End If
    
                End If
        
            End If
        
        
            com = BtOpGetNext
        Loop
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "良品返品処理終了", Me.hwnd, 0)
    End If



'    SEIKYU.QuickSort 1, SEIKYU.UpperBound(1), ColSYUKA_YMD, 0, XTYPE_STRING
        
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "合計処理中", Me.hwnd, 0)


    Set TDBGrid1.Array = SEIKYU
    
    
'    TDBGrid1.Bookmark = Null
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.ScrollBars = dbgAutomatic








    Text1(ptxGK_SYUKO_KOURYO).Text = Format(ToRoundUp(GK_SYUKO_KOURYO, 0), "#,##0")
    Text1(ptxGK_SYUKA_KOURYO).Text = Format(ToRoundUp(GK_SYUKA_KOURYO, 0), "#,##0")
    Text1(ptxGK_NYUKA_KOURYO).Text = Format(ToRoundUp(GK_NYUKA_KOURYO, 0), "#,##0")
    Text1(ptxGK_RYOHEN_KOURYO).Text = Format(ToRoundUp(GK_RYOHEN_KOURYO, 0), "#,##0")
    Text1(ptxGK_KOURYO).Text = Format(ToRoundUp(GK_SYUKO_KOURYO, 0) + _
                                ToRoundUp(GK_SYUKA_KOURYO, 0) + _
                                ToRoundUp(GK_NYUKA_KOURYO, 0) + _
                                ToRoundUp(GK_RYOHEN_KOURYO, 0), "#,##0")

    Text1(ptxGK_SYOHIN_KOURYO).Text = Format(ToRoundUp(GK_SYOHIN_KOURYO, 0), "#,##0")
    Text1(ptxGK_SYOHIN_SHIZAI).Text = Format(ToRoundUp(GK_SYOHIN_SHIZAI, 0), "#,##0")
    Text1(ptxGK_SYOHIN).Text = Format(ToRoundUp(GK_SYOHIN_KOURYO, 0) + ToRoundUp(GK_SYOHIN_SHIZAI, 0), "#,##0")

    Text1(ptxGK_SEIKYU).Text = Format(ToRoundUp(GK_SYUKO_KOURYO, 0) + _
                                        ToRoundUp(GK_SYUKA_KOURYO, 0) + _
                                        ToRoundUp(GK_NYUKA_KOURYO, 0) + _
                                        ToRoundUp(GK_RYOHEN_KOURYO, 0) + _
                                        ToRoundUp(GK_SYOHIN_KOURYO, 0) + _
                                        ToRoundUp(GK_SYOHIN_SHIZAI, 0), "#,##0")



    GK_ZEI_KIN = Int((CDbl(Text1(ptxGK_SEIKYU).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                            CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10)


    Text1(ptxGK_ZEI_KIN).Text = Format(GK_ZEI_KIN, "#,##0")

    Text1(ptxGK_SEIKYU_KIN).Text = Format(CDbl(Text1(ptxGK_SEIKYU).Text) + GK_ZEI_KIN, "#,##0")
    
    
    Text1(ptxGK_SYUKA_CNT).Text = Format(GK_SYUKA_CNT, "#,##0")     '2017.12.27
    Text1(ptxGK_SYUKA_QTY).Text = Format(GK_SYUKA_QTY, "#,##0")     '2017.12.27
    Text1(ptxGK_NYUKA_CNT).Text = Format(GK_NYUKA_CNT, "#,##0")     '2017.12.27
    Text1(ptxGK_NYUKA_QTY).Text = Format(GK_NYUKA_QTY, "#,##0")     '2017.12.27
    
    
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "集計終了", Me.hwnd, 0)

    Call Input_UnLock




    Update_Proc = False


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


Private Sub SHORI_Click(Index As Integer)
    Select Case Index
        Case 0 To 6
            Command1(Index).Value = True

'        Case 2      '画面印刷
        
'            Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)   2017.12.27

    End Select

End Sub

Private Sub SHORI_MENU_Click(Index As Integer)

'    Select Case Index
'        Case 0 To 5
'            Command1(Index).Value = True

'        Case 2      '画面印刷
        
'            Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)   2017.12.27

'    End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub
Private Function SYU_Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           出荷データ--＞Grid
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim INV_F       As Boolean

Dim READ_NEXT   As Boolean

Dim wkS_KOUSU_BAIKA    As String   '2009.06.10
Dim wkS_SHIZAI_BAIKA   As String   '2009.06.10

    
    SYU_Grid_Set_Proc = True

    

    SEIKYU.ReDim Min_Row, Row, Min_Col, Max_Col
    
    '伝票日付
    SEIKYU(Row, ColSYUKA_YMD) = Mid(StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode), 7, 2)
    
    
    '伝票№
    SEIKYU(Row, ColDEN_NO) = StrConv(Y_SYUREC.DEN_NO, vbUnicode)
    '出荷先
    SEIKYU(Row, ColMUKE_CODE) = StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & " " & StrConv(Y_SYUREC.MUKE_NAME, vbUnicode)
    
    '品番
    SEIKYU(Row, ColHIN_GAI) = StrConv(Y_SYUREC.HIN_NO, vbUnicode)
    
    '数量
    SEIKYU(Row, ColSURYO) = Format(CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)), "#0")
    
    '出庫工料
    
    If Trim(StrConv(Y_SYUREC.HTANABAN, vbUnicode)) = "" Then
    
    
    
        '2008.08.20 ↓
        If StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2" Then
    
            READ_NEXT = False
        
        
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                    '2009.06.10
                    If Not IsDate(Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 7, 2)) Then
                        
                        wkS_KOUSU_BAIKA = StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)
                        wkS_SHIZAI_BAIKA = StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)
                
                    Else
                
                        If StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) < StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode) Then
                
                            wkS_KOUSU_BAIKA = StrConv(ITEMREC.BEF_S_KOUSU_BAIKA, vbUnicode)
                            wkS_SHIZAI_BAIKA = StrConv(ITEMREC.BEF_S_SHIZAI_BAIKA, vbUnicode)
                
                        Else
                            wkS_KOUSU_BAIKA = StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)
                            wkS_SHIZAI_BAIKA = StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)
                        
                        End If
                
                    End If
                                    
                    'If Not IsNumeric(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) Then
                    If Not IsNumeric(wkS_KOUSU_BAIKA) Then
                    '2009.06.10
                            
                        READ_NEXT = True
                
                    
                    Else
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
                                        Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                                        Exit Function
                                
                                End Select
                            
                            Case BtErrKeyNotFound
                            
                            
                                INV_F = True
                                            
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                                Exit Function
                        
                        End Select
                    End If
                
                Case BtErrKeyNotFound
                    
                    READ_NEXT = True

                
                    Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                    Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                    '2009.06.10
                    Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "00000000.00")
                    Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "00000000.00")
                    '2009.06.10
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            
            End Select
        
        
            If READ_NEXT Then
                Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    
                    
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
                                        Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                                        Exit Function
                                
                                End Select
                            
                            Case BtErrKeyNotFound
                            
                            
                                INV_F = True
                                            
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                                Exit Function
                        
                        End Select
                    
                    
                    Case BtErrKeyNotFound
                        
                        INV_F = True
                    
                    
                        Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                        Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                        '2009.06.10
                        Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "00000000.00")
                        Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "00000000.00")
                        '2009.06.10
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                
                End Select
            End If
        
        
        Else
        '2008.08.20 ↑
    
    
    
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
            
            INV_F = False
            
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                
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
                                    Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_OUT_TANKA, "00000000.00")
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                                    Exit Function
                            
                            End Select
                        
                        Case BtErrKeyNotFound
                            Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_OUT_TANKA, "00000000.00")
                        
                            INV_F = True
                        
                                        
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                            Exit Function
                    
                    End Select
                
                
                Case BtErrKeyNotFound
                    
                    Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                    
                    Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_OUT_TANKA, "00000000.00")
                
                    Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                    Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                
                    '2009.06.10
                    Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "00000000.00")
                    Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "00000000.00")
                    '2009.06.10
                
                    INV_F = True
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            
            End Select
        End If
    
    Else
    
    
        '2008.08.20 ↓
        If StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2" Then
    
            READ_NEXT = False
        
        
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                    If Not IsNumeric(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) Then
                            
                        READ_NEXT = True
                
                    
                    Else
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
                                        Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                                        Exit Function
                                
                                End Select
                            
                            Case BtErrKeyNotFound
                            
                            
                                INV_F = True
                                            
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                                Exit Function
                        
                        End Select
                    End If
                
                Case BtErrKeyNotFound
                    
                    READ_NEXT = True

                
                    Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                    Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                    '2009.06.10
                    Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "00000000.00")
                    Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "00000000.00")
                    '2009.06.10
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            
            End Select
        
        
            If READ_NEXT Then
                Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    
                    
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
                                        Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                                        Exit Function
                                
                                End Select
                            
                            Case BtErrKeyNotFound
                            
                            
                                INV_F = True
                                            
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                                Exit Function
                        
                        End Select
                    
                    
                    Case BtErrKeyNotFound
                        
                        INV_F = True
                    
                    
                        Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                        Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                        '2009.06.10
                        Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "00000000.00")
                        Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "00000000.00")
                        '2009.06.10
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                
                End Select
            End If
        
        
        Else
        '2008.08.20 ↑
    
    
    
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                
                
                
                Case BtErrKeyNotFound
                    
                
                
                    Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                    Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                    '2009.06.10
                    Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "00000000.00")
                    Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "00000000.00")
                    '2009.06.10
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            
            End Select
        
        
        
        
            Call UniCode_Conv(K0_SOKO.Soko_No, Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2))
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
                            Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                            Exit Function
                    
                    End Select
                
                Case BtErrKeyNotFound
                
                
                    INV_F = True
                                
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                    Exit Function
            
            End Select
        End If
    
    End If
    
    If INV_F Then
        
        


        
        
        Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, INV_IO_TANKA_No)
        sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
            
                Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_IO_TANKA_No, "")
                Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_OUT_TANKA, "00000000.00")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                Exit Function
        End Select
    End If
    
    
    
    SEIKYU(Row, ColSYUKO_KOURYO) = Format(CDbl(StrConv(SE_LOC_TANKA_M_REC.SE_OUT_TANKA, vbUnicode)), "#,##0.00")
    
    
    '合計値　加算
 '   GK_SYUKO_KOURYO = GK_SYUKO_KOURYO + Int(CDbl(StrConv(SE_LOC_TANKA_M_REC.SE_OUT_TANKA, vbUnicode)) + 0.9)
    GK_SYUKO_KOURYO = GK_SYUKO_KOURYO + CDbl(SEIKYU(Row, ColSYUKO_KOURYO))
    '品名
    SEIKYU(Row, ColHIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
    
    
    '出荷工料
    Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))
    Call UniCode_Conv(K0_MTS.SS_CODE, "")
    
    INV_F = False
    
    
    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    Select Case sts
        Case BtNoErr
        
        
            Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, StrConv(MTSREC.SYUKA_KBN, vbUnicode))
            sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
            Select Case sts
                Case BtNoErr
                
                
                Case BtErrKeyNotFound
                    Call UniCode_Conv(SE_SHIP_TANKA_M_REC.SE_SYUKA_NAME, "")
                    Call UniCode_Conv(SE_SHIP_TANKA_M_REC.SE_TANKA, "00000000.00")
                
                    INV_F = True
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                    Exit Function
            
            End Select
        
        
        
        Case BtErrKeyNotFound
            
            INV_F = True
            
            Call UniCode_Conv(SE_SHIP_TANKA_M_REC.SE_SYUKA_NAME, "")
            Call UniCode_Conv(SE_SHIP_TANKA_M_REC.SE_TANKA, "00000000.00")
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
            Exit Function
    
    End Select
    
    
    If INV_F Then
        
        If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
        
            Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, INV_KBN11)
            sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
            Select Case sts
                Case BtNoErr
                    INV_F = False
                Case BtErrKeyNotFound
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                    Exit Function
            End Select
        
        
        End If
        
    End If
        
    If INV_F Then
        
        If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2" Then
        
            Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, INV_KBN12)
            sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
            Select Case sts
                Case BtNoErr
                    INV_F = False
                Case BtErrKeyNotFound
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                    Exit Function
            End Select
        
        
        End If
        
    End If
        
    If INV_F Then
        
        If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
        
            Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, INV_KBN71)
            sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
            Select Case sts
                Case BtNoErr
                    INV_F = False
                Case BtErrKeyNotFound
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                    Exit Function
            End Select
        
        
        End If
        
    End If
        
        
        
    If INV_F Then
        Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, INV_SYUKA_KBN)
        sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
            
                Call UniCode_Conv(SE_SHIP_TANKA_M_REC.SE_SYUKA_KBN, "")
                Call UniCode_Conv(SE_SHIP_TANKA_M_REC.SE_TANKA, "00000000.00")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                Exit Function
        End Select
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    SEIKYU(Row, ColSYUKA_KOURYO) = Format(CDbl(StrConv(SE_SHIP_TANKA_M_REC.SE_TANKA, vbUnicode)), "#,##0.00")
    '合計値　加算
    GK_SYUKA_KOURYO = GK_SYUKA_KOURYO + CDbl(SEIKYU(Row, ColSYUKA_KOURYO))
    
    '入庫工料
    SEIKYU(Row, ColNYUKA_KOURYO) = ""
    '良品返品
    SEIKYU(Row, ColRYOHEN_KOURYO) = ""
    
    
    
    '2009.06.10
    
    
If Trim(StrConv(Y_SYUREC.HIN_NO, vbUnicode)) = "304SPN-6" Then
    Debug.Print
End If
    
    
    If Not IsDate(Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 1, 4) & "/" & _
                    Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 5, 2) & "/" & _
                    Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 7, 2)) Then
        
        wkS_KOUSU_BAIKA = StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)
        wkS_SHIZAI_BAIKA = StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)

    Else



        If StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) < StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode) Then

            wkS_KOUSU_BAIKA = StrConv(ITEMREC.BEF_S_KOUSU_BAIKA, vbUnicode)
            wkS_SHIZAI_BAIKA = StrConv(ITEMREC.BEF_S_SHIZAI_BAIKA, vbUnicode)

        Else
            wkS_KOUSU_BAIKA = StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)
            wkS_SHIZAI_BAIKA = StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)
        
        End If

    End If
    '2009.06.10
    
    
    
    
    
    
    
    '商品化　工料
    '2009.06.10
    'If IsNumeric(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) Then
    '    SEIKYU(Row, ColSYOHIN_KOURYO) = Format(CDbl(StrConv(Y_SYUREC.SURYO, vbUnicode)) * CDbl(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)), "#,##0.00")
    If IsNumeric(wkS_KOUSU_BAIKA) Then
        SEIKYU(Row, ColSYOHIN_KOURYO) = Format(CDbl(StrConv(Y_SYUREC.SURYO, vbUnicode)) * CDbl(wkS_KOUSU_BAIKA), "#,##0.00")
        GK_SYOHIN_KOURYO = GK_SYOHIN_KOURYO + CDbl(SEIKYU(Row, ColSYOHIN_KOURYO))
    '2009.06.10
    Else
        SEIKYU(Row, ColSYOHIN_KOURYO) = "0.00"
    End If
    '商品化　資材
    '2009.06.10
    'If IsNumeric(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)) Then
    '    SEIKYU(Row, ColSYOHIN_SHIZAI) = Format(CDbl(StrConv(Y_SYUREC.SURYO, vbUnicode)) * CDbl(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)), "#,##0.00")
    If IsNumeric(wkS_SHIZAI_BAIKA) Then
        SEIKYU(Row, ColSYOHIN_SHIZAI) = Format(CDbl(StrConv(Y_SYUREC.SURYO, vbUnicode)) * CDbl(wkS_SHIZAI_BAIKA), "#,##0.00")
        GK_SYOHIN_SHIZAI = GK_SYOHIN_SHIZAI + CDbl(SEIKYU(Row, ColSYOHIN_SHIZAI))
    '2009.06.10
    Else
        SEIKYU(Row, ColSYOHIN_SHIZAI) = "0.00"
    End If
    
    
    GK_SYUKA_CNT = GK_SYUKA_CNT + 1                                         '2017.12.27
    If IsNumeric(StrConv(Y_SYUREC.SURYO, vbUnicode)) Then                       '2017.12.27
        GK_SYUKA_QTY = GK_SYUKA_QTY + CDbl(StrConv(Y_SYUREC.SURYO, vbUnicode))  '2017.12.27
    End If                                                                      '2017.12.27
        
    SYU_Grid_Set_Proc = False
End Function


Private Function NYU_Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           入荷情報--＞Grid
'----------------------------------------------------------------------------

Dim sts     As Integer
Dim INV_F   As Boolean

    
    NYU_Grid_Set_Proc = True

    

    SEIKYU.ReDim Min_Row, Row, Min_Col, Max_Col
    
    '伝票日付
    SEIKYU(Row, ColSYUKA_YMD) = Mid(StrConv(Y_GLICSREC.SYUKA_YMD, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(Y_GLICSREC.SYUKA_YMD, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(Y_GLICSREC.SYUKA_YMD, vbUnicode), 7, 2)
    
    
    '伝票№
    SEIKYU(Row, ColDEN_NO) = StrConv(Y_GLICSREC.DEN_NO, vbUnicode)
    '出荷先
    SEIKYU(Row, ColMUKE_CODE) = StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)
    '品番
    SEIKYU(Row, ColHIN_GAI) = StrConv(Y_GLICSREC.HIN_NO, vbUnicode)
        
    
    
    '数量
    'SEIKYU(Row, ColDEN_NO) = Format(CLng(StrConv(Y_GLICSREC.SURYO, vbUnicode)), "#0")  '2017.12.27
    SEIKYU(Row, ColSURYO) = Format(CLng(StrConv(Y_GLICSREC.SURYO, vbUnicode)), "#0")    '2017.12.27
    
    '出庫工料
    
    SEIKYU(Row, ColSYUKA_KOURYO) = ""
    '出荷工料
    SEIKYU(Row, ColSYUKO_KOURYO) = ""
    
    '入庫工料
    INV_F = False
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_GLICSREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_GLICSREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_GLICSREC.HIN_NO, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        
        
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
                            Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                            Exit Function
                    
                    End Select
                
                Case BtErrKeyNotFound
                    INV_F = True
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                    Exit Function
            
            End Select
        
        
        Case BtErrKeyNotFound
            INV_F = True
        
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    
    End Select
    
    If INV_F Then
        Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, INV_IO_TANKA_No)
        sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_IO_TANKA_No, "")
            
            
            
                Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_IN_TANKA, "00000000.00")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                Exit Function
        End Select
    End If
    
    SEIKYU(Row, ColNYUKA_KOURYO) = Format(CDbl(StrConv(SE_LOC_TANKA_M_REC.SE_IN_TANKA, vbUnicode)), "#0.00")
    '合計値　加算
    GK_NYUKA_KOURYO = GK_NYUKA_KOURYO + CDbl(SEIKYU(Row, ColNYUKA_KOURYO))
    
    
    
    '品名
    SEIKYU(Row, ColHIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
    
    '良品返品
    SEIKYU(Row, ColRYOHEN_KOURYO) = ""
    '商品化　工料
    SEIKYU(Row, ColSYOHIN_KOURYO) = ""
    '商品化　資材
    SEIKYU(Row, ColSYOHIN_SHIZAI) = ""
    
        
    GK_NYUKA_CNT = GK_NYUKA_CNT + 1                                             '2017.12.27
    
    If IsNumeric(StrConv(Y_GLICSREC.SURYO, vbUnicode)) Then                         '2017.12.27
        GK_NYUKA_QTY = GK_NYUKA_QTY + CDbl(StrConv(Y_GLICSREC.SURYO, vbUnicode))    '2017.12.27
    End If                                                                          '2017.12.27
    
    NYU_Grid_Set_Proc = False
End Function


Private Function COVER_Proc() As Integer
'----------------------------------------------------------------------------
'                   請求書（表示）出力
'----------------------------------------------------------------------------
Dim i                   As Integer
Dim j                   As Integer
Dim Row                 As Integer
    
    
Dim sts                 As Integer
Dim com                 As Integer
    
Dim End_Date            As String


Dim GK_KINGAKU          As Double
Dim WK_TANKA            As Double
Dim ZEI_KIN             As Long

    
Dim Skip_F              As Boolean


'Dim excelApplication    As excel.Application   '2015.07.06
'Dim excelWorkBooks      As excel.Workbooks
'Dim excelWorkBook       As excel.Workbook      '2015.07.06
'Dim excelSheet          As excel.Worksheet     '2015.07.06

Dim excelApplication    As Object               '2015.07.06
Dim excelWorkBook       As Object               '2015.07.06
Dim excelSheet          As Object               '2015.07.06


    

    COVER_Proc = True
    
    Call Input_Lock
    



    
    Set excelApplication = CreateObject("Excel.Application")
'2008.05.16    excelApplication.Visible = True


    
    Set excelWorkBook = excelApplication.Workbooks.Add
'    Set excelSheet = excelWorkBook.Worksheets.Add
    Set excelSheet = excelWorkBook.Worksheets(1)
    

    
    excelApplication.StandardFontSize = 11
    
    excelApplication.StandardFont = "ＭＳ　ゴシック"

    
    
'    excelSheet.Application.Select
'    With excelSheet.Application.Selection.Font
'        .NAME = "ＭＳ　ゴシック"
'        .FontStyle = "標準"
'        .Size = 11
'    End With
    
    'ページ設定
    With excelSheet.Application.ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    
    '列の幅
    excelSheet.Application.Columns(1).Select
    excelSheet.Application.Selection.ColumnWidth = 7.25
    excelSheet.Application.Columns(2).Select
    excelSheet.Application.Selection.ColumnWidth = 36.13
    excelSheet.Application.Columns(3).Select
    excelSheet.Application.Selection.ColumnWidth = 5.38
    excelSheet.Application.Columns(4).Select
    excelSheet.Application.Selection.ColumnWidth = 12.13
    excelSheet.Application.Columns(5).Select
    excelSheet.Application.Selection.ColumnWidth = 13.38
    excelSheet.Application.Columns(6).Select
    excelSheet.Application.Selection.ColumnWidth = 15
    
    '行の幅
    excelSheet.Application.Rows(1).Select
    excelSheet.Application.Selection.RowHeight = 24
    excelSheet.Application.Rows("3:4").Select
    excelSheet.Application.Selection.RowHeight = 14.25
    excelSheet.Application.Rows(12).Select
    excelSheet.Application.Selection.RowHeight = 27
    excelSheet.Application.Rows("14:31").Select
    excelSheet.Application.Selection.RowHeight = 27
    
    'セルの結合
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(1, 6)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(1, 6)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(1, 6)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(1, 6)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "太字"
        .Size = 20
        .Underline = xlUnderlineStyleSingle
    End With
    excelSheet.Application.Cells(1, 1).Value = "請 求 書"
    
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 3)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 3)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 3)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 3)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "標準"
        .Size = 14
    End With
    excelSheet.Application.Cells(12, 1).Value = "合 計 金 額"
    
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 4), excelSheet.Application.Cells(12, 6)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 4), excelSheet.Application.Cells(12, 6)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 4), excelSheet.Application.Cells(12, 6)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 4), excelSheet.Application.Cells(12, 6)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "標準"
        .Size = 14
    End With
    excelSheet.Application.Cells(12, 4).Value = ""
    
    excelSheet.Application.Range(excelSheet.Application.Cells(29, 1), excelSheet.Application.Cells(29, 4)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(29, 1), excelSheet.Application.Cells(29, 4)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(29, 1), excelSheet.Application.Cells(29, 4)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(29, 1), excelSheet.Application.Cells(29, 4)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "標準"
        .Size = 11
    End With
    excelSheet.Application.Cells(29, 1).Value = "税 抜 き 金 額"
    
    excelSheet.Application.Range(excelSheet.Application.Cells(30, 1), excelSheet.Application.Cells(30, 4)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(30, 1), excelSheet.Application.Cells(30, 4)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(30, 1), excelSheet.Application.Cells(30, 4)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(30, 1), excelSheet.Application.Cells(30, 4)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "標準"
        .Size = 11
    End With
    excelSheet.Application.Cells(30, 1).Value = "消    費    税"
    
    excelSheet.Application.Range(excelSheet.Application.Cells(31, 1), excelSheet.Application.Cells(31, 4)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(31, 1), excelSheet.Application.Cells(31, 4)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(31, 1), excelSheet.Application.Cells(31, 4)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(31, 1), excelSheet.Application.Cells(31, 4)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "標準"
        .Size = 11
    End With
    excelSheet.Application.Cells(31, 1).Value = "税 込 み 金 額"
    
    
    
    '罫線
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 6)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    excelSheet.Application.Range(excelSheet.Application.Cells(14, 1), excelSheet.Application.Cells(31, 6)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    '固定項目（見出し）
    excelSheet.Application.Range(excelSheet.Application.Cells(14, 1), excelSheet.Application.Cells(14, 6)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(14, 1), excelSheet.Application.Cells(14, 6)).VerticalAlignment = xlCenter
    excelSheet.Application.Cells(14, 1).Value = "月/日"
    excelSheet.Application.Cells(14, 2).Value = "品     名"
    excelSheet.Application.Cells(14, 3).Value = "数 量"
    excelSheet.Application.Cells(14, 4).Value = "単  価"
    excelSheet.Application.Cells(14, 5).Value = "金　額"
    excelSheet.Application.Cells(14, 6).Value = "摘　要"
    '固定項目（INI）
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 6), excelSheet.Application.Cells(2, 6)).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(2, 6).Value = Left(Format(Now, "YYYY年MM月DD日"), 8) & SHIMEBI & "日"
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(3, 1)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "太字"
        .Size = 11
    End With
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 6), excelSheet.Application.Cells(9, 6)).HorizontalAlignment = xlRight
    
    excelSheet.Application.Cells(3, 1).Value = Name1
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(4, 1)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "太字"
        .Size = 11
        .Underline = xlUnderlineStyleSingle
    End With
    excelSheet.Application.Cells(4, 1).Value = Name2
    
    excelSheet.Application.Range(excelSheet.Application.Cells(8, 2), excelSheet.Application.Cells(8, 2)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "太字"
        .Size = 11
        .Underline = xlUnderlineStyleSingle
    End With
    excelSheet.Application.Cells(8, 2).Value = ITEM
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 6), excelSheet.Application.Cells(7, 6)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "ＭＳ　ゴシック"
        .FontStyle = "太字"
        .Size = 9
    End With
    excelSheet.Application.Cells(4, 6).Value = ADDR1
    excelSheet.Application.Cells(5, 6).Value = ADDR2
    excelSheet.Application.Cells(6, 6).Value = SYAMEI
    excelSheet.Application.Cells(7, 6).Value = BIKOU1
    excelSheet.Application.Cells(8, 6).Value = BIKOU2
    excelSheet.Application.Cells(9, 6).Value = BIKOU3
    
    
    
    For i = 0 To UBound(MEISAI_TBL)
        MEISAI_TBL(i).KINGAKU = 0
    Next i
    
    
    
    
    '------------------------------------------------------------------------   '過日分出荷予定の読み込み
    Call UniCode_Conv(K1_DEL_SYU.KEY_SYUKA_YMD, Format(Text1(ptxS_Date).Text, "YYYYMMDD"))
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        sts = BTRV(com, DEL_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K1_DEL_SYU, Len(K1_DEL_SYU), 1)
        Select Case sts
            Case BtNoErr
            
            
                If Format(Text1(ptxE_Date), "YYYYMMDD") < StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "出荷予定")
                Exit Function
        End Select

        Skip_F = False

'2008.05.16        If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
'2008.05.16
'2008.05.16            If (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Or _
'2008.05.16                (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2") Or _
'2008.05.16                (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Then
'2008.05.16
'2008.05.16            Else
'2008.05.16                Skip_F = True
'2008.05.16
'2008.05.16            End If
'2008.05.16
'2008.05.16        End If
    
    
    
        '2008.05.16
        If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
        
            If (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Or _
                (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2") Then
            
            Else
                Skip_F = True
            
            End If
        
        End If
        '2008.05.16
    
    
    
    
    
    
    
        If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
        
            If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
            
                If Not IsNumeric(StrConv(Y_SYUREC.LK_SEQ_NO, vbUnicode)) Then
                    Skip_F = True
                End If
            End If
        
            If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2" Then
            
                If Trim(StrConv(Y_SYUREC.KAN_YMD, vbUnicode)) = "" Then
                    Skip_F = True
                End If
            End If
        
            If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
            
                If Trim(StrConv(Y_SYUREC.KAN_YMD, vbUnicode)) = "" Then
                    Skip_F = True
                End If
            End If
        
        
        End If
    
    
        If Not Skip_F Then
    
            If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" And _
                (StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "25" Or StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "29") Then
                If StrConv(Y_SYUREC.JGYOBU, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) And _
                    StrConv(Y_SYUREC.NAIGAI, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
                
                '--------- 合計集計
Call LOG_OUT(LOG_F, StrConv(Y_SYUREC.ID_NO, vbUnicode))
                
                    If Cover_Total_Proc(1) Then
                        Exit Function
                
                    End If
                
                '--------- 合計集計
                
                End If
            
            End If
        
        End If
        
        
        com = BtOpGetNext
    Loop
    '------------------------------------------------------------------------   '出荷予定の読み込み
        
    Call UniCode_Conv(K0_Y_SYU.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, "")
    
    com = BtOpGetGreater
    
    Do
        
        DoEvents
        
        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "出荷予定")
                Exit Function
        End Select

    
        
        Skip_F = False

'2008.05.16        If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
'2008.05.16
'2008.05.16            If (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Or _
'2008.05.16                (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2") Or _
'2008.05.16                (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Then
'2008.05.16
'2008.05.16            Else
'2008.05.16                Skip_F = True
'2008.05.16
'2008.05.16            End If
'2008.05.16
'2008.05.16        End If
    
        
        '2008.05.16
        If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
        
            If (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Or _
                (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2") Then
            
            Else
                Skip_F = True
            
            End If
        
        End If
        '2008.05.16
    
    
    
        If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
        
            If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
            
                If Not IsNumeric(StrConv(Y_SYUREC.LK_SEQ_NO, vbUnicode)) Then
                    Skip_F = True
                End If
            End If
        
            If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2" Then
            
                If Trim(StrConv(Y_SYUREC.KAN_YMD, vbUnicode)) = "" Then
                    Skip_F = True
                End If
            End If
        
            If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
            
                If Trim(StrConv(Y_SYUREC.KAN_YMD, vbUnicode)) = "" Then
                    Skip_F = True
                End If
            End If
        
        
        End If
    
    
        If Not Skip_F Then
        
        
            If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" And _
                (StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "25" Or StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "29") Then
            
            
                If Format(Text1(ptxS_Date).Text, "YYYYMMDD") > StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) Or _
                    Format(Text1(ptxE_Date).Text, "YYYYMMDD") < StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) Then
                Else
            
                    If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                        StrConv(Y_SYUREC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
                    Else
                    
                '--------- 合計集計
Call LOG_OUT(LOG_F, StrConv(Y_SYUREC.ID_NO, vbUnicode))
                
                    If Cover_Total_Proc(1) Then
                        Exit Function
                
                    End If
                
                
                
                
                
                
                '--------- 合計集計
                    
                    
                    End If
            
                End If
            End If
        
        End If
    
    
    
        com = BtOpGetNext
    Loop
    
    
    
    '------------------------------------------------------------------------   'Y_GLICSの読み込み
        
    Call UniCode_Conv(K0_Y_GLICS.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_Y_GLICS.SYUKA_YMD, Format(Text1(ptxS_Date).Text, "YYYYMMDD"))
    Call UniCode_Conv(K0_Y_GLICS.TEXT_NO, "")
    
    com = BtOpGetGreater
    
    Do
        
        DoEvents
        
        
        sts = BTRV(com, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
        Select Case sts
            Case BtNoErr
                
                If StrConv(Y_GLICSREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Then
                    Exit Do
                End If
                
                If StrConv(Y_GLICSREC.SYUKA_YMD, vbUnicode) > Format(Text1(ptxE_Date).Text, "YYYYMMDD") Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "Y_GLICS")
                Exit Function
        End Select


        Skip_F = True
        For i = 0 To UBound(JGYOBU_T)               '収支区分のチェック
            If StrConv(Y_GLICSREC.JGYOBU, vbUnicode) = JGYOBU_T(i).CODE Then
                For j = 0 To UBound(Soko_T, 2)
                    If Trim(StrConv(Y_GLICSREC.H_SOKO, vbUnicode)) = Trim(Soko_T(i, j).HS_SOKO) Then
                        Skip_F = False
                        Exit For
                    End If
                Next j
                Exit For
            End If
        Next i
    
        
        '2008.11.27 "4"追加
        If StrConv(Y_GLICSREC.IO_KBN, vbUnicode) <> "1" And StrConv(Y_GLICSREC.IO_KBN, vbUnicode) <> "4" Then
            Skip_F = True
        End If
    
    
        If StrConv(Y_GLICSREC.PM_KBN, vbUnicode) = "-" Then
            Skip_F = True
        End If
    
    
        
        
        
        
        If Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) = "01B11" Or _
            Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) = "01C11" Then
        Else
            Skip_F = True
        End If
       
        
        
        
        
        
        If Not Skip_F Then
    
            
            If StrConv(Y_GLICSREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                StrConv(Y_GLICSREC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
            
            Else
                If Cover_Total_Proc(2) Then
                    Exit Function
                End If
            End If
        End If
    
    
    
    
        com = BtOpGetNext
    Loop
    
    
    
    '------------------------------------------------------------------------   'Y_GLICSの読み込み
        
    Call UniCode_Conv(K0_Y_GLICS.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_Y_GLICS.SYUKA_YMD, Format(Text1(ptxS_Date).Text, "YYYYMMDD"))
    Call UniCode_Conv(K0_Y_GLICS.TEXT_NO, "")
    
    com = BtOpGetGreater
    
    Do
        
        DoEvents
        
        sts = BTRV(com, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(Y_GLICSREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Then
                    Exit Do
                End If
            
            
                If Format(Text1(ptxE_Date).Text, "YYYYMMDD") < StrConv(Y_GLICSREC.SYUKA_YMD, vbUnicode) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "Y_GLICS")
                Exit Function
        End Select

        Skip_F = True
        For i = 0 To UBound(JGYOBU_T)               '収支区分のチェック
            If StrConv(Y_GLICSREC.JGYOBA, vbUnicode) = JGYOBU_T(i).CODE Then
                For j = 0 To UBound(Soko_T, 2)
                    If Trim(Y_GLICSREC.H_SOKO) = Trim(Soko_T(i, j).HS_SOKO) Then
                        Skip_F = False
                        Exit For
                    End If
                Next j
                Exit For
            End If
        Next i
    
        
        If StrConv(Y_GLICSREC.IO_KBN, vbUnicode) <> "1" Then
            Skip_F = True
        End If
    
    
        If StrConv(Y_GLICSREC.PM_KBN, vbUnicode) = "-" Then
            Skip_F = True
        End If
    
    
        
        
        
        If Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) = "02B11" Or _
            Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) = "02C11" Then
        Else
            Skip_F = True
        End If
       
        
        
        
        
        
        If Not Skip_F Then
    
        Else
            
            If StrConv(Y_GLICSREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                StrConv(Y_GLICSREC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
            
            
                If Cover_Total_Proc(3) Then
                    Exit Function
                End If
            End If
        End If
    
    
    
    
        com = BtOpGetNext
    Loop
    
    
    
    
    
    
    For i = 0 To UBound(MEISAI_TBL)
        '月／日
        excelSheet.Application.Range(excelSheet.Application.Cells(15 + i, 1), excelSheet.Application.Cells(15 + i, 1)).HorizontalAlignment = xlCenter
        excelSheet.Application.Range(excelSheet.Application.Cells(15 + i, 1), excelSheet.Application.Cells(15 + i, 1)).Select
        excelSheet.Application.Selection.NumberFormatLocal = "@"
        excelSheet.Application.Selection.HorizontalAlignment = xlCenter
        excelSheet.Application.Cells(15 + i, 1).Value = Format(CInt(Mid(Format(Now, "YYYYMMDD"), 5, 2)), "#") & "/" & SHIMEBI
        '品名
        excelSheet.Application.Cells(15 + i, 2).Value = Trim(MEISAI_TBL(i).HIN_NAME) & "(" & Text1(ptxS_Date).Text & "～" & Text1(ptxE_Date).Text & ")"
        '数量
        excelSheet.Application.Range(excelSheet.Application.Cells(15 + i, 3), excelSheet.Application.Cells(15 + i, 3)).Select
        excelSheet.Application.Selection.NumberFormatLocal = "#,##0"
        excelSheet.Application.Cells(15 + i, 3).Value = 1
        '単価～金額
        excelSheet.Application.Range(excelSheet.Application.Cells(15 + i, 4), excelSheet.Application.Cells(15 + i, 5)).Select
        excelSheet.Application.Selection.NumberFormatLocal = "#,##0"
        excelSheet.Application.Cells(15 + i, 4).Value = ToRoundUp(MEISAI_TBL(i).KINGAKU, 0)
        excelSheet.Application.Cells(15 + i, 5).Value = ToRoundUp(MEISAI_TBL(i).KINGAKU, 0)
        '摘要
        excelSheet.Application.Cells(15 + i, 6).Value = Trim(MEISAI_TBL(i).TEKIYO)
    
    Next i
    
    
    GK_KINGAKU = 0
    For i = 0 To UBound(MEISAI_TBL)
        GK_KINGAKU = GK_KINGAKU + ToRoundUp(MEISAI_TBL(i).KINGAKU, 0)
    Next i
    
    
    
    '税抜き金額
    excelSheet.Application.Range(excelSheet.Application.Cells(29, 5), excelSheet.Application.Cells(31, 5)).Select
    excelSheet.Application.Selection.NumberFormatLocal = "#,##0;""▲ ""#,##0"
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[-14]C:R[-1]C)"
    '消費税
    ZEI_KIN = Int((excelSheet.Application.Cells(29, 5) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10)
    excelSheet.Application.Cells(30, 5).Value = ZEI_KIN
    '税込み金額
    excelSheet.Application.Cells(31, 5).Value = GK_KINGAKU + ZEI_KIN
    '合計金額
    excelSheet.Application.Cells(12, 4).Value = Format(GK_KINGAKU + ZEI_KIN, "\\#,##0")


    excelApplication.Visible = True



'    excelApplication.Quit

    Set excelSheet = Nothing
    Set excelWorkBook = Nothing
'    Set excelWorkBooks = Nothing
    Set excelApplication = Nothing


    
    Call Input_UnLock
    COVER_Proc = False
    

End Function


Private Function SYU_DETAIL_Proc() As Integer
'----------------------------------------------------------------------------
'                   ＥＸＣＥＬ（出荷明細）出力
'----------------------------------------------------------------------------


Dim Row                 As Long
    
    
Dim sts                 As Integer
Dim com                 As Integer
    
Dim i                   As Long
    
Dim End_Date            As String

Dim s_test_now          As String

Dim Skip_F              As Boolean


'Dim excelApplication    As excel.Application   '2015.07.06
'Dim excelWorkBooks      As excel.Workbooks
'Dim excelWorkBook       As excel.Workbook      '2015.07.06
'Dim excelSheet          As excel.Worksheet     '2015.07.06
    
Dim excelApplication    As Object               '2015.07.06
Dim excelWorkBook       As Object               '2015.07.06
Dim excelSheet          As Object               '2015.07.06
    
    
    
On Error GoTo Error_Proc
    
    
s_test_now = Format(Now, "YYYY/MM/DD HH:MM:SS")
    
    SYU_DETAIL_Proc = True
    
    Call Input_Lock
    
    Set excelApplication = CreateObject("Excel.Application")
'excelApplication.Visible = True

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[請求システム]出荷明細出力開始" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
        
    
    
    
    Set excelWorkBook = excelApplication.Workbooks.Add
    
    
    Set excelSheet = excelWorkBook.Worksheets(1)
    
    
    
    excelApplication.StandardFontSize = 11
    
    excelApplication.StandardFont = "ＭＳ　ゴシック"
    
    

    excelSheet.Application.Calculation = xlManual
    excelSheet.Application.MaxChange = 0.001



'    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(1, 4)).Select
'    With excelSheet.Application.Selection.Font
'        .Size = 16
'    End With
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(1, 4)).Font.Size = 16
    
    
    
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[請求システム]出荷明細 見出し設定" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
    
    
    
    excelSheet.Application.Cells(1, 1).Value = "出荷明細表" & _
                                    Trim(StrConv(P_KANRIREC.CENTER_NAME, vbUnicode)) & _
                                    "（" & Text1(ptxS_Date).Text & "～" & _
                                    Text1(ptxE_Date).Text & "）"
    
    
    '列の幅
    excelSheet.Application.Columns(1).ColumnWidth = 7.75
    
    
    'セルの結合
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 12), excelSheet.Application.Cells(2, 13)).HorizontalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 12), excelSheet.Application.Cells(2, 13)).VerticalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 12), excelSheet.Application.Cells(2, 13)).MergeCells = True
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 15), excelSheet.Application.Cells(2, 16)).HorizontalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 15), excelSheet.Application.Cells(2, 16)).VerticalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 15), excelSheet.Application.Cells(2, 16)).MergeCells = True
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 18), excelSheet.Application.Cells(2, 19)).HorizontalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 18), excelSheet.Application.Cells(2, 19)).VerticalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 18), excelSheet.Application.Cells(2, 19)).MergeCells = True
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 20), excelSheet.Application.Cells(2, 21)).HorizontalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 20), excelSheet.Application.Cells(2, 21)).VerticalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 20), excelSheet.Application.Cells(2, 21)).MergeCells = True
    
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 13), excelSheet.Application.Cells(2, 13)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 13), excelSheet.Application.Cells(2, 13)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 13), excelSheet.Application.Cells(2, 13)).MergeCells = True
    
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 14), excelSheet.Application.Cells(2, 15)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 14), excelSheet.Application.Cells(2, 15)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 14), excelSheet.Application.Cells(2, 15)).MergeCells = True
   
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 16), excelSheet.Application.Cells(2, 17)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 16), excelSheet.Application.Cells(2, 17)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 16), excelSheet.Application.Cells(2, 17)).MergeCells = True
    
    
    
    
    
    '２行目見出し設定
'    excelSheet.Application.Cells(2, 12).Value = "出庫工料"
'    excelSheet.Application.Cells(2, 15).Value = "出荷工料"
'    excelSheet.Application.Cells(2, 18).Value = "個装工料"
'    excelSheet.Application.Cells(2, 20).Value = "個装箱代"
    
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 10), excelSheet.Application.Cells(2, 12)).HorizontalAlignment = xlCenter
    excelSheet.Application.Cells(2, 10).Value = "出庫工料"
    excelSheet.Application.Cells(2, 11).Value = "出荷工料"
    excelSheet.Application.Cells(2, 12).Value = "個装"
    excelSheet.Application.Cells(2, 13).Value = "切替"
    excelSheet.Application.Cells(2, 14).Value = "個装工料"
    excelSheet.Application.Cells(2, 16).Value = "個装箱代"
    
    
    '３行目見出し設定
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(3, 1)).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(3, 1).Value = "№"
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 2), excelSheet.Application.Cells(3, 21)).HorizontalAlignment = xlCenter
    
'    excelSheet.Application.Cells(3, 2).Value = "ID-№"
'    excelSheet.Application.Cells(3, 3).Value = "出荷日"
'    excelSheet.Application.Cells(3, 4).Value = "伝№"
'    excelSheet.Application.Cells(3, 5).Value = "出荷先"
'    excelSheet.Application.Cells(3, 6).Value = "出荷先名"
'    excelSheet.Application.Cells(3, 7).Value = "品番"
'    excelSheet.Application.Cells(3, 8).Value = "品名"
'    excelSheet.Application.Cells(3, 9).Value = "数量"
'    excelSheet.Application.Cells(3, 10).Value = "棚番"
'    excelSheet.Application.Cells(3, 11).Value = "出庫区分"
'    excelSheet.Application.Cells(3, 12).Value = "単価"
'    excelSheet.Application.Cells(3, 13).Value = "金額"
'    excelSheet.Application.Cells(3, 14).Value = "出荷区分"
'    excelSheet.Application.Cells(3, 15).Value = "単価"
'    excelSheet.Application.Cells(3, 16).Value = "金額"
'    excelSheet.Application.Cells(3, 17).Value = "個装形態"
'    excelSheet.Application.Cells(3, 18).Value = "単価"
'    excelSheet.Application.Cells(3, 19).Value = "金額"
'    excelSheet.Application.Cells(3, 20).Value = "単価"
'    excelSheet.Application.Cells(3, 21).Value = "金額"
    
    
    excelSheet.Application.Cells(3, 2).Value = "ID-№"
    excelSheet.Application.Cells(3, 3).Value = "出荷日"
    excelSheet.Application.Cells(3, 4).Value = "伝№"
    excelSheet.Application.Cells(3, 5).Value = "出荷先"
    excelSheet.Application.Cells(3, 6).Value = "出荷先名"
    excelSheet.Application.Cells(3, 7).Value = "品番"
    excelSheet.Application.Cells(3, 8).Value = "品名"
    excelSheet.Application.Cells(3, 9).Value = "数量"
    excelSheet.Application.Cells(3, 10).Value = "金額"
    excelSheet.Application.Cells(3, 11).Value = "金額"
    excelSheet.Application.Cells(3, 12).Value = "形態"
    excelSheet.Application.Cells(3, 13).Value = "区分"
    excelSheet.Application.Cells(3, 14).Value = "単価"
    excelSheet.Application.Cells(3, 15).Value = "金額"
    excelSheet.Application.Cells(3, 16).Value = "単価"
    excelSheet.Application.Cells(3, 17).Value = "金額"

    
    
    
    
    
    
    
    
    
    '見出し　罫線
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 21)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 21)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 21)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 21)).Borders(xlEdgeLeft).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 21)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 21)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 21)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 21)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 21)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 21)).Borders(xlEdgeBottom).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 21)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 21)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 21)).Borders(xlEdgeRight).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 21)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 21)).Borders(xlInsideVertical).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 21)).Borders(xlInsideVertical).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 21)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(3, 12), excelSheet.Application.Cells(3, 13)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(3, 12), excelSheet.Application.Cells(3, 13)).Borders(xlDiagonalUp).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(3, 12), excelSheet.Application.Cells(3, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(3, 12), excelSheet.Application.Cells(3, 13)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(3, 12), excelSheet.Application.Cells(3, 13)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(3, 15), excelSheet.Application.Cells(3, 16)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(3, 15), excelSheet.Application.Cells(3, 16)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(3, 15), excelSheet.Application.Cells(3, 16)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(3, 15), excelSheet.Application.Cells(3, 16)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(3, 15), excelSheet.Application.Cells(3, 16)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(3, 18), excelSheet.Application.Cells(3, 21)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(3, 18), excelSheet.Application.Cells(3, 21)).Borders(xlDiagonalUp).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(3, 18), excelSheet.Application.Cells(3, 21)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(3, 18), excelSheet.Application.Cells(3, 21)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(3, 18), excelSheet.Application.Cells(3, 21)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 17)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 17)).Borders(xlDiagonalUp).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 17)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 17)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 17)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic


    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 17)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 17)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 17)).Borders(xlEdgeTop).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 17)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 17)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 17)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 17)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 17)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 17)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 17)).Borders(xlInsideVertical).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 17)).Borders(xlInsideVertical).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 17)).Borders(xlInsideVertical).ColorIndex = xlAutomatic





    excelSheet.Application.Range(excelSheet.Application.Cells(3, 14), excelSheet.Application.Cells(3, 17)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 14), excelSheet.Application.Cells(3, 17)).Borders(xlDiagonalUp).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 14), excelSheet.Application.Cells(3, 17)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 14), excelSheet.Application.Cells(3, 17)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 14), excelSheet.Application.Cells(3, 17)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
    
    
    
    
   
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[請求システム]出荷明細 ﾍﾟｰｼﾞ設定" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
    
    '右下にページ追加 2009.02.20
    excelSheet.Application.ActiveSheet.PageSetup.RightFooter = "&P/&N"
    
    
    
    'ｼｰﾄ名変更　2009.02.20
'    excelSheet.Application.ActiveSheet.NAME = "③④⑤⑥出荷明細"       2009.06.17
    excelSheet.Application.ActiveSheet.NAME = SYUKA_SHEET_TITLE         '2009.06.17
    
    
    
    
    'ページヘッダー固定 2009.02.20
    excelSheet.Application.ActiveSheet.PageSetup.PrintTitleRows = "$2:$3"
    
    '余白
    excelSheet.Application.ActiveSheet.PageSetup.LeftMargin = excelSheet.Application.InchesToPoints(0)
    excelSheet.Application.ActiveSheet.PageSetup.RightMargin = excelSheet.Application.InchesToPoints(0)
    excelSheet.Application.ActiveSheet.PageSetup.TopMargin = excelSheet.Application.InchesToPoints(0)
    excelSheet.Application.ActiveSheet.PageSetup.BottomMargin = excelSheet.Application.InchesToPoints(0.393700787401575)
    
    excelSheet.Application.ActiveSheet.PageSetup.HeaderMargin = excelSheet.Application.InchesToPoints(0)
    excelSheet.Application.ActiveSheet.PageSetup.FooterMargin = excelSheet.Application.InchesToPoints(0)
    
    
    
    
    '印刷　横
'    excelSheet.Application.ActiveSheet.PageSetup.Orientation = xlLandscape
    '横--＞縦 2009.07.30
    excelSheet.Application.ActiveSheet.PageSetup.Orientation = xlPortrait
    
    '印刷　拡大率
    excelSheet.Application.ActiveSheet.PageSetup.Zoom = False
    excelSheet.Application.ActiveSheet.PageSetup.FitToPagesWide = 1
    excelSheet.Application.ActiveSheet.PageSetup.FitToPagesTall = False
    '枠線なし   2009.06.19
    excelSheet.Application.ActiveWindow.DisplayGridlines = False
    
    
    '印刷　中央 2009.07.30
    excelSheet.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[請求システム]出荷明細 過日分出力" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
    
    Row = 3
        
    
    
    '------------------------------------------------------------------------   '過日分出荷予定の読み込み
    Call UniCode_Conv(K1_DEL_SYU.KEY_SYUKA_YMD, Format(Text1(ptxS_Date).Text, "YYYYMMDD"))
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        sts = BTRV(com, DEL_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K1_DEL_SYU, Len(K1_DEL_SYU), 1)
        Select Case sts
            Case BtNoErr
            
            
                If Format(Text1(ptxE_Date), "YYYYMMDD") < StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "出荷予定")
                Exit Function
        End Select

        Skip_F = False

'2008.05.16        If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
'2008.05.16
'2008.05.16            If (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Or _
'2008.05.16                (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2") Or _
'2008.05.16                (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Then
'2008.05.16
'2008.05.16            Else
'2008.05.16                Skip_F = True
'2008.05.16
'2008.05.16            End If
'2008.05.16
'2008.05.16        End If
    
    
        '2008.05.16
        If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then

            If (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Or _
                (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2") Then

            Else
                Skip_F = True

            End If

        End If
        '2008.05.16
    
    
    
        If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
        
            If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
            
                If Not IsNumeric(StrConv(Y_SYUREC.LK_SEQ_NO, vbUnicode)) Then
                    Skip_F = True
                End If
            End If
        
            If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2" Then
            
                If Trim(StrConv(Y_SYUREC.KAN_YMD, vbUnicode)) = "" Then
                    Skip_F = True
                End If
            End If
        
            If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
            
                If Trim(StrConv(Y_SYUREC.KAN_YMD, vbUnicode)) = "" Then
                    Skip_F = True
                End If
            End If
        
        
        End If
    
    
        If Not Skip_F Then
            If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" And _
                (StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "25" Or StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "29") Then
                If StrConv(Y_SYUREC.JGYOBU, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) And _
                    StrConv(Y_SYUREC.NAIGAI, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
            
            
            
                    Row = Row + 1
                
                
If Right(Format(Row - 3, 0), 2) = "00" Or Right(Format(Row - 3, 0), 2) = "50" Then
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]出荷明細 過日分出力" & "出力件数　= " & Row - 3, Me.hwnd, 0)
    DoEvents
End If
                
                    If SYU_Excel_Set_Proc(Row, excelApplication, excelWorkBook, excelSheet) Then
                        Exit Function
                    End If
                
                End If
            End If
        End If
        
        
        com = BtOpGetNext
    Loop
    '------------------------------------------------------------------------   '出荷予定の読み込み
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[請求システム]出荷明細 当日分出力" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
        
    Call UniCode_Conv(K0_Y_SYU.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, "")
    
    com = BtOpGetGreater
    
    Do
        
        DoEvents
        
        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "出荷予定")
                Exit Function
        End Select

        Skip_F = False
        
        
'2008.05.16        If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
'2008.05.16
'2008.05.16            If (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Or _
'2008.05.16                (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2") Or _
'2008.05.16                (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Then
'2008.05.16
'2008.05.16            Else
'2008.05.16                Skip_F = True
'2008.05.16
'2008.05.16            End If
'2008.05.16
'2008.05.16        End If
        
        
            
        '2008.05.16
        If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
        
            If (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Or _
                (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2") Then
            Else
                Skip_F = True
            
            End If
        
        End If
        '2008.05.16
        
        
        
        If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
        
            If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
            
                If Not IsNumeric(StrConv(Y_SYUREC.LK_SEQ_NO, vbUnicode)) Then
                    Skip_F = True
                End If
            End If
        
            If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2" Then
            
                If Trim(StrConv(Y_SYUREC.KAN_YMD, vbUnicode)) = "" Then
                    Skip_F = True
                End If
            End If
        
            If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
            
                If Trim(StrConv(Y_SYUREC.KAN_YMD, vbUnicode)) = "" Then
                    Skip_F = True
                End If
            End If
        
        
        End If
    
        
        If Not Skip_F Then
            If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" And _
                (StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "25" Or StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "29") Then
            
            
                If Format(Text1(ptxS_Date).Text, "YYYYMMDD") > StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) Or _
                    Format(Text1(ptxE_Date).Text, "YYYYMMDD") < StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) Then
                Else
            
                    If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                        StrConv(Y_SYUREC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
                    Else
                        Row = Row + 1
                        

If Right(Format(Row - 3, 0), 2) = "00" Or Right(Format(Row - 3, 0), 2) = "50" Then
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]出荷明細 当日分出力" & "出力件数　= " & Row - 3, Me.hwnd, 0)
    DoEvents
End If
                        
                        If SYU_Excel_Set_Proc(Row, excelApplication, excelWorkBook, excelSheet) Then
                            Exit Function
                        End If
                    End If
            
                End If
            End If
        End If
    
    
    
        com = BtOpGetNext
    Loop
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[請求システム]出荷明細 ﾃﾞｰﾀ出力終了" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
    
    excelSheet.Application.Columns("B:U").EntireColumn.AutoFit
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[請求システム]出荷明細 EXCEL SORT開始" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
    
    '2009.06.10
'    excelSheet.Application.Range(excelSheet.Application.Cells(4, 2), excelSheet.Application.Cells(Row, 21)).Sort _
'                key1:=excelSheet.Application.Cells(4, 7), Order1:=xlAscending, _
'                key2:=excelSheet.Application.Cells(4, 3), Order2:=xlAscending, _
'                key3:=excelSheet.Application.Cells(4, 2), Order1:=xlAscending, _
'                Header:=xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
'                sortmethod:=xlPinYin
'                sortmethod:=xlPinYin, dataoption1:=xlSortNormal, dataoption2:=xlSortTextAsNumbers, dataoption3:=xlSortTextAsNumbers
    '2009.06.10
                
    '2009.07.30 SORTKEY 変更
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 2), excelSheet.Application.Cells(Row, 21)).Sort _
                key1:=excelSheet.Application.Cells(4, 3), Order1:=xlAscending, _
                key2:=excelSheet.Application.Cells(4, 2), Order2:=xlAscending, _
                Header:=xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
                SortMethod:=xlPinYin
'                sortmethod:=xlPinYin, dataoption1:=xlSortNormal, dataoption2:=xlSortTextAsNumbers, dataoption3:=xlSortTextAsNumbers
                
                
    
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[請求システム]出荷明細 EXCEL 集計開始" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
    
    
    Row = Row + 1
    
    '件数合計
    excelSheet.Application.Cells(Row, 1).Font.Size = 11
    excelSheet.Application.Cells(Row, 1).Value = "件数合計"
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 1), excelSheet.Application.Cells(Row + 1, 1)).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(Row + 1, 1).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Cells(Row + 1, 1).Value = Row - 4
    
    
    
    '数量合計
    excelSheet.Application.Cells(Row, 9).Value = "個数合計"
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 9), excelSheet.Application.Cells(Row + 1, 9)).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(Row + 1, 9).NumberFormatLocal = "#,##0_ "
    
    
    
    
    '2009.06.19
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 9), excelSheet.Application.Cells(Row + 1, 9)).FormulaR1C1 = "=SUM(R[" & ((Row) * -1) + 3 & "]C:R[-2]C)"
   
    
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 9), excelSheet.Application.Cells(Row + 1, 9)).Select
''    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((Row) * -1) + 3 & "]C:R[-1]C)"
'    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((Row) * -1) + 3 & "]C:R[-2]C)"
    
    '出庫工料　金額合計
    
'2009.06.17
'    excelSheet.Application.Cells(Row, 12).Value = "③出庫費用合計"
'    If Len(SYUKA_SHEET_TITLE) < 4 Then
'        excelSheet.Application.Cells(Row, 12).Value = "出庫費用合計"
'    Else
'        excelSheet.Application.Cells(Row, 12).Value = Mid(SYUKA_SHEET_TITLE, 1, 1) & "出庫費用合計"
'    End If
    If Len(SYUKA_SHEET_TITLE) < 4 Then
        excelSheet.Application.Cells(Row, 10).Value = "出庫費用合計"
    Else
        excelSheet.Application.Cells(Row, 10).Value = Mid(SYUKA_SHEET_TITLE, 1, 1) & "出庫費用合計"
    End If
'2009.06.17
    
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 13), excelSheet.Application.Cells(Row + 1, 13)).HorizontalAlignment = xlRight
'    excelSheet.Application.Cells(Row + 1, 13).NumberFormatLocal = "#,##0_ "
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 13), excelSheet.Application.Cells(Row + 1, 13)).FormulaR1C1 = "=ROUNDUP(ROUNDDOWN(SUM(R[" & ((Row) * -1) + 3 & "]C:R[-2]C),2),0)"


    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 10), excelSheet.Application.Cells(Row + 1, 10)).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(Row + 1, 10).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 10), excelSheet.Application.Cells(Row + 1, 10)).FormulaR1C1 = "=ROUNDUP(ROUNDDOWN(SUM(R[" & ((Row) * -1) + 3 & "]C:R[-2]C),2),0)"



'2009.06.17
'    excelSheet.Application.ActiveCell.FormulaR1C1 = "=ROUNDUP(SUM(R[" & ((Row) * -1) + 3 & "]C:R[-1]C),0)"
'    excelSheet.Application.ActiveCell.FormulaR1C1 = "=ROUNDUP(ROUNDDOWN(SUM(R[" & ((Row) * -1) + 3 & "]C:R[-2]C),2),0)"
'2009.06.17
    
    
    
    
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 12), excelSheet.Application.Cells(Row + 1, 12)).HorizontalAlignment = xlRight
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row, 13)).MergeCells = True
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 12), excelSheet.Application.Cells(Row + 1, 13)).MergeCells = True
    
    
    '出荷工料　金額合計
'2009.06.17
'    excelSheet.Application.Cells(Row, 15).Value = "④出荷費用合計"
'    If Len(SYUKA_SHEET_TITLE) < 4 Then
'        excelSheet.Application.Cells(Row, 15).Value = "出荷費用合計"
'    Else
'        excelSheet.Application.Cells(Row, 15).Value = Mid(SYUKA_SHEET_TITLE, 2, 1) & "出荷費用合計"
'    End If
    If Len(SYUKA_SHEET_TITLE) < 4 Then
        excelSheet.Application.Cells(Row, 11).Value = "出荷費用合計"
    Else
        excelSheet.Application.Cells(Row, 11).Value = Mid(SYUKA_SHEET_TITLE, 2, 1) & "出荷費用合計"
    End If
'2009.06.17
    
'    excelSheet.Application.Cells(Row + 1, 16).NumberFormatLocal = "#,##0_ "
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 16), excelSheet.Application.Cells(Row + 1, 16)).FormulaR1C1 = "=ROUNDUP(ROUNDDOWN(SUM(R[" & ((Row) * -1) + 3 & "]C:R[-2]C),2),0)"
    excelSheet.Application.Cells(Row + 1, 11).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 11), excelSheet.Application.Cells(Row + 1, 11)).FormulaR1C1 = "=ROUNDUP(ROUNDDOWN(SUM(R[" & ((Row) * -1) + 3 & "]C:R[-2]C),2),0)"

'2009.06.17
'    excelSheet.Application.ActiveCell.FormulaR1C1 = "=ROUNDUP(SUM(R[" & ((Row) * -1) + 3 & "]C:R[-1]C),0)"
'    excelSheet.Application.ActiveCell.FormulaR1C1 = "=ROUNDUP(ROUNDDOWN(SUM(R[" & ((Row) * -1) + 3 & "]C:R[-2]C),2),0)"
'2009.06.17
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 16), excelSheet.Application.Cells(Row + 1, 16)).HorizontalAlignment = xlRight
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row, 16)).MergeCells = True
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 15), excelSheet.Application.Cells(Row + 1, 16)).MergeCells = True
    
    
    '個装工料　金額合計
'2009.06.17
'    excelSheet.Application.Cells(Row, 18).Value = "⑤商品化工料合計"
'    If Len(SYUKA_SHEET_TITLE) < 4 Then
'        excelSheet.Application.Cells(Row, 18).Value = "商品化工料合計"
'    Else
'        excelSheet.Application.Cells(Row, 18).Value = Mid(SYUKA_SHEET_TITLE, 3, 1) & "商品化工料合計"
'    End If
    If Len(SYUKA_SHEET_TITLE) < 4 Then
        excelSheet.Application.Cells(Row, 14).Value = "商品化工料合計"
    Else
        excelSheet.Application.Cells(Row, 14).Value = Mid(SYUKA_SHEET_TITLE, 3, 1) & "商品化工料合計"
    End If
'2009.06.17
    
    
    
    
    
'    excelSheet.Application.Cells(Row + 1, 19).NumberFormatLocal = "#,##0_ "
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 19), excelSheet.Application.Cells(Row + 1, 19)).FormulaR1C1 = "=ROUNDUP(ROUNDDOWN(SUM(R[" & ((Row) * -1) + 3 & "]C:R[-2]C),2),0)"


    excelSheet.Application.Cells(Row + 1, 15).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 15), excelSheet.Application.Cells(Row + 1, 15)).FormulaR1C1 = "=ROUNDUP(ROUNDDOWN(SUM(R[" & ((Row) * -1) + 3 & "]C:R[-2]C),2),0)"



'2009.06.17
'    excelSheet.Application.ActiveCell.FormulaR1C1 = "=ROUNDUP(SUM(R[" & ((Row) * -1) + 3 & "]C:R[-1]C),0)"
'    excelSheet.Application.ActiveCell.FormulaR1C1 = "=ROUNDUP(ROUNDDOWN(SUM(R[" & ((Row) * -1) + 3 & "]C:R[-2]C),2),0)"
'2009.06.17
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 19), excelSheet.Application.Cells(Row + 1, 19)).HorizontalAlignment = xlRight
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 18), excelSheet.Application.Cells(Row, 19)).MergeCells = True
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 18), excelSheet.Application.Cells(Row + 1, 19)).MergeCells = True
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 15), excelSheet.Application.Cells(Row + 1, 15)).HorizontalAlignment = xlRight
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row, 15)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 14), excelSheet.Application.Cells(Row + 1, 15)).MergeCells = True
    
    
    '個装箱代　金額合計
'2009.06.17
'    excelSheet.Application.Cells(Row, 20).Value = "⑥商品化箱代合計"
'    If Len(SYUKA_SHEET_TITLE) < 4 Then
'        excelSheet.Application.Cells(Row, 20).Value = "商品化箱代合計"
'    Else
'        excelSheet.Application.Cells(Row, 20).Value = Mid(SYUKA_SHEET_TITLE, 4, 1) & "商品化箱代合計"
'    End If


    If Len(SYUKA_SHEET_TITLE) < 4 Then
        excelSheet.Application.Cells(Row, 16).Value = "商品化箱代合計"
    Else
        excelSheet.Application.Cells(Row, 16).Value = Mid(SYUKA_SHEET_TITLE, 4, 1) & "商品化箱代合計"
    End If

'2009.06.17
'    excelSheet.Application.Cells(Row + 1, 21).NumberFormatLocal = "#,##0_ "
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 21), excelSheet.Application.Cells(Row + 1, 21)).FormulaR1C1 = "=ROUNDUP(ROUNDDOWN(SUM(R[" & ((Row) * -1) + 3 & "]C:R[-2]C),2),0)"
    excelSheet.Application.Cells(Row + 1, 17).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 17), excelSheet.Application.Cells(Row + 1, 17)).FormulaR1C1 = "=ROUNDUP(ROUNDDOWN(SUM(R[" & ((Row) * -1) + 3 & "]C:R[-2]C),2),0)"

'2009.06.17
'    excelSheet.Application.ActiveCell.FormulaR1C1 = "=ROUNDUP(SUM(R[" & ((Row) * -1) + 3 & "]C:R[-1]C),0)"
'    excelSheet.Application.ActiveCell.FormulaR1C1 = "=ROUNDUP(ROUNDDOWN(SUM(R[" & ((Row) * -1) + 3 & "]C:R[-2]C),2),0)"
'2009.06.17
    
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 21), excelSheet.Application.Cells(Row + 1, 21)).HorizontalAlignment = xlRight
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 20), excelSheet.Application.Cells(Row, 21)).MergeCells = True
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 20), excelSheet.Application.Cells(Row + 1, 21)).MergeCells = True
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 17), excelSheet.Application.Cells(Row + 1, 17)).HorizontalAlignment = xlRight
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 16), excelSheet.Application.Cells(Row, 17)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 16), excelSheet.Application.Cells(Row + 1, 17)).MergeCells = True
    
    
    '罫線
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeLeft).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeBottom).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeRight).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlInsideHorizontal).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
'
'
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeLeft).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeBottom).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeRight).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlInsideHorizontal).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
'
'
'
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeLeft).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeBottom).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeRight).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlInsideHorizontal).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
'
'
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeLeft).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeBottom).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeRight).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeLeft).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeBottom).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeRight).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
'
'
'
'
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeLeft).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeBottom).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeRight).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
'
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeLeft).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeBottom).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeRight).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
'
'
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeLeft).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeBottom).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeRight).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
'
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 18), excelSheet.Application.Cells(Row + 1, 19)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 18), excelSheet.Application.Cells(Row + 1, 19)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 18), excelSheet.Application.Cells(Row + 1, 19)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 18), excelSheet.Application.Cells(Row + 1, 19)).Borders(xlEdgeLeft).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 18), excelSheet.Application.Cells(Row + 1, 19)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 18), excelSheet.Application.Cells(Row + 1, 19)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 18), excelSheet.Application.Cells(Row + 1, 19)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 18), excelSheet.Application.Cells(Row + 1, 19)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 18), excelSheet.Application.Cells(Row + 1, 19)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 18), excelSheet.Application.Cells(Row + 1, 19)).Borders(xlEdgeBottom).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 18), excelSheet.Application.Cells(Row + 1, 19)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 18), excelSheet.Application.Cells(Row + 1, 19)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 18), excelSheet.Application.Cells(Row + 1, 19)).Borders(xlEdgeRight).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 18), excelSheet.Application.Cells(Row + 1, 19)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
'
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 20), excelSheet.Application.Cells(Row + 1, 21)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 20), excelSheet.Application.Cells(Row + 1, 21)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 20), excelSheet.Application.Cells(Row + 1, 21)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 20), excelSheet.Application.Cells(Row + 1, 21)).Borders(xlEdgeLeft).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 20), excelSheet.Application.Cells(Row + 1, 21)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 20), excelSheet.Application.Cells(Row + 1, 21)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 20), excelSheet.Application.Cells(Row + 1, 21)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 20), excelSheet.Application.Cells(Row + 1, 21)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 20), excelSheet.Application.Cells(Row + 1, 21)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 20), excelSheet.Application.Cells(Row + 1, 21)).Borders(xlEdgeBottom).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 20), excelSheet.Application.Cells(Row + 1, 21)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 20), excelSheet.Application.Cells(Row + 1, 21)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 20), excelSheet.Application.Cells(Row + 1, 21)).Borders(xlEdgeRight).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 20), excelSheet.Application.Cells(Row + 1, 21)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
'
    
''''
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlDiagonalUp).LineStyle = xlNone
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlInsideHorizontal).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
    
    
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlDiagonalUp).LineStyle = xlNone
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 8)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
    





    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlDiagonalUp).LineStyle = xlNone
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlInsideHorizontal).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row + 1, 9)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic




    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 10)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 10)).Borders(xlDiagonalUp).LineStyle = xlNone
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 10)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 10)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 10)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 10)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 10)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 10)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 10)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 10)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 10)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row + 1, 10)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 11), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 11), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlDiagonalUp).LineStyle = xlNone
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 11), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 11), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 11), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 11), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 11), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 11), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 11), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 11), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 11), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 11), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 11), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 11), excelSheet.Application.Cells(Row + 1, 11)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlDiagonalUp).LineStyle = xlNone
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 12), excelSheet.Application.Cells(Row + 1, 13)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    
    
    
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlDiagonalUp).LineStyle = xlNone
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 14), excelSheet.Application.Cells(Row + 1, 14)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 15)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 15)).Borders(xlDiagonalUp).LineStyle = xlNone
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 15)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 15)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 15)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 15)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 15)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 15)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 15)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 15)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 15)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 15)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 15)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 15), excelSheet.Application.Cells(Row + 1, 15)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 16), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 16), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlDiagonalUp).LineStyle = xlNone
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 16), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 16), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 16), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 16), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 16), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 16), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 16), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 16), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 16), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 16), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 16), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 16), excelSheet.Application.Cells(Row + 1, 16)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlDiagonalUp).LineStyle = xlNone
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 17), excelSheet.Application.Cells(Row + 1, 17)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
    
    
    






''''



'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlDiagonalDown).LineStyle = xlNone
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlEdgeLeft).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlEdgeTop).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlEdgeTop).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlEdgeTop).ColorIndex = xlAutomatic
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlEdgeBottom).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlEdgeRight).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlEdgeRight).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlInsideVertical).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlInsideVertical).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlInsideVertical).ColorIndex = xlAutomatic
'
'
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlInsideHorizontal).Weight = xlThin
'    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(Row - 1, 21)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic

    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlDiagonalUp).LineStyle = xlNone

    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlEdgeLeft).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlEdgeTop).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlEdgeTop).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlEdgeTop).ColorIndex = xlAutomatic


    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlEdgeBottom).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlEdgeRight).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlEdgeRight).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlEdgeRight).ColorIndex = xlAutomatic

    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlInsideVertical).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlInsideVertical).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlInsideVertical).ColorIndex = xlAutomatic


    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlInsideHorizontal).Weight = xlThin
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 17)).Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
    
    
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 12), excelSheet.Application.Cells(Row + 1, 21)).Font.FontStyle = "太字"
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 12), excelSheet.Application.Cells(Row + 1, 21)).Font.Size = 14
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 10), excelSheet.Application.Cells(Row + 1, 17)).Font.FontStyle = "太字"
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, 10), excelSheet.Application.Cells(Row + 1, 17)).Font.Size = 14
    
    
    
    
'2009.07.30列幅を固定
'    excelSheet.Application.Columns("I:I").EntireColumn.AutoFit
'    excelSheet.Application.Columns(10).ColumnWidth = 16
'    excelSheet.Application.Columns(11).ColumnWidth = 16

    excelSheet.Application.Columns(1).ColumnWidth = 6
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row + 1, 1)).ShrinkToFit = True
    
    excelSheet.Application.Columns(6).ColumnWidth = 23
    excelSheet.Application.Columns(8).ColumnWidth = 20
    
    excelSheet.Application.Columns(9).ColumnWidth = 6
    excelSheet.Application.Columns(9).ShrinkToFit = True
    
    excelSheet.Application.Columns(10).ColumnWidth = 9
    excelSheet.Application.Columns(10).ShrinkToFit = True
    
    excelSheet.Application.Columns(11).ColumnWidth = 9
    excelSheet.Application.Columns(11).ShrinkToFit = True
    
    excelSheet.Application.Columns(12).ColumnWidth = 6
    excelSheet.Application.Columns(12).ShrinkToFit = True
    
    excelSheet.Application.Columns(14).ColumnWidth = 8
    excelSheet.Application.Columns(14).ShrinkToFit = True
    
    excelSheet.Application.Columns(15).ColumnWidth = 11
    excelSheet.Application.Columns(15).ShrinkToFit = True
    
    excelSheet.Application.Columns(16).ColumnWidth = 8
    excelSheet.Application.Columns(16).ShrinkToFit = True
    
    excelSheet.Application.Columns(17).ColumnWidth = 11
    excelSheet.Application.Columns(17).ShrinkToFit = True
    


'2009.07.30列幅を固定


    
        
    excelSheet.Application.Range("A4").Select
    excelSheet.Application.ActiveWindow.FreezePanes = True
        
    '印刷範囲 2009.07.30
    excelSheet.Application.ActiveSheet.PageSetup.PrintArea = "$A$1:$Q$" & Row + 1
        
        
excelApplication.Visible = True
    
    



    Set excelSheet = Nothing
    
    Set excelWorkBook = Nothing
    

'    excelApplication.Quit
    
    Set excelApplication = Nothing


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]出荷明細 EXCEL出力終了" & s_test_now & " " & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
    
    Call Input_UnLock
    
On Error GoTo 0
    
    
    SYU_DETAIL_Proc = False
    Exit Function
    
Error_Proc:
    
    MsgBox "処理異常　発生=" & Err.Number & "　処理を中断します。"
    
excelApplication.Visible = True
    
    



    Set excelSheet = Nothing
    
    Set excelWorkBook = Nothing
    

'    excelApplication.Quit
    
    Set excelApplication = Nothing


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]出荷明細 EXCEL出力異常終了" & s_test_now & " " & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
    
    Call Input_UnLock
    
On Error GoTo 0
    
    SYU_DETAIL_Proc = False

End Function
'Private Function SYU_Excel_Set_Proc(Row As Long, excelApplication As excel.Application, excelWorkBook As excel.Workbook, excelSheet As excel.Worksheet) As Integer
Private Function SYU_Excel_Set_Proc(Row As Long, excelApplication As Object, excelWorkBook As Object, excelSheet As Object) As Integer
'----------------------------------------------------------------------------
'           出荷データ--＞EXCEL
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim INV_F       As Boolean
    
Dim READ_NEXT   As Boolean
    
    
Dim wkS_KOUSU_BAIKA    As String   '2009.06.10
Dim wkS_SHIZAI_BAIKA   As String   '2009.06.10
    
    
    SYU_Excel_Set_Proc = True
        
    'セルの書式設定
''    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row, 1)).Select
''    excelSheet.Application.ActiveCell.FormulaR1C1 = "=ROW()-3"
    
''    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 5)).Select
''    excelSheet.Application.Selection.NumberFormatLocal = "@"
''    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 10)).Select
''    excelSheet.Application.Selection.NumberFormatLocal = "@"
    
    excelSheet.Application.Cells(Row, 1).Value = Row - 3
    
'    excelSheet.Application.Cells(Row, 2).NumberFormatLocal = "@"
'    excelSheet.Application.Cells(Row, 3).NumberFormatLocal = "@"
'    excelSheet.Application.Cells(Row, 4).NumberFormatLocal = "@"
'    excelSheet.Application.Cells(Row, 5).NumberFormatLocal = "@"
'    excelSheet.Application.Cells(Row, 10).NumberFormatLocal = "@"
    
    
    excelSheet.Application.Cells(Row, 2).NumberFormatLocal = "@"
    excelSheet.Application.Cells(Row, 3).NumberFormatLocal = "@"
    excelSheet.Application.Cells(Row, 4).NumberFormatLocal = "@"
    excelSheet.Application.Cells(Row, 5).NumberFormatLocal = "@"
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row, 5)).HorizontalAlignment = xlLeft
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 13), excelSheet.Application.Cells(Row, 13)).HorizontalAlignment = xlCenter
    
'    excelSheet.Application.Cells(Row, 9).NumberFormatLocal = "#,##0_ "
'    excelSheet.Application.Cells(Row, 12).NumberFormatLocal = "#,##0.00_ "
'    excelSheet.Application.Cells(Row, 13).NumberFormatLocal = "#,##0.00_ "
'    excelSheet.Application.Cells(Row, 15).NumberFormatLocal = "#,##0.00_ "
'    excelSheet.Application.Cells(Row, 16).NumberFormatLocal = "#,##0.00_ "
'    excelSheet.Application.Cells(Row, 18).NumberFormatLocal = "#,##0.00_ "
'    excelSheet.Application.Cells(Row, 19).NumberFormatLocal = "#,##0.00_ "
'    excelSheet.Application.Cells(Row, 20).NumberFormatLocal = "#,##0.00_ "
'    excelSheet.Application.Cells(Row, 21).NumberFormatLocal = "#,##0.00_ "

    excelSheet.Application.Cells(Row, 9).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Cells(Row, 10).NumberFormatLocal = "#,##0.00_ "
    excelSheet.Application.Cells(Row, 11).NumberFormatLocal = "#,##0.00_ "
    excelSheet.Application.Cells(Row, 14).NumberFormatLocal = "#,##0.00_ "
    excelSheet.Application.Cells(Row, 15).NumberFormatLocal = "#,##0.00_ "
    excelSheet.Application.Cells(Row, 16).NumberFormatLocal = "#,##0.00_ "
    excelSheet.Application.Cells(Row, 17).NumberFormatLocal = "#,##0.00_ "



    'ID-No
    excelSheet.Application.Cells(Row, 2).Value = Trim(StrConv(Y_SYUREC.ID_NO, vbUnicode))
    '出荷日
    excelSheet.Application.Cells(Row, 3).Value = Mid(StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode), 7, 2)
    '伝票№
    excelSheet.Application.Cells(Row, 4).Value = Trim(StrConv(Y_SYUREC.DEN_NO, vbUnicode))
    '出荷先
    excelSheet.Application.Cells(Row, 5).Value = Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))
    '出荷先名
    excelSheet.Application.Cells(Row, 6).Value = Trim(StrConv(Y_SYUREC.MUKE_NAME, vbUnicode))
    '品番
    excelSheet.Application.Cells(Row, 7).Value = Trim(StrConv(Y_SYUREC.HIN_NO, vbUnicode))
    '品名
    excelSheet.Application.Cells(Row, 8).Value = Trim(StrConv(Y_SYUREC.HIN_NAME, vbUnicode))
    '数量
    excelSheet.Application.Cells(Row, 9).Value = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode))
Debug.Print StrConv(Y_SYUREC.HIN_NO, vbUnicode) & " " & StrConv(Y_SYUREC.HAN_KBN, vbUnicode)
    '棚番
    If Trim(StrConv(Y_SYUREC.HTANABAN, vbUnicode)) = "" Then
        
        
        
        '2008.08.20 ↓
        If StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2" Then
    
            READ_NEXT = False
        
        
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                    '2009.06.10
                    If Not IsDate(Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 7, 2)) Then
                        
                        wkS_KOUSU_BAIKA = StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)
                        wkS_SHIZAI_BAIKA = StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)
                
                    Else
                
                        If StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) < StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode) Then
                
                            wkS_KOUSU_BAIKA = StrConv(ITEMREC.BEF_S_KOUSU_BAIKA, vbUnicode)
                            wkS_SHIZAI_BAIKA = StrConv(ITEMREC.BEF_S_SHIZAI_BAIKA, vbUnicode)
                
                        Else
                            wkS_KOUSU_BAIKA = StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)
                            wkS_SHIZAI_BAIKA = StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)
                        
                        End If
                
                    End If
                                    
                    'If Not IsNumeric(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) Then
                    If Not IsNumeric(wkS_KOUSU_BAIKA) Then
                    '2009.06.10
                            
                            
                        READ_NEXT = True
                
                    
                    Else
                        
                        
                        '海外単価使用
'                        excelSheet.Application.Cells(Row, 22).Value = "海外単価使用"
                        excelSheet.Application.Cells(Row, 18).Value = "海外単価使用"
                        
                        
                        
                        
                        
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
                                        Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                                        Exit Function
                                
                                End Select
                            
                            Case BtErrKeyNotFound
                            
                            
                                INV_F = True
                                            
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                                Exit Function
                        
                        End Select
                    End If
                
                Case BtErrKeyNotFound
                    
                    READ_NEXT = True

                
                    Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                    Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                
                    '2009.06.10
                    Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "00000000.00")
                    Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "00000000.00")
                    
                    
                    Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, "")
                    '2009.06.10
                
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            
            End Select
        
            '2009.07.28
            READ_NEXT = True
            excelSheet.Application.Cells(Row, 18).Value = ""
            '2009.07.28
        
            If READ_NEXT Then
                Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    
                    
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
                                        Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                                        Exit Function
                                
                                End Select
                            
                            Case BtErrKeyNotFound
                            
                            
                                INV_F = True
                                            
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                                Exit Function
                        
                        End Select
                    
                    
                    Case BtErrKeyNotFound
                        
                        INV_F = True
                    
                    
                        Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                        Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                        
                        '2009.06.10
                        Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "00000000.00")
                        Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "00000000.00")
                        
                        Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, "")
                        
                        '2009.06.10
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                
                End Select
            End If
        
        
        Else
        '2008.08.20 ↑
        
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
                    Call UniCode_Conv(ITEMREC.ST_REN, "")
                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
                
                    Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                    Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                
                    '2009.06.10
                    Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "00000000.00")
                    Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "00000000.00")
                    
                    Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, "")
                    
                    '2009.06.10
                
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
        End If
    Else
        
        
        '2008.08.20 ↓
        If StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2" Then
    
            READ_NEXT = False
        
        
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                    If Not IsNumeric(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) Then
                            
                        READ_NEXT = True
                
                    
                    Else
                        
                        '海外単価使用
'                        excelSheet.Application.Cells(Row, 22).Value = "海外単価使用"
                        excelSheet.Application.Cells(Row, 18).Value = "海外単価使用"
                        
                        
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
                                        Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                                        Exit Function
                                
                                End Select
                            
                            Case BtErrKeyNotFound
                            
                            
                                INV_F = True
                                            
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                                Exit Function
                        
                        End Select
                    End If
                
                Case BtErrKeyNotFound
                    
                    READ_NEXT = True

                
                    Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                    Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                
                    '2009.06.10
                    Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "00000000.00")
                    Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "00000000.00")
                    
                    Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, "")
                    
                    '2009.06.10
                
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            
            End Select
        
            '2009.07.28
            READ_NEXT = True
            excelSheet.Application.Cells(Row, 18).Value = ""
            '2009.07.28
        
            If READ_NEXT Then
                Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    
                    
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
                                        Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                                        Exit Function
                                
                                End Select
                            
                            Case BtErrKeyNotFound
                            
                            
                                INV_F = True
                                            
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                                Exit Function
                        
                        End Select
                    
                    
                    Case BtErrKeyNotFound
                        
                        INV_F = True
                    
                    
                        Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                        Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                        
                        '2009.06.10
                        Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "00000000.00")
                        Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "00000000.00")
                        
                        Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, "")
                        
                        '2009.06.10
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                
                End Select
            End If
        
        
        Else
        '2008.08.20 ↑
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                
                
                    Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                    Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                    '2009.06.10
                    Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "00000000.00")
                    Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "00000000.00")
                    
                    Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, "")
                    
                    '2009.06.10
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
        
        
            Call UniCode_Conv(ITEMREC.ST_SOKO, Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2))
            Call UniCode_Conv(ITEMREC.ST_RETU, Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 3, 2))
            Call UniCode_Conv(ITEMREC.ST_REN, Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 4, 2))
            Call UniCode_Conv(ITEMREC.ST_DAN, Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 6, 2))
        
        End If
    
    End If
    
    
'    excelSheet.Application.Cells(Row, 10).Value = Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) & _
'                                        Trim(StrConv(ITEMREC.ST_RETU, vbUnicode)) & _
'                                        Trim(StrConv(ITEMREC.ST_REN, vbUnicode)) & _
'                                        Trim(StrConv(ITEMREC.ST_DAN, vbUnicode))
    '出庫区分
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
            
                Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_IO_TANKA_No, "")
                Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_OUT_TANKA, "00000000.00")
            
                Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_Name, "")
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                Exit Function
        End Select
    End If
    
    '2009.06.10
    
    
If Trim(StrConv(Y_SYUREC.HIN_NO, vbUnicode)) = "304SPN-6" Then
    Debug.Print
End If
    
    
    If Not IsDate(Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 1, 4) & "/" & _
                    Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 5, 2) & "/" & _
                    Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 7, 2)) Then
        
        wkS_KOUSU_BAIKA = StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)
        wkS_SHIZAI_BAIKA = StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)

    Else



        If StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) < StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode) Then

            wkS_KOUSU_BAIKA = StrConv(ITEMREC.BEF_S_KOUSU_BAIKA, vbUnicode)
            wkS_SHIZAI_BAIKA = StrConv(ITEMREC.BEF_S_SHIZAI_BAIKA, vbUnicode)

        Else
            wkS_KOUSU_BAIKA = StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)
            wkS_SHIZAI_BAIKA = StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)
        
        End If

    End If
    '2009.06.10
    
    
    
    
    '出庫区分
'    excelSheet.Application.Cells(Row, 11).Value = Trim(StrConv(SE_LOC_TANKA_M_REC.SE_Name, vbUnicode))
    '出庫工料　単価
'    If IsNumeric(StrConv(SE_LOC_TANKA_M_REC.SE_OUT_TANKA, vbUnicode)) Then
'        excelSheet.Application.Cells(Row, 12).Value = CDbl(StrConv(SE_LOC_TANKA_M_REC.SE_OUT_TANKA, vbUnicode))
'    Else
'        excelSheet.Application.Cells(Row, 12).Value = 0
'    End If
    
    If IsNumeric(StrConv(SE_LOC_TANKA_M_REC.SE_OUT_TANKA, vbUnicode)) Then
        excelSheet.Application.Cells(Row, 10).Value = CDbl(StrConv(SE_LOC_TANKA_M_REC.SE_OUT_TANKA, vbUnicode))
    Else
        excelSheet.Application.Cells(Row, 10).Value = 0
    End If
    
    
    '出庫工料　金額
'    excelSheet.Application.Cells(Row, 13).Value = CDbl(StrConv(SE_LOC_TANKA_M_REC.SE_OUT_TANKA, vbUnicode))
    
    
    
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, k + 1), excelSheet.Application.Cells(Row, k + 1)).Select
'    excelSheet.Application.Cells(Row, k + 1).NumberFormatLocal = "#,##0.0_ "
'    excelSheet.Application.ActiveCell.FormulaR1C1 = "=RC[-1]*RC[" & -k + 1 & "]"
'    excelSheet.Application.Cells(Row, 13).FormulaR1C1 = "=RC[-1]"
    '出荷区分
    INV_F = False
    Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))
    Call UniCode_Conv(K0_MTS.SS_CODE, "")
    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    Select Case sts
        Case BtNoErr
        
        
            Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, StrConv(MTSREC.SYUKA_KBN, vbUnicode))
            sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
            Select Case sts
                Case BtNoErr
                
                
                Case BtErrKeyNotFound
                    INV_F = True
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                    Exit Function
            
            End Select
        
        
        
        Case BtErrKeyNotFound
            INV_F = True
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
            Exit Function
    
    End Select
    
                    
    If INV_F Then
        
        If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
        
            Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, INV_KBN11)
            sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
            Select Case sts
                Case BtNoErr
                    INV_F = False
                Case BtErrKeyNotFound
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                    Exit Function
            End Select
        
        
        End If
        
    End If
        
    If INV_F Then
        
        If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2" Then
        
            Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, INV_KBN12)
            sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
            Select Case sts
                Case BtNoErr
                    INV_F = False
                Case BtErrKeyNotFound
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                    Exit Function
            End Select
        
        
        End If
        
    End If
        
    If INV_F Then
        
        If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
        
            Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, INV_KBN71)
            sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
            Select Case sts
                Case BtNoErr
                    INV_F = False
                Case BtErrKeyNotFound
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                    Exit Function
            End Select
        
        
        End If
        
    End If
        
        
        
    If INV_F Then
        Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, INV_SYUKA_KBN)
        sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
            
                Call UniCode_Conv(SE_SHIP_TANKA_M_REC.SE_SYUKA_KBN, "")
                Call UniCode_Conv(SE_SHIP_TANKA_M_REC.SE_TANKA, "00000000.00")
            
            
                Call UniCode_Conv(SE_SHIP_TANKA_M_REC.SE_SYUKA_NAME, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                Exit Function
        End Select
    End If
                    
    
'    excelSheet.Application.Cells(Row, 14).Value = Trim(StrConv(SE_SHIP_TANKA_M_REC.SE_SYUKA_NAME, vbUnicode))
    
    
    '出荷工料　単価
'    If IsNumeric(StrConv(SE_SHIP_TANKA_M_REC.SE_TANKA, vbUnicode)) Then
'        excelSheet.Application.Cells(Row, 15).Value = CDbl(StrConv(SE_SHIP_TANKA_M_REC.SE_TANKA, vbUnicode))
'    Else
'        excelSheet.Application.Cells(Row, 15).Value = 0
'    End If
    If IsNumeric(StrConv(SE_SHIP_TANKA_M_REC.SE_TANKA, vbUnicode)) Then
        excelSheet.Application.Cells(Row, 11).Value = CDbl(StrConv(SE_SHIP_TANKA_M_REC.SE_TANKA, vbUnicode))
    Else
        excelSheet.Application.Cells(Row, 11).Value = 0
    End If
    '出荷工料　金額
'    excelSheet.Application.Cells(Row, 16).Value = CDbl(StrConv(SE_LOC_TANKA_M_REC.SE_OUT_TANKA, vbUnicode))
'    excelSheet.Application.Cells(Row, 16).FormulaR1C1 = "=RC[-1]"
    
    
    '個装形態
'    excelSheet.Application.Cells(Row, 17).Value = Trim(StrConv(Y_SYUREC.KOSO_KEITAI, vbUnicode))
    excelSheet.Application.Cells(Row, 12).Value = Trim(StrConv(Y_SYUREC.KOSO_KEITAI, vbUnicode))
    
    '切替区分   2009.06.10
    If Not IsDate(Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 1, 4) & "/" & _
                    Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 5, 2) & "/" & _
                    Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 7, 2)) Then
    Else
        If StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) < StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode) Then
        Else
            excelSheet.Application.Cells(Row, 13).Value = Trim(StrConv(ITEMREC.KIRIKAE_KBN, vbUnicode))
        End If
    End If
    '個装工料　単価
    If IsNumeric(wkS_KOUSU_BAIKA) Then
        '2009.06.10
        'excelSheet.Application.Cells(Row, 18).Value = CDbl(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode))
'        excelSheet.Application.Cells(Row, 18).Value = CDbl(wkS_KOUSU_BAIKA)
        excelSheet.Application.Cells(Row, 14).Value = CDbl(wkS_KOUSU_BAIKA)
    Else
'        excelSheet.Application.Cells(Row, 18).Value = 0
        excelSheet.Application.Cells(Row, 14).Value = 0
    End If
    '個装工料　金額
    
'    If IsNumeric(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) Then
'        excelSheet.Application.Cells(Row, 19).Value = CDbl(StrConv(Y_SYUREC.SURYO, vbUnicode)) * CDbl(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode))
'    Else
'        excelSheet.Application.Cells(Row, 19).Value = 0
'    End If
    
'     excelSheet.Application.Cells(Row, 19).Value = "=RC[-1]*RC[-10]"
     excelSheet.Application.Cells(Row, 15).Value = "=RC[-1]*RC[-6]"
    
    
    '個装箱代　単価
    If IsNumeric(wkS_SHIZAI_BAIKA) Then
        '2009.06.10
        'excelSheet.Application.Cells(Row, 20).Value = CDbl(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode))
'        excelSheet.Application.Cells(Row, 20).Value = CDbl(wkS_SHIZAI_BAIKA)
        excelSheet.Application.Cells(Row, 16).Value = CDbl(wkS_SHIZAI_BAIKA)
    Else
'        excelSheet.Application.Cells(Row, 20).Value = 0
        excelSheet.Application.Cells(Row, 16).Value = 0
    End If
    '個装箱代　金額
'    If IsNumeric(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)) Then
'        excelSheet.Application.Cells(Row, 21).Value = CDbl(StrConv(Y_SYUREC.SURYO, vbUnicode)) * CDbl(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode))
'    Else
'        excelSheet.Application.Cells(Row, 21).Value = 0
'    End If
'    excelSheet.Application.Cells(Row, 21).Value = "=RC[-1]*RC[-12]"
    excelSheet.Application.Cells(Row, 17).Value = "=RC[-1]*RC[-8]"
    
    
    '罫線
'    excelSheet.Range(Cells(Row, 1), Cells(Row, 21)).Select
'    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'    With Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With Selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With Selection.Borders(xlEdgeBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With Selection.Borders(xlInsideVertical)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With






















    If Not IsNumeric(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)) Then
        Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
    End If




    SYU_Excel_Set_Proc = False

End Function
Private Function KAMOKU_DETAIL_Proc() As Integer
'----------------------------------------------------------------------------
'                   ＥＸＣＥＬ（科目振替明細）出力
'                   2008.05.21
'----------------------------------------------------------------------------


Dim Row                 As Integer
    
    
Dim sts                 As Integer
Dim com                 As Integer
    
Dim i                   As Integer
    
Dim End_Date            As String

Dim s_test_now          As String

Dim Skip_F              As Boolean


'Dim excelApplication    As excel.Application       '2015.07.06
'Dim excelWorkBooks      As excel.Workbooks
'Dim excelWorkBook       As excel.Workbook          '2015.07.06
'Dim excelSheet          As excel.Worksheet         '2015.07.06
    
    
Dim excelApplication    As Object                   '2015.07.06
Dim excelWorkBook       As Object                   '2015.07.06
Dim excelSheet          As Object                   '2015.07.06
    
    
    
s_test_now = Format(Now, "YYYY/MM/DD HH:MM:SS")
    
    KAMOKU_DETAIL_Proc = True
    
    Call Input_Lock
    
    Set excelApplication = CreateObject("Excel.Application")
'''2008.05.16    excelApplication.Visible = True

        
    
    
    Set excelWorkBook = excelApplication.Workbooks.Add
    
    
    Set excelSheet = excelWorkBook.Worksheets(1)
    
    excelApplication.StandardFontSize = 11
    
    excelApplication.StandardFont = "ＭＳ　ゴシック"
    
    

    excelSheet.Application.Calculation = xlManual
    excelSheet.Application.MaxChange = 0.001

    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(1, 4)).Select
    With excelSheet.Application.Selection.Font
        .Size = 16
    End With
    excelSheet.Application.Cells(1, 1).Value = "科目振替明細表" & _
                                    Trim(StrConv(P_KANRIREC.CENTER_NAME, vbUnicode)) & _
                                    "（" & StrConv(Text1(ptxS_Date).Text, vbWide) & "～" & _
                                    StrConv(Text1(ptxE_Date).Text, vbWide) & "）"
    
    
    
    '列の幅
    excelSheet.Application.Columns(1).Select
    excelSheet.Application.Selection.ColumnWidth = 4.88
    'セルの結合
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 12), excelSheet.Application.Cells(2, 13)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 12), excelSheet.Application.Cells(2, 13)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 12), excelSheet.Application.Cells(2, 13)).MergeCells = True
   
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 15), excelSheet.Application.Cells(2, 16)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 15), excelSheet.Application.Cells(2, 16)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 15), excelSheet.Application.Cells(2, 16)).MergeCells = True
    
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 18), excelSheet.Application.Cells(2, 19)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 18), excelSheet.Application.Cells(2, 19)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 18), excelSheet.Application.Cells(2, 19)).MergeCells = True
    
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 20), excelSheet.Application.Cells(2, 21)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 20), excelSheet.Application.Cells(2, 21)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 20), excelSheet.Application.Cells(2, 21)).MergeCells = True
    
    '２行目見出し設定
    excelSheet.Application.Cells(2, 12).Value = "出庫工料"
    excelSheet.Application.Cells(2, 15).Value = "出荷工料"
    excelSheet.Application.Cells(2, 18).Value = "個装工料"
    excelSheet.Application.Cells(2, 20).Value = "個装箱代"
    '３行目見出し設定
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(3, 1)).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(3, 1).Value = "№"
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 2), excelSheet.Application.Cells(3, 21)).HorizontalAlignment = xlCenter
    excelSheet.Application.Cells(3, 2).Value = "ID-№"
    excelSheet.Application.Cells(3, 3).Value = "出荷日"
    excelSheet.Application.Cells(3, 4).Value = "伝№"
    excelSheet.Application.Cells(3, 5).Value = "出荷先"
    excelSheet.Application.Cells(3, 6).Value = "出荷先名"
    excelSheet.Application.Cells(3, 7).Value = "品番"
    excelSheet.Application.Cells(3, 8).Value = "品名"
    excelSheet.Application.Cells(3, 9).Value = "数量"
    excelSheet.Application.Cells(3, 10).Value = "棚番"
    excelSheet.Application.Cells(3, 11).Value = "出庫区分"
    excelSheet.Application.Cells(3, 12).Value = "単価"
    excelSheet.Application.Cells(3, 13).Value = "金額"
    excelSheet.Application.Cells(3, 14).Value = "出荷区分"
    excelSheet.Application.Cells(3, 15).Value = "単価"
    excelSheet.Application.Cells(3, 16).Value = "金額"
    excelSheet.Application.Cells(3, 17).Value = "個装形態"
    excelSheet.Application.Cells(3, 18).Value = "単価"
    excelSheet.Application.Cells(3, 19).Value = "金額"
    excelSheet.Application.Cells(3, 20).Value = "単価"
    excelSheet.Application.Cells(3, 21).Value = "金額"
    '見出し　罫線
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 21)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 12), excelSheet.Application.Cells(3, 13)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
   
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 15), excelSheet.Application.Cells(3, 16)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
   
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 18), excelSheet.Application.Cells(3, 21)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
   
    
    
    
    '右下にページ追加 2009.02.20
    excelSheet.Application.ActiveSheet.PageSetup.RightFooter = "&P/&N"
    'ページヘッダー固定 2009.02.20
    excelSheet.Application.ActiveSheet.PageSetup.PrintTitleRows = "$2:$3"
        
    
    
    
    
    Row = 3
        
    
    
    '------------------------------------------------------------------------   '過日分出荷予定の読み込み
    Call UniCode_Conv(K1_DEL_SYU.KEY_SYUKA_YMD, Format(Text1(ptxS_Date).Text, "YYYYMMDD"))
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        sts = BTRV(com, DEL_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K1_DEL_SYU, Len(K1_DEL_SYU), 1)
        Select Case sts
            Case BtNoErr
            
            
                If Format(Text1(ptxE_Date), "YYYYMMDD") < StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "出荷予定")
                Exit Function
        End Select

        Skip_F = False

        If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
            If (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Then

            Else
                Skip_F = True

            End If

        End If
    
    
    
    
    
        If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
        
            If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
            
                If Not IsNumeric(StrConv(Y_SYUREC.LK_SEQ_NO, vbUnicode)) Then
                    Skip_F = True
                End If
            End If
        
            If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2" Then
            
                If Trim(StrConv(Y_SYUREC.KAN_YMD, vbUnicode)) = "" Then
                    Skip_F = True
                End If
            End If
        
            If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
            
                If Trim(StrConv(Y_SYUREC.KAN_YMD, vbUnicode)) = "" Then
                    Skip_F = True
                End If
            End If
        
        
        End If
    
    
        If Not Skip_F Then
            If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" And _
                (StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "25" Or StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "29") Then
                If StrConv(Y_SYUREC.JGYOBU, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) And _
                    StrConv(Y_SYUREC.NAIGAI, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
            
            
            
                    Row = Row + 1
                
                    If KAMOKU_Excel_Set_Proc(Row, excelApplication, excelWorkBook, excelSheet) Then
                        Exit Function
                    End If
                
                End If
            End If
        End If
        
        
        com = BtOpGetNext
    Loop
    '------------------------------------------------------------------------   '出荷予定の読み込み
        
    Call UniCode_Conv(K0_Y_SYU.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, "")
    
    com = BtOpGetGreater
    
    Do
        
        DoEvents
        
        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "出荷予定")
                Exit Function
        End Select

        Skip_F = False
        
        
        If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
            If (StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1") Then

            Else
                Skip_F = True

            End If

        End If
        
        
        
        If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" Then
        
            If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
            
                If Not IsNumeric(StrConv(Y_SYUREC.LK_SEQ_NO, vbUnicode)) Then
                    Skip_F = True
                End If
            End If
        
            If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2" Then
            
                If Trim(StrConv(Y_SYUREC.KAN_YMD, vbUnicode)) = "" Then
                    Skip_F = True
                End If
            End If
        
            If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
            
                If Trim(StrConv(Y_SYUREC.KAN_YMD, vbUnicode)) = "" Then
                    Skip_F = True
                End If
            End If
        
        
        End If
    
        
        If Not Skip_F Then
            If StrConv(Y_SYUREC.JGYOBA, vbUnicode) = "00036003" And _
                (StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "25" Or StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "29") Then
            
            
                If Format(Text1(ptxS_Date).Text, "YYYYMMDD") > StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) Or _
                    Format(Text1(ptxE_Date).Text, "YYYYMMDD") < StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) Then
                Else
            
                    If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                        StrConv(Y_SYUREC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
                    Else
                        Row = Row + 1
                        If KAMOKU_Excel_Set_Proc(Row, excelApplication, excelWorkBook, excelSheet) Then
                            Exit Function
                        End If
                    End If
            
                End If
            End If
        End If
    
    
    
        com = BtOpGetNext
    Loop
    
    
    
    
    Row = Row + 1
    
    '合計
    excelSheet.Application.Cells(Row, 1).Value = "合計"
    
    '数量合計
    excelSheet.Application.Cells(Row, 9).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row, 9)).Select
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C)"
    
    '出庫工料　金額合計
    excelSheet.Application.Cells(Row, 13).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 13), excelSheet.Application.Cells(Row, 13)).Select
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C)"
    
    '出荷工料　金額合計
    excelSheet.Application.Cells(Row, 16).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 16), excelSheet.Application.Cells(Row, 16)).Select
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C)"
    
    '個装工料　金額合計
    excelSheet.Application.Cells(Row, 19).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 19), excelSheet.Application.Cells(Row, 19)).Select
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C)"
    
    '個装箱代　金額合計
    excelSheet.Application.Cells(Row, 21).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 21), excelSheet.Application.Cells(Row, 21)).Select
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C)"
    '罫線

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row, 21)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(Row, 21)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    
    
    
    excelSheet.Application.Columns("B:U").EntireColumn.AutoFit
    
    
    
        
    excelApplication.Visible = True
    
    



    Set excelSheet = Nothing
    
    Set excelWorkBook = Nothing
    

    
    
    Set excelApplication = Nothing


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        s_test_now & " " & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
    
    Call Input_UnLock
    KAMOKU_DETAIL_Proc = False
    

End Function
'Private Function KAMOKU_Excel_Set_Proc(Row As Integer, excelApplication As excel.Application, excelWorkBook As excel.Workbook, excelSheet As excel.Worksheet) As Integer
Private Function KAMOKU_Excel_Set_Proc(Row As Integer, excelApplication As Object, excelWorkBook As Object, excelSheet As Object) As Integer
'----------------------------------------------------------------------------
'           出荷データ--＞EXCEL
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim INV_F   As Boolean
    
    KAMOKU_Excel_Set_Proc = True
        
    'セルの書式設定
''    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row, 1)).Select
''    excelSheet.Application.ActiveCell.FormulaR1C1 = "=ROW()-3"
    
''    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 5)).Select
''    excelSheet.Application.Selection.NumberFormatLocal = "@"
''    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 10)).Select
''    excelSheet.Application.Selection.NumberFormatLocal = "@"
    
    excelSheet.Application.Cells(Row, 1).Value = Row - 3
    
    excelSheet.Application.Cells(Row, 2).NumberFormatLocal = "@"
    excelSheet.Application.Cells(Row, 3).NumberFormatLocal = "@"
    excelSheet.Application.Cells(Row, 4).NumberFormatLocal = "@"
    excelSheet.Application.Cells(Row, 5).NumberFormatLocal = "@"
    excelSheet.Application.Cells(Row, 10).NumberFormatLocal = "@"
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row, 5)).HorizontalAlignment = xlLeft
    
    excelSheet.Application.Cells(Row, 9).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Cells(Row, 12).NumberFormatLocal = "#,##0.00_ "
    
    excelSheet.Application.Cells(Row, 13).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Cells(Row, 15).NumberFormatLocal = "#,##0.00_ "
    excelSheet.Application.Cells(Row, 16).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Cells(Row, 18).NumberFormatLocal = "#,##0.00_ "
    excelSheet.Application.Cells(Row, 19).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Cells(Row, 20).NumberFormatLocal = "#,##0.00_ "
    excelSheet.Application.Cells(Row, 21).NumberFormatLocal = "#,##0_ "

    'ID-No
    excelSheet.Application.Cells(Row, 2).Value = Trim(StrConv(Y_SYUREC.ID_NO, vbUnicode))
    '出荷日
    excelSheet.Application.Cells(Row, 3).Value = Mid(StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode), 7, 2)
    '伝票№
    excelSheet.Application.Cells(Row, 4).Value = Trim(StrConv(Y_SYUREC.DEN_NO, vbUnicode))
    '出荷先
    excelSheet.Application.Cells(Row, 5).Value = Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))
    '出荷先名
    excelSheet.Application.Cells(Row, 6).Value = Trim(StrConv(Y_SYUREC.MUKE_NAME, vbUnicode))
    '品番
    excelSheet.Application.Cells(Row, 7).Value = Trim(StrConv(Y_SYUREC.HIN_NO, vbUnicode))
    '品名
    excelSheet.Application.Cells(Row, 8).Value = Trim(StrConv(Y_SYUREC.HIN_NAME, vbUnicode))
    '数量
    excelSheet.Application.Cells(Row, 9).Value = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode))
    '棚番
    If Trim(StrConv(Y_SYUREC.HTANABAN, vbUnicode)) = "" Then
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
            
                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                Call UniCode_Conv(ITEMREC.ST_RETU, "")
                Call UniCode_Conv(ITEMREC.ST_REN, "")
                Call UniCode_Conv(ITEMREC.ST_DAN, "")
            
                Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        End Select
    Else
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
            
            
                Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        End Select
    
    
        Call UniCode_Conv(ITEMREC.ST_SOKO, Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2))
        Call UniCode_Conv(ITEMREC.ST_RETU, Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 3, 2))
        Call UniCode_Conv(ITEMREC.ST_REN, Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 4, 2))
        Call UniCode_Conv(ITEMREC.ST_DAN, Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 6, 2))
    
    
    
    End If
    
    
    excelSheet.Application.Cells(Row, 10).Value = Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) & _
                                        Trim(StrConv(ITEMREC.ST_RETU, vbUnicode)) & _
                                        Trim(StrConv(ITEMREC.ST_REN, vbUnicode)) & _
                                        Trim(StrConv(ITEMREC.ST_DAN, vbUnicode))
    '出庫区分
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
            
                Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_IO_TANKA_No, "")
                Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_OUT_TANKA, "00000000.00")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                Exit Function
        End Select
    End If
    
    
    
    
    
    '出庫区分
    excelSheet.Application.Cells(Row, 11).Value = Trim(StrConv(SE_LOC_TANKA_M_REC.SE_IO_TANKA_No, vbUnicode))
    '出庫工料　単価
    If IsNumeric(StrConv(SE_LOC_TANKA_M_REC.SE_OUT_TANKA, vbUnicode)) Then
        excelSheet.Application.Cells(Row, 12).Value = CDbl(StrConv(SE_LOC_TANKA_M_REC.SE_OUT_TANKA, vbUnicode))
    Else
        excelSheet.Application.Cells(Row, 12).Value = 0
    End If
    '出庫工料　金額
    excelSheet.Application.Cells(Row, 13).Value = Int(CDbl(excelSheet.Cells(Row, 12).Value) + 0.9)
    '出荷区分
    INV_F = False
    Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))
    Call UniCode_Conv(K0_MTS.SS_CODE, "")
    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    Select Case sts
        Case BtNoErr
        
        
            Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, StrConv(MTSREC.SYUKA_KBN, vbUnicode))
            sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
            Select Case sts
                Case BtNoErr
                
                
                Case BtErrKeyNotFound
                    INV_F = True
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                    Exit Function
            
            End Select
        
        
        
        Case BtErrKeyNotFound
            INV_F = True
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
            Exit Function
    
    End Select
    
                    
    If INV_F Then
        
        If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
        
            Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, INV_KBN11)
            sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
            Select Case sts
                Case BtNoErr
                    INV_F = False
                Case BtErrKeyNotFound
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                    Exit Function
            End Select
        
        
        End If
        
    End If
        
    If INV_F Then
        
        If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2" Then
        
            Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, INV_KBN12)
            sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
            Select Case sts
                Case BtNoErr
                    INV_F = False
                Case BtErrKeyNotFound
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                    Exit Function
            End Select
        
        
        End If
        
    End If
        
    If INV_F Then
        
        If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
        
            Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, INV_KBN71)
            sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
            Select Case sts
                Case BtNoErr
                    INV_F = False
                Case BtErrKeyNotFound
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                    Exit Function
            End Select
        
        
        End If
        
    End If
        
        
        
    If INV_F Then
        Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, INV_SYUKA_KBN)
        sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
            
                Call UniCode_Conv(SE_SHIP_TANKA_M_REC.SE_SYUKA_KBN, "")
                Call UniCode_Conv(SE_SHIP_TANKA_M_REC.SE_TANKA, "00000000.00")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                Exit Function
        End Select
    End If
                    
    
    excelSheet.Application.Cells(Row, 14).Value = Trim(StrConv(SE_SHIP_TANKA_M_REC.SE_SYUKA_KBN, vbUnicode))
    
    
    '出荷工料　単価
    If IsNumeric(StrConv(SE_SHIP_TANKA_M_REC.SE_TANKA, vbUnicode)) Then
        excelSheet.Application.Cells(Row, 15).Value = CDbl(StrConv(SE_SHIP_TANKA_M_REC.SE_TANKA, vbUnicode))
    Else
        excelSheet.Application.Cells(Row, 15).Value = 0
    End If
    '出荷工料　金額
    excelSheet.Application.Cells(Row, 16).Value = Int(CDbl(excelSheet.Application.Cells(Row, 15).Value) + 0.9)
    '個装形態
    excelSheet.Application.Cells(Row, 17).Value = Trim(StrConv(Y_SYUREC.KOSO_KEITAI, vbUnicode))
    '個装工料　単価
    If IsNumeric(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) Then
        excelSheet.Application.Cells(Row, 18).Value = CDbl(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode))
    Else
        excelSheet.Application.Cells(Row, 18).Value = 0
    End If
    '個装工料　金額
    excelSheet.Application.Cells(Row, 19).Value = Int(CDbl(excelSheet.Application.Cells(Row, 9).Value) * CDbl(excelSheet.Application.Cells(Row, 18).Value) + 0.9)
    '個装箱代　単価
    If IsNumeric(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)) Then
        excelSheet.Application.Cells(Row, 20).Value = CDbl(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode))
    Else
        excelSheet.Application.Cells(Row, 20).Value = 0
    End If
    '個装箱代　金額
    excelSheet.Application.Cells(Row, 21).Value = Int(CDbl(excelSheet.Application.Cells(Row, 9).Value) * CDbl(excelSheet.Application.Cells(Row, 20).Value) + 0.9)
    '罫線
'    excelSheet.Range(Cells(Row, 1), Cells(Row, 21)).Select
'    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'    With Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With Selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With Selection.Borders(xlEdgeBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With Selection.Borders(xlInsideVertical)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With






















    If Not IsNumeric(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)) Then
        Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
    End If

    KAMOKU_Excel_Set_Proc = False

End Function

Private Function Cover_Total_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'           出荷データより鏡用の金額集計
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim INV_F       As Boolean
    
Dim READ_NEXT   As Boolean  '2008.08.20
    
Dim wkS_KOUSU_BAIKA    As String   '2009.06.10
Dim wkS_SHIZAI_BAIKA   As String   '2009.06.10
    
    
    
    Cover_Total_Proc = True
    
    
    Select Case Mode
        Case 1
'-------------------------------    出荷
            '出庫工料
            
            INV_F = False
            
            
            If Trim(StrConv(Y_SYUREC.HTANABAN, vbUnicode)) = "" Then
            
            
            
                '2008.08.20 ↓
                If StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2" Then
            
                    READ_NEXT = False
                
                
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI)
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                            '2009.06.10
                            If Not IsDate(Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 1, 4) & "/" & _
                                            Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 5, 2) & "/" & _
                                            Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 7, 2)) Then
                                
                                wkS_KOUSU_BAIKA = StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)
                                wkS_SHIZAI_BAIKA = StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)
                        
                            Else
                        
                                If StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) < StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode) Then
                        
                                    wkS_KOUSU_BAIKA = StrConv(ITEMREC.BEF_S_KOUSU_BAIKA, vbUnicode)
                                    wkS_SHIZAI_BAIKA = StrConv(ITEMREC.BEF_S_SHIZAI_BAIKA, vbUnicode)
                        
                                Else
                                    wkS_KOUSU_BAIKA = StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)
                                    wkS_SHIZAI_BAIKA = StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)
                                
                                End If
                        
                            End If
                                            
                            'If Not IsNumeric(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) Then
                            If Not IsNumeric(wkS_KOUSU_BAIKA) Then
                            '2009.06.10
                                    
                                READ_NEXT = True
                        
                            
                            Else
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
                                                Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                                                Exit Function
                                        
                                        End Select
                                    
                                    Case BtErrKeyNotFound
                                    
                                    
                                        INV_F = True
                                                    
                                    
                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                                        Exit Function
                                
                                End Select
                            End If
                        
                        Case BtErrKeyNotFound
                            
                            READ_NEXT = True

                        
                            Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                            Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                            '2009.06.10
                            Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "00000000.00")
                            Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "00000000.00")
                            '2009.06.10
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Exit Function
                    
                    End Select
                
                
                    If READ_NEXT Then
                        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            
                            
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
                                                Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                                                Exit Function
                                        
                                        End Select
                                    
                                    Case BtErrKeyNotFound
                                    
                                    
                                        INV_F = True
                                                    
                                    
                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                                        Exit Function
                                
                                End Select
                            
                            
                            Case BtErrKeyNotFound
                                
                                INV_F = True
                            
                            
                                Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                                Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                                '2009.06.10
                                Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "00000000.00")
                                Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "00000000.00")
                                '2009.06.10
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                Exit Function
                        
                        End Select
                    End If
                
                
                
                
                
                
                
                Else
                '2008.08.20 ↑
            
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        
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
                                            Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                                            Exit Function
                                    
                                    End Select
                                
                                Case BtErrKeyNotFound
                                
                                
                                    INV_F = True
                                                
                                
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                                    Exit Function
                            
                            End Select
                        
                        
                        Case BtErrKeyNotFound
                            
                            INV_F = True
                        
                        
                            Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                            Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                            '2009.06.10
                            Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "00000000.00")
                            Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "00000000.00")
                            '2009.06.10
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Exit Function
                    
                    End Select
                
                End If
            Else
                
                '2008.08.20 ↓
                If StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2" Then
            
                    READ_NEXT = False
                
                
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI)
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                            If Not IsNumeric(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) Then
                                    
                                READ_NEXT = True
                        
                            
                            Else
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
                                                Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                                                Exit Function
                                        
                                        End Select
                                    
                                    Case BtErrKeyNotFound
                                    
                                    
                                        INV_F = True
                                                    
                                    
                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                                        Exit Function
                                
                                End Select
                            End If
                        
                        Case BtErrKeyNotFound
                            
                            READ_NEXT = True

                        
                            Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                            Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                        
                            '2009.06.10
                            Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "00000000.00")
                            Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "00000000.00")
                            '2009.06.10
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Exit Function
                    
                    End Select
                
                
                    If READ_NEXT Then
                        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            
                            
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
                                                Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                                                Exit Function
                                        
                                        End Select
                                    
                                    Case BtErrKeyNotFound
                                    
                                    
                                        INV_F = True
                                                    
                                    
                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                                        Exit Function
                                
                                End Select
                            
                            
                            Case BtErrKeyNotFound
                                
                                INV_F = True
                            
                            
                                Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                                Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                            
                                '2009.06.10
                                Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "00000000.00")
                                Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "00000000.00")
                                '2009.06.10
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                Exit Function
                        
                        End Select
                    End If
                
                
                
                
                
                
                
                Else
                '2008.08.20 ↑
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        
                        
                        
                        Case BtErrKeyNotFound
                            
                        
                        
                            Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                            Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                            '2009.06.10
                            Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "00000000.00")
                            Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "00000000.00")
                            '2009.06.10
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Exit Function
                    
                    End Select
                
                
                
                
                    Call UniCode_Conv(K0_SOKO.Soko_No, Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2))
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
                                    Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                                    Exit Function
                            
                            End Select
                        
                        Case BtErrKeyNotFound
                        
                        
                            INV_F = True
                                        
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                            Exit Function
                    
                    End Select
                
                End If
            
            End If
            
            If INV_F Then
                Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, INV_IO_TANKA_No)
                sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    
                        Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_IO_TANKA_No, "")
                        Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_OUT_TANKA, "00000000.00")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                        Exit Function
                End Select
            End If
            
            '合計値　加算
            MEISAI_TBL(0).KINGAKU = MEISAI_TBL(0).KINGAKU + CDbl(StrConv(SE_LOC_TANKA_M_REC.SE_OUT_TANKA, vbUnicode))
            
            
            
            '出荷工料
            Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))
            Call UniCode_Conv(K0_MTS.SS_CODE, "")
            
            INV_F = False
                        
            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
            Select Case sts
                Case BtNoErr
                
                
                    Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, StrConv(MTSREC.SYUKA_KBN, vbUnicode))
                    sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        
                        Case BtErrKeyNotFound
                        
                            INV_F = True
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                            Exit Function
                    
                    End Select
                
                
                
                Case BtErrKeyNotFound
                    INV_F = True
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                    Exit Function
            
            End Select
            
            
            If INV_F Then
                
                If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
                
                    Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, INV_KBN11)
                    sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
                    Select Case sts
                        Case BtNoErr
                            INV_F = False
                        Case BtErrKeyNotFound
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                            Exit Function
                    End Select
                
                
                End If
                
            End If
            
            
            
            If INV_F Then
                
                If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "2" Then
                
                    Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, INV_KBN12)
                    sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
                    Select Case sts
                        Case BtNoErr
                            INV_F = False
                        Case BtErrKeyNotFound
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                            Exit Function
                    End Select
                
                
                End If
                
            End If
                
            If INV_F Then
                
                If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "7" And StrConv(Y_SYUREC.HAN_KBN, vbUnicode) = "1" Then
                
                    Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, INV_KBN71)
                    sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
                    Select Case sts
                        Case BtNoErr
                            INV_F = False
                        Case BtErrKeyNotFound
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                            Exit Function
                    End Select
                
                
                End If
                
            End If
                
                
                
            If INV_F Then
                Call UniCode_Conv(K0_SE_SHIP_TANKA_M.SE_SYUKA_KBN, INV_SYUKA_KBN)
                sts = BTRV(BtOpGetEqual, SE_SHIP_TANKA_M_POS, SE_SHIP_TANKA_M_REC, Len(SE_SHIP_TANKA_M_REC), K0_SE_SHIP_TANKA_M, Len(K0_SE_SHIP_TANKA_M), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    
                        Call UniCode_Conv(SE_SHIP_TANKA_M_REC.SE_SYUKA_KBN, "")
                        Call UniCode_Conv(SE_SHIP_TANKA_M_REC.SE_TANKA, "00000000.00")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "出荷先別単価設定")
                        Exit Function
                End Select
            End If
            
            '2009.06.10
            
            
        If Trim(StrConv(Y_SYUREC.HIN_NO, vbUnicode)) = "304SPN-6" Then
            Debug.Print
        End If
            
            
            If Not IsDate(Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 1, 4) & "/" & _
                            Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 5, 2) & "/" & _
                            Mid(StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode), 7, 2)) Then
                
                wkS_KOUSU_BAIKA = StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)
                wkS_SHIZAI_BAIKA = StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)
        
            Else
        
        
        
                If StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode) < StrConv(ITEMREC.TANKA_KIRIKAE_DT, vbUnicode) Then
        
                    wkS_KOUSU_BAIKA = StrConv(ITEMREC.BEF_S_KOUSU_BAIKA, vbUnicode)
                    wkS_SHIZAI_BAIKA = StrConv(ITEMREC.BEF_S_SHIZAI_BAIKA, vbUnicode)
        
                Else
                    wkS_KOUSU_BAIKA = StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)
                    wkS_SHIZAI_BAIKA = StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)
                
                End If
        
            End If
            '2009.06.10
            
            
            
            
            
            '合計値　加算
            MEISAI_TBL(1).KINGAKU = MEISAI_TBL(1).KINGAKU + CDbl(StrConv(SE_SHIP_TANKA_M_REC.SE_TANKA, vbUnicode))
            
            '商品化　工料
            '2009.06.10
            'If IsNumeric(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) Then
            '    MEISAI_TBL(2).KINGAKU = MEISAI_TBL(2).KINGAKU + CDbl(StrConv(Y_SYUREC.SURYO, vbUnicode)) * CDbl(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode))
            If IsNumeric(wkS_KOUSU_BAIKA) Then
                MEISAI_TBL(2).KINGAKU = MEISAI_TBL(2).KINGAKU + CDbl(StrConv(Y_SYUREC.SURYO, vbUnicode)) * CDbl(wkS_KOUSU_BAIKA)
            '2009.06.10
            Else
            End If
            '商品化　資材
            '2009.06.10
            'If IsNumeric(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)) Then
            '    MEISAI_TBL(3).KINGAKU = MEISAI_TBL(3).KINGAKU + CDbl(StrConv(Y_SYUREC.SURYO, vbUnicode)) * CDbl(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode))
            If IsNumeric(wkS_SHIZAI_BAIKA) Then
                MEISAI_TBL(3).KINGAKU = MEISAI_TBL(3).KINGAKU + CDbl(StrConv(Y_SYUREC.SURYO, vbUnicode)) * CDbl(wkS_SHIZAI_BAIKA)
            '2009.06.10
            Else
            End If
    
        Case 2
'-------------------------------    入庫
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_GLICSREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_GLICSREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_GLICSREC.HIN_NO, vbUnicode))
            
            INV_F = False
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                
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
                                    Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                                    Exit Function
                            
                            End Select
                        
                        Case BtErrKeyNotFound
                            INV_F = True
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                            Exit Function
                    
                    End Select
                
                
                Case BtErrKeyNotFound
                    INV_F = True
                
                    Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            
            End Select
            
            If INV_F Then
                Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, INV_IO_TANKA_No)
                sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_IO_TANKA_No, "")
                    
                        Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_IN_TANKA, "00000000.00")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                        Exit Function
                End Select
            End If
            
            
            '入庫　工料
            MEISAI_TBL(4).KINGAKU = MEISAI_TBL(4).KINGAKU + CDbl(StrConv(SE_LOC_TANKA_M_REC.SE_IN_TANKA, vbUnicode))
    
    
    
    
       Case 3
'-------------------------------    良品返品
'            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_GLICSREC.JGYOBU, vbUnicode))
'            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_GLICSREC.NAIGAI, vbUnicode))
'            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_GLICSREC.HIN_NO, vbUnicode))
            
'            INV_F = False
'            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'            Select Case sts
'                Case BtNoErr
'
'
'                    Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
'                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
'                    Select Case sts
'                        Case BtNoErr
'
'
                            Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, RYOHEN)
                            sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                                    Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_IN_TANKA, "00000000.00")
                                Case Else
                                   Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                                    Exit Function
                            
                            End Select
                        
'                        Case BtErrKeyNotFound
'                            Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_IN_TANKA, "00000000.00")
'
'                        Case Else
'                            Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
'                            Exit Function
'
'                    End Select
'
'
'                Case BtErrKeyNotFound
'                    Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_IN_TANKA, "00000000.00")
'
'                    Call UniCode_Conv(ITEMREC.HIN_NAME, "")
'
'                Case Else
'                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
'                    Exit Function
'
'            End Select
            
            If INV_F Then
                Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, INV_IO_TANKA_No)
                sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_IO_TANKA_No, "")
                    
                        Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_IN_TANKA, "00000000.00")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                        Exit Function
                End Select
            End If
            
            
            '入庫　工料
            MEISAI_TBL(5).KINGAKU = MEISAI_TBL(5).KINGAKU + CDbl(StrConv(SE_LOC_TANKA_M_REC.SE_IN_TANKA, vbUnicode))
        
    End Select
    
    Cover_Total_Proc = False

End Function
Private Function NYU_DETAIL_Proc() As Integer
'----------------------------------------------------------------------------
'                   ＥＸＣＥＬ（入庫明細）出力
'----------------------------------------------------------------------------


Dim Row                 As Integer
    
    
Dim sts                 As Integer
Dim com                 As Integer
    
Dim End_Date            As String

Dim s_test_now          As String

Dim Skip_Flg            As Boolean
    
Dim i                   As Integer
Dim j                   As Integer
    
    
'Dim excelApplication    As excel.Application   '2015.07.06
'Dim excelWorkBook       As excel.Workbook      '2015.07.06
'Dim excelSheet          As excel.Worksheet     '2015.07.06
    
Dim excelApplication    As Object               '2015.07.06
Dim excelWorkBook       As Object               '2015.07.06
Dim excelSheet          As Object               '2015.07.06
    
    
    
s_test_now = Format(Now, "YYYY/MM/DD HH:MM:SS")
    
    NYU_DETAIL_Proc = True
    
    Call Input_Lock
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[請求システム]入庫明細出力開始" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
    
    
    Set excelApplication = CreateObject("Excel.Application")
    
    

'2008.05.16 excelApplication.Visible = True

        
    
    
    Set excelWorkBook = excelApplication.Workbooks.Add
    
    
    Set excelSheet = excelWorkBook.Worksheets(1)
    
    
    
    
    excelApplication.StandardFontSize = 11
    
    excelApplication.StandardFont = "ＭＳ　ゴシック"
    
    

    excelSheet.Application.Calculation = xlManual
    excelSheet.Application.MaxChange = 0.001


    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(1, 4)).Select
    With excelSheet.Application.Selection.Font
                
        .Size = 16
    End With
    
    excelSheet.Application.Cells(1, 1).Value = "入庫明細表　" & _
                                    Trim(StrConv(P_KANRIREC.CENTER_NAME, vbUnicode)) & _
                                    "（" & Text1(ptxS_Date).Text & "～" & _
                                    Text1(ptxE_Date).Text & "）"
    
    
    
    
    
    
    '列の幅
    excelSheet.Application.Columns(1).Select
    excelSheet.Application.Selection.ColumnWidth = 4.88
    'セルの結合
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 10), excelSheet.Application.Cells(2, 11)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 10), excelSheet.Application.Cells(2, 11)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 10), excelSheet.Application.Cells(2, 11)).MergeCells = True
    
    '２行目見出し設定
    excelSheet.Application.Cells(2, 10).Value = "入庫工料"
    '３行目見出し設定
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(3, 1)).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(3, 1).Value = "№"
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 2), excelSheet.Application.Cells(3, 11)).HorizontalAlignment = xlCenter
    excelSheet.Application.Cells(3, 2).Value = "入庫日"
    excelSheet.Application.Cells(3, 3).Value = "伝№"
    excelSheet.Application.Cells(3, 4).Value = "相手先"
    excelSheet.Application.Cells(3, 5).Value = "品番"
    excelSheet.Application.Cells(3, 6).Value = "品名"
    excelSheet.Application.Cells(3, 7).Value = "数量"
    excelSheet.Application.Cells(3, 8).Value = "棚番"
    excelSheet.Application.Cells(3, 9).Value = "入庫区分"
    excelSheet.Application.Cells(3, 10).Value = "単価"
    excelSheet.Application.Cells(3, 11).Value = "金額"
    '見出し　罫線
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 11)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 10), excelSheet.Application.Cells(3, 11)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    '右下にページ追加 2009.02.20
    excelSheet.Application.ActiveSheet.PageSetup.RightFooter = "&P/&N"
    'ｼｰﾄ名変更　2009.02.20
'    excelSheet.Application.ActiveSheet.NAME = "⑦入庫"             2009.06.17
    excelSheet.Application.ActiveSheet.NAME = NYUKO_SHEET_TITLE     '2009.06.17
    'ページヘッダー固定 2009.02.20
    excelSheet.Application.ActiveSheet.PageSetup.PrintTitleRows = "$2:$3"
    
    
    '余白
    excelSheet.Application.ActiveSheet.PageSetup.LeftMargin = excelSheet.Application.InchesToPoints(0)
    excelSheet.Application.ActiveSheet.PageSetup.RightMargin = excelSheet.Application.InchesToPoints(0)
    excelSheet.Application.ActiveSheet.PageSetup.TopMargin = excelSheet.Application.InchesToPoints(0)
    excelSheet.Application.ActiveSheet.PageSetup.BottomMargin = excelSheet.Application.InchesToPoints(0.393700787401575)
    
    excelSheet.Application.ActiveSheet.PageSetup.HeaderMargin = excelSheet.Application.InchesToPoints(0)
    excelSheet.Application.ActiveSheet.PageSetup.FooterMargin = excelSheet.Application.InchesToPoints(0)
    
    
    '印刷　横→縦
'    excelSheet.Application.ActiveSheet.PageSetup.Orientation = xlLandscape
    excelSheet.Application.ActiveSheet.PageSetup.Orientation = xlPortrait
    

    
    '印刷　拡大率
    excelSheet.Application.ActiveSheet.PageSetup.Zoom = False
    excelSheet.Application.ActiveSheet.PageSetup.FitToPagesWide = 1
    excelSheet.Application.ActiveSheet.PageSetup.FitToPagesTall = False
    
    '印刷　中央 2009.07.30
    excelSheet.Application.ActiveSheet.PageSetup.CenterHorizontally = True
    
    '枠線なし   2009.06.19
    excelSheet.Application.ActiveWindow.DisplayGridlines = False
    
    Row = 3
    
    '------------------------------------------------------------------------   'Y_GLICSの読み込み
        
    Call UniCode_Conv(K0_Y_GLICS.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_Y_GLICS.SYUKA_YMD, Format(Text1(ptxS_Date).Text, "YYYYMMDD"))
    Call UniCode_Conv(K0_Y_GLICS.TEXT_NO, "")
    
    com = BtOpGetGreater
    
    Do
        
        DoEvents
        
        sts = BTRV(com, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(Y_GLICSREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Then
                    Exit Do
                End If
            
            
                If Format(Text1(ptxE_Date).Text, "YYYYMMDD") < StrConv(Y_GLICSREC.SYUKA_YMD, vbUnicode) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "Y_GLICS")
                Exit Function
        End Select

        Skip_Flg = True
        For i = 0 To UBound(JGYOBU_T)               '収支区分のチェック
            If StrConv(Y_GLICSREC.JGYOBU, vbUnicode) = JGYOBU_T(i).CODE Then
                For j = 0 To UBound(Soko_T, 2)
                    If Trim(StrConv(Y_GLICSREC.H_SOKO, vbUnicode)) = Trim(Soko_T(i, j).HS_SOKO) Then
                        Skip_Flg = False
                        Exit For
                    End If
                Next j
                Exit For
            End If
        Next i
    
    

    
        '2008.11.27 "4"追加
        If StrConv(Y_GLICSREC.IO_KBN, vbUnicode) <> "1" And StrConv(Y_GLICSREC.IO_KBN, vbUnicode) <> "4" Then
            Skip_Flg = True
        End If
    
    
        If StrConv(Y_GLICSREC.PM_KBN, vbUnicode) = "-" Then
            Skip_Flg = True
        End If
    
    
'        If Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) = "36003" Then
'            Skip_Flg = True
'        End If
    
'        If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "PP" Then
'            Skip_Flg = True
'        End If
'
'
'
'
'        Select Case StrConv(Y_GLICSREC.JGYOBU, vbUnicode)
'            Case SOJIKI                         '掃除機
'
'
'                If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "KM" Then
'                    Skip_Flg = True
'                End If
'
'                If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "KK" Then
'                    Skip_Flg = True
'                End If
'
'                If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "GG" Then
'                    Skip_Flg = True
'                End If
'
'                If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "SS" Then
'                    Skip_Flg = True
'                End If
'
'                If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 5) = "0090K" Then
'                    Skip_Flg = True
'                End If
'
'                If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 5) = "0092H" Then
'                    Skip_Flg = True
'                End If
'
'                If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "AA" Then
'                    Skip_Flg = True
'                End If
'
'
'
'            Case DENKA, SUIHAN, SENTAKU         '電化、炊飯、洗濯機（アイロン）
'
'
'                Select Case MyCenter
'
'                    Case "O"
'
'                        If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "01" Then
'                            Skip_Flg = True
'                        End If
'
'                        If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 3) = "H33" Then
'                            Skip_Flg = True
'                        End If
'                        If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 3) = "H22" Then
'                            Skip_Flg = True
'                        End If
'
'                        If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "05" Then
'                            Skip_Flg = True
'                        End If
'
'                        If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "08" Then
'                            Skip_Flg = True
'                        End If
'
'                        If StrConv(Y_GLICSREC.JGYOBU, vbUnicode) = DENKA Then
'
'                            If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) <> "02" And _
'                                Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) <> "G11" And _
'                                Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) <> "G22" Then
'                                Skip_Flg = True
'                            End If
'                        End If
'
'                        If (StrConv(Y_GLICSREC.JGYOBU, vbUnicode) = SUIHAN Or _
'                            StrConv(Y_GLICSREC.JGYOBU, vbUnicode) = SENTAKU) Then
'                            If (Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "P3" Or _
'                                Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "S3") Then
'                                Skip_Flg = True
'                            End If
'                        End If
'
'
'
'                        If (StrConv(Y_GLICSREC.JGYOBU, vbUnicode) = SUIHAN Or _
'                            StrConv(Y_GLICSREC.JGYOBU, vbUnicode) = SENTAKU) Then
'                            If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "RO" Then
'                                Skip_Flg = True
'                            End If
'                        End If
'
'                        If (StrConv(Y_GLICSREC.JGYOBU, vbUnicode) = SUIHAN Or _
'                            StrConv(Y_GLICSREC.JGYOBU, vbUnicode) = SENTAKU) Then
'                            If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "07" Then
'                                Skip_Flg = True
'                            End If
'                        End If
'
'
'
'
'
'
'                    Case "F"
'
'                        If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "P2" Then
'                            Skip_Flg = True
'                        End If''''''
'
'                        If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "U2" Then
'                            Skip_Flg = True
'                        End If
'
'
'                        If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 3) <> "904" Then
'                            If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 1) = "9" Then
'                              Skip_Flg = True
'                            End If
'                        End If
'
'                End Select
'        End Select
        
        
        
        If Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) = "01B11" Or _
            Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) = "01C11" Then
        Else
            Skip_Flg = True
        End If
        
        
        
        
        
        
        
        If Not Skip_Flg Then
    
            
            If StrConv(Y_GLICSREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                StrConv(Y_GLICSREC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
            Else
        
                Row = Row + 1
                
If Right(Format(Row - 3, 0), 2) = "00" Or Right(Format(Row - 3, 0), 2) = "50" Then
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]入庫明細 出力中　" & "出力件数　= " & Row - 3, Me.hwnd, 0)
    DoEvents
End If
                
                
                
                If NYU_Excel_Set_Proc(Row, excelApplication, excelWorkBook, excelSheet) Then
                    Exit Function
                End If
            End If
        
        End If
    
    
    
    
        com = BtOpGetNext
    Loop
    
    
    
    
    
    Row = Row + 1
    
    '合計
    excelSheet.Application.Cells(Row, 1).Value = "合計"
    
    '数量合計
 '   excelSheet.Application.Cells(Row, 7).NumberFormatLocal = "#,##0_ "
 '   excelSheet.Application.Range(excelSheet.Application.Cells(Row, 7), excelSheet.Application.Cells(Row, 7)).Select
 '   excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C)"
    
    '出庫工料　金額合計
 '   excelSheet.Application.Cells(Row, 11).NumberFormatLocal = "#,##0.00_ "
 '   excelSheet.Application.Range(excelSheet.Application.Cells(Row, 11), excelSheet.Application.Cells(Row, 11)).Select
 '   excelSheet.Application.ActiveCell.FormulaR1C1 = "=ROUNDUP(SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C),0)"
    
    
 '   excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).HorizontalAlignment = xlRight
 '   excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).MergeCells = True
    
    
    '罫線
 '   excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row, 11)).Select
 '   excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
 '   excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
 '   With excelSheet.Application.Selection.Borders(xlEdgeLeft)
 '       .LineStyle = xlContinuous
 '       .Weight = xlThin
 '       .ColorIndex = xlAutomatic
 '   End With
 '   With excelSheet.Application.Selection.Borders(xlEdgeTop)
 '       .LineStyle = xlContinuous
 '       .Weight = xlThin
 '       .ColorIndex = xlAutomatic
 '   End With
 '   With excelSheet.Application.Selection.Borders(xlEdgeBottom)
 '       .LineStyle = xlContinuous
 '       .Weight = xlThin
 '       .ColorIndex = xlAutomatic
 '   End With
 '   With excelSheet.Application.Selection.Borders(xlEdgeRight)
 '       .LineStyle = xlContinuous
 '       .Weight = xlThin
 '       .ColorIndex = xlAutomatic
 '   End With
 '   With excelSheet.Application.Selection.Borders(xlInsideVertical)
 '       .LineStyle = xlContinuous
 '       .Weight = xlThin
 '       .ColorIndex = xlAutomatic
 '   End With
    
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(Row - 1, 11)).Select
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(Row - 1, 11)).Select
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(Row - 1, 11)).Select
    
    
    
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
     End With
    With excelSheet.Application.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
         .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row, 1)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
     End With
    With excelSheet.Application.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
         .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 3)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
         .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 4), excelSheet.Application.Cells(Row, 7)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
         .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    
    'セルの結合
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 8), excelSheet.Application.Cells(Row, 9)).HorizontalAlignment = xlRight
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 8), excelSheet.Application.Cells(Row, 9)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 8), excelSheet.Application.Cells(Row, 9)).MergeCells = True
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).HorizontalAlignment = xlRight
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).VerticalAlignment = xlCenter
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 11)).MergeCells = True
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 8), excelSheet.Application.Cells(Row, 11)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
         .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 11)).HorizontalAlignment = xlRight
    
    '入庫件数
    excelSheet.Application.Cells(Row, 2).Value = "入庫件数"
    excelSheet.Application.Cells(Row, 3).Value = Format(Row - 4, "#,##0")
    
    '入庫個数
    excelSheet.Application.Cells(Row, 6).Value = "入庫個数"
    excelSheet.Application.Cells(Row, 7).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 7), excelSheet.Application.Cells(Row, 7)).Select
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C)"
    
    '出庫工料　金額合計

'2009.06.17
'    excelSheet.Application.Cells(Row, 8).Value = "⑦入庫費用合計"
    If Len(NYUKO_SHEET_TITLE) < 1 Then
        excelSheet.Application.Cells(Row, 8).Value = "入庫費用合計"
    Else
        excelSheet.Application.Cells(Row, 8).Value = Left(NYUKO_SHEET_TITLE, 1) & "入庫費用合計"
    End If
'2009.06.17
    
    
    
    
    
    
    
    excelSheet.Application.Cells(Row, 11).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 11), excelSheet.Application.Cells(Row, 11)).Select
    
    
    With excelSheet.Application.Selection.Font
        .FontStyle = "太字"
        .Size = 14
    End With
    
    
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=ROUNDUP(ROUNDDOWN(SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C),2),0)"
    
    
    
    excelSheet.Application.Columns("B:U").EntireColumn.AutoFit
    
    excelSheet.Application.Range("A4").Select
    excelSheet.Application.ActiveWindow.FreezePanes = True
    
    '印刷範囲 2009.07.30
    excelSheet.Application.ActiveSheet.PageSetup.PrintArea = "$A$1:$K$" & Row
    
    
    
    excelApplication.Visible = True


    



    Set excelSheet = Nothing
    
    Set excelWorkBook = Nothing
    

    
    Set excelApplication = Nothing

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[請求システム]入庫明細出力開始" & s_test_now & " " & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)



    
    Call Input_UnLock
    NYU_DETAIL_Proc = False
    

End Function


'Private Function NYU_Excel_Set_Proc(Row As Integer, excelApplication As excel.Application, excelWorkBook As excel.Workbook, excelSheet As excel.Worksheet) As Integer
Private Function NYU_Excel_Set_Proc(Row As Integer, excelApplication As Object, excelWorkBook As Object, excelSheet As Object) As Integer


'----------------------------------------------------------------------------
'           Y_GLICS--＞EXCEL
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim INV_F       As Boolean
    
Dim ST_SOKO     As String * 2
Dim ST_RETU     As String * 2
Dim ST_REN      As String * 2
Dim ST_DAN      As String * 2
    
    
    NYU_Excel_Set_Proc = True
        
    'セルの書式設定
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row, 1)).Select
'    excelSheet.Application.ActiveCell.FormulaR1C1 = "=ROW()-3"
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 3)).Select
'    excelSheet.Application.Selection.NumberFormatLocal = "@"
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 8), excelSheet.Application.Cells(Row, 7)).Select
'    excelSheet.Application.Selection.NumberFormatLocal = "@"
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row, 7)).Select
'    excelSheet.Application.Selection.NumberFormatLocal = "@"
    
    excelSheet.Application.Cells(Row, 1).Value = Row - 3
    
    excelSheet.Application.Cells(Row, 2).NumberFormatLocal = "@"
    excelSheet.Application.Cells(Row, 3).NumberFormatLocal = "@"
    excelSheet.Application.Cells(Row, 8).NumberFormatLocal = "@"
    excelSheet.Application.Cells(Row, 9).NumberFormatLocal = "@"
    
    
    excelSheet.Application.Cells(Row, 7).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Cells(Row, 10).NumberFormatLocal = "#,##0.00_ "
    excelSheet.Application.Cells(Row, 11).NumberFormatLocal = "#,##0.00_ "

    '出荷日(入庫日)
    excelSheet.Application.Cells(Row, 2).Value = Mid(StrConv(Y_GLICSREC.SYUKA_YMD, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(Y_GLICSREC.SYUKA_YMD, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(Y_GLICSREC.SYUKA_YMD, vbUnicode), 7, 2)
    '伝票№
    excelSheet.Application.Cells(Row, 3).Value = Trim(StrConv(Y_GLICSREC.DEN_NO, vbUnicode))
    '相手先
    excelSheet.Application.Cells(Row, 4).Value = Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode))
    '品番
    excelSheet.Application.Cells(Row, 5).Value = Trim(StrConv(Y_GLICSREC.HIN_NO, vbUnicode))
    '品名
    excelSheet.Application.Cells(Row, 6).Value = Trim(StrConv(Y_GLICSREC.HIN_NAME, vbUnicode))
    '数量
    excelSheet.Application.Cells(Row, 7).Value = CLng(StrConv(Y_GLICSREC.SURYO, vbUnicode))
    '棚番
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_GLICSREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_GLICSREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_GLICSREC.HIN_NO, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
        
            Call UniCode_Conv(ITEMREC.ST_SOKO, "")
            Call UniCode_Conv(ITEMREC.ST_RETU, "")
            Call UniCode_Conv(ITEMREC.ST_REN, "")
            Call UniCode_Conv(ITEMREC.ST_DAN, "")
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    End Select
    excelSheet.Cells(Row, 8).Value = Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) & _
                                        Trim(StrConv(ITEMREC.ST_RETU, vbUnicode)) & _
                                        Trim(StrConv(ITEMREC.ST_REN, vbUnicode)) & _
                                        Trim(StrConv(ITEMREC.ST_DAN, vbUnicode))
    '入庫区分
    
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
                Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_IO_TANKA_No, "")
            
                Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_Name, "")
            
                Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_IN_TANKA, "00000000.00")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                Exit Function
        End Select
    End If
    
    
    '入庫区分
    excelSheet.Application.Cells(Row, 9).Value = Trim(StrConv(SE_LOC_TANKA_M_REC.SE_Name, vbUnicode))
    '入庫工料　単価
    If IsNumeric(StrConv(SE_LOC_TANKA_M_REC.SE_IN_TANKA, vbUnicode)) Then
        excelSheet.Application.Cells(Row, 10).Value = CDbl(StrConv(SE_LOC_TANKA_M_REC.SE_IN_TANKA, vbUnicode))
    Else
        excelSheet.Application.Cells(Row, 10).Value = 0
    End If
    '入庫工料　金額
'    excelSheet.Application.Cells(Row, 11).Value = Int(CDbl(excelSheet.Application.Cells(Row, 10).Value) + 0.9)
    excelSheet.Application.Cells(Row, 11).FormulaR1C1 = "=RC[-1]"


    NYU_Excel_Set_Proc = False

End Function


Private Function RYOHEN_Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           良品返品--＞Grid
'----------------------------------------------------------------------------

Dim sts     As Integer
Dim INV_F   As Boolean
    
    RYOHEN_Grid_Set_Proc = True

    

    SEIKYU.ReDim Min_Row, Row, Min_Col, Max_Col
    
    '伝票日付
    SEIKYU(Row, ColSYUKA_YMD) = Mid(StrConv(Y_GLICSREC.SYUKA_YMD, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(Y_GLICSREC.SYUKA_YMD, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(Y_GLICSREC.SYUKA_YMD, vbUnicode), 7, 2)
    
    
    '伝票№
    SEIKYU(Row, ColDEN_NO) = StrConv(Y_GLICSREC.DEN_NO, vbUnicode)
    '出荷先
    SEIKYU(Row, ColMUKE_CODE) = StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)
    '品番
    SEIKYU(Row, ColHIN_GAI) = StrConv(Y_GLICSREC.HIN_NO, vbUnicode)
        
    
    
    '数量
    SEIKYU(Row, ColDEN_NO) = Format(CLng(StrConv(Y_GLICSREC.SURYO, vbUnicode)), "#0")
    
    '出庫工料
    
    SEIKYU(Row, ColSYUKA_KOURYO) = ""
    '出荷工料
    SEIKYU(Row, ColSYUKO_KOURYO) = ""
    
    '入庫工料
'    INV_F = False
'    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_GLICSREC.JGYOBU, vbUnicode))
'    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_GLICSREC.NAIGAI, vbUnicode))
'    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_GLICSREC.HIN_NO, vbUnicode))
'    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'    Select Case sts
'        Case BtNoErr
'
'
'            Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
'            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
'            Select Case sts
'                Case BtNoErr
                
                
                    Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, RYOHEN)
                    sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            INV_F = True
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "入出庫単価単価設定マスタ")
                            Exit Function
                    
                    End Select
                
'                Case BtErrKeyNotFound
'                    INV_F = True
'
'                Case Else
'                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
'                    Exit Function
'
'            End Select
'
'
'        Case BtErrKeyNotFound
'            INV_F = True
'
'            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
'
'        Case Else
'            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
'            Exit Function
'
'    End Select
    
    If INV_F Then
        Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, INV_IO_TANKA_No)
        sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_IO_TANKA_No, "")
            
                Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_IN_TANKA, "00000000.00")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                Exit Function
        End Select
    End If
    
    SEIKYU(Row, ColRYOHEN_KOURYO) = Format(CDbl(StrConv(SE_LOC_TANKA_M_REC.SE_IN_TANKA, vbUnicode)), "#0.00")
    '合計値　加算
    GK_RYOHEN_KOURYO = GK_RYOHEN_KOURYO + Int(CDbl(StrConv(SE_LOC_TANKA_M_REC.SE_IN_TANKA, vbUnicode)) + 0.9)
    
    
    
    '品名
    SEIKYU(Row, ColHIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
    
    '良品返品
    SEIKYU(Row, ColRYOHEN_KOURYO) = ""
    '商品化　工料
    SEIKYU(Row, ColSYOHIN_KOURYO) = ""
    '商品化　資材
    SEIKYU(Row, ColSYOHIN_SHIZAI) = ""


    RYOHEN_Grid_Set_Proc = False

End Function


Private Function RYOHEN_DETAIL_Proc() As Integer
'----------------------------------------------------------------------------
'                   ＥＸＣＥＬ（良品返品）出力
'----------------------------------------------------------------------------


Dim Row                 As Integer
    
    
Dim sts                 As Integer
Dim com                 As Integer
    
Dim End_Date            As String

Dim s_test_now          As String

Dim Skip_Flg            As Boolean
    
Dim i                   As Integer
Dim j                   As Integer
    
    
'Dim excelApplication    As excel.Application   '2015.07.06
'Dim excelWorkBook       As excel.Workbook      '2015.07.06
'Dim excelSheet          As excel.Worksheet     '2015.07.06
Dim excelApplication    As Object               '2015.07.06
Dim excelWorkBook       As Object               '2015.07.06
Dim excelSheet          As Object               '2015.07.06
    
    
s_test_now = Format(Now, "YYYY/MM/DD HH:MM:SS")
    
    RYOHEN_DETAIL_Proc = True
    
    Call Input_Lock
    
    Set excelApplication = CreateObject("Excel.Application")
'2008.05.16    excelApplication.Visible = True

        
    
    
    Set excelWorkBook = excelApplication.Workbooks.Add
    
    
    Set excelSheet = excelWorkBook.Worksheets(1)
    
    
    
    
    excelApplication.StandardFontSize = 11
    
    excelApplication.StandardFont = "ＭＳ　ゴシック"
    
    

    excelSheet.Application.Calculation = xlManual
    excelSheet.Application.MaxChange = 0.001


    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(1, 4)).Select
    With excelSheet.Application.Selection.Font
                
        .Size = 16
    End With
    
    excelSheet.Application.Cells(1, 1).Value = "良品返品明細表" & _
                                    Trim(StrConv(P_KANRIREC.CENTER_NAME, vbUnicode)) & _
                                    "（" & StrConv(Text1(ptxS_Date).Text, vbWide) & "～" & _
                                    StrConv(Text1(ptxE_Date).Text, vbWide) & "）"
    
    
    
    '列の幅
    excelSheet.Application.Columns(1).Select
    excelSheet.Application.Selection.ColumnWidth = 4.88
    'セルの結合
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 10), excelSheet.Application.Cells(2, 11)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 10), excelSheet.Application.Cells(2, 11)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 10), excelSheet.Application.Cells(2, 11)).MergeCells = True
    
    '２行目見出し設定
    excelSheet.Application.Cells(2, 10).Value = "良品返品工料"
    '３行目見出し設定
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(3, 1)).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(3, 1).Value = "№"
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 2), excelSheet.Application.Cells(3, 11)).HorizontalAlignment = xlCenter
    excelSheet.Application.Cells(3, 2).Value = "入庫日"
    excelSheet.Application.Cells(3, 3).Value = "伝№"
    excelSheet.Application.Cells(3, 4).Value = "相手先"
    excelSheet.Application.Cells(3, 5).Value = "品番"
    excelSheet.Application.Cells(3, 6).Value = "品名"
    excelSheet.Application.Cells(3, 7).Value = "数量"
    excelSheet.Application.Cells(3, 8).Value = "棚番"
    excelSheet.Application.Cells(3, 9).Value = "入庫区分"
    excelSheet.Application.Cells(3, 10).Value = "単価"
    excelSheet.Application.Cells(3, 11).Value = "金額"
    '見出し　罫線
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 11)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 10), excelSheet.Application.Cells(3, 11)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    
    
    
    '右下にページ追加 2009.02.20
    excelSheet.Application.ActiveSheet.PageSetup.RightFooter = "&P/&N"
    'ページヘッダー固定 2009.02.20
    excelSheet.Application.ActiveSheet.PageSetup.PrintTitleRows = "$2:$3"
    
    
    
    
    
    
    Row = 3
    
    '------------------------------------------------------------------------   'Y_GLICSの読み込み
        
    Call UniCode_Conv(K0_Y_GLICS.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_Y_GLICS.SYUKA_YMD, Format(Text1(ptxS_Date).Text, "YYYYMMDD"))
    Call UniCode_Conv(K0_Y_GLICS.TEXT_NO, "")
    
    com = BtOpGetGreater
    
    Do
        
        DoEvents
        
        sts = BTRV(com, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(Y_GLICSREC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Then
                    Exit Do
                End If
            
            
                If Format(Text1(ptxE_Date).Text, "YYYYMMDD") < StrConv(Y_GLICSREC.SYUKA_YMD, vbUnicode) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "Y_GLICS")
                Exit Function
        End Select

        Skip_Flg = True
        For i = 0 To UBound(JGYOBU_T)               '収支区分のチェック
            If StrConv(Y_GLICSREC.JGYOBA, vbUnicode) = JGYOBU_T(i).CODE Then
                For j = 0 To UBound(Soko_T, 2)
                    If Trim(Y_GLICSREC.H_SOKO) = Trim(Soko_T(i, j).HS_SOKO) Then
                        Skip_Flg = False
                        Exit For
                    End If
                Next j
                Exit For
            End If
        Next i
    
        
        If StrConv(Y_GLICSREC.IO_KBN, vbUnicode) <> "1" Then
            Skip_Flg = True
        End If
    
    
        If StrConv(Y_GLICSREC.PM_KBN, vbUnicode) = "-" Then
            Skip_Flg = True
        End If
    
    
'        If Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) = "36003" Then
'            Skip_Flg = True
'        End If
'
'        If Left(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode), 2) = "PP" Then
'            Skip_Flg = True
'        End If
        
        
        
        If Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) = "0221B" Or _
            Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode)) = "0221C" Then
        Else
            Skip_Flg = True
        End If
        
        
        
        
        If Not Skip_Flg Then
    
            
            If StrConv(Y_GLICSREC.NAIGAI, vbUnicode) = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Then
            
            
                Row = Row + 1
                If RYOHEN_Excel_Set_Proc(Row, excelApplication, excelWorkBook, excelSheet) Then
                    Exit Function
                End If
            End If
        End If
    
    
    
    
        com = BtOpGetNext
    Loop
    
    
    Row = Row + 1
    
    '合計
    excelSheet.Application.Cells(Row, 1).Value = "合計"
    
    '数量合計
    excelSheet.Application.Cells(Row, 7).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 7), excelSheet.Application.Cells(Row, 7)).Select
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C)"
    
    '出庫工料　金額合計
    excelSheet.Application.Cells(Row, 11).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 11), excelSheet.Application.Cells(Row, 11)).Select
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C)"
    
    
    '罫線
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row, 11)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(Row, 11)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    
    
    
    excelSheet.Application.Columns("B:U").EntireColumn.AutoFit
    
    excelApplication.Visible = True


    



    Set excelSheet = Nothing
    
    Set excelWorkBook = Nothing
    

    
    Set excelApplication = Nothing


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        s_test_now & " " & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
    
    Call Input_UnLock
    RYOHEN_DETAIL_Proc = False
    

End Function

'Private Function RYOHEN_Excel_Set_Proc(Row As Integer, excelApplication As excel.Application, excelWorkBook As excel.Workbook, excelSheet As excel.Worksheet) As Integer
Private Function RYOHEN_Excel_Set_Proc(Row As Integer, excelApplication As Object, excelWorkBook As Object, excelSheet As Object) As Integer


'----------------------------------------------------------------------------
'           Y_GLICS--＞EXCEL
'----------------------------------------------------------------------------
Dim INV_F       As Boolean
Dim sts         As Integer
    
Dim ST_SOKO     As String * 2
Dim ST_RETU     As String * 2
Dim ST_REN      As String * 2
Dim ST_DAN      As String * 2
    
    
    RYOHEN_Excel_Set_Proc = True
        
    'セルの書式設定
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row, 1)).Select
'    excelSheet.Application.ActiveCell.FormulaR1C1 = "=ROW()-3"
    
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 3)).Select
'    excelSheet.Application.Selection.NumberFormatLocal = "@"
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 8), excelSheet.Application.Cells(Row, 7)).Select
'    excelSheet.Application.Selection.NumberFormatLocal = "@"
'    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row, 7)).Select
'    excelSheet.Application.Selection.NumberFormatLocal = "@"
    
    excelSheet.Application.Cells(Row, 1).Value = Row - 3
    
    excelSheet.Application.Cells(Row, 2).NumberFormatLocal = "@"
    excelSheet.Application.Cells(Row, 3).NumberFormatLocal = "@"
    excelSheet.Application.Cells(Row, 8).NumberFormatLocal = "@"
    excelSheet.Application.Cells(Row, 9).NumberFormatLocal = "@"
    
    
    excelSheet.Application.Cells(Row, 7).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Cells(Row, 10).NumberFormatLocal = "#,##0.00_ "
    excelSheet.Application.Cells(Row, 11).NumberFormatLocal = "#,##0_ "

    '出荷日(入庫日)
    excelSheet.Application.Cells(Row, 2).Value = Mid(StrConv(Y_GLICSREC.SYUKA_YMD, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(Y_GLICSREC.SYUKA_YMD, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(Y_GLICSREC.SYUKA_YMD, vbUnicode), 7, 2)
    '伝票№
    excelSheet.Application.Cells(Row, 3).Value = Trim(StrConv(Y_GLICSREC.DEN_NO, vbUnicode))
    '相手先
    excelSheet.Application.Cells(Row, 4).Value = Trim(StrConv(Y_GLICSREC.YOSAN_FROM, vbUnicode))
    '品番
    excelSheet.Application.Cells(Row, 5).Value = Trim(StrConv(Y_GLICSREC.HIN_NO, vbUnicode))
    '品名
    excelSheet.Application.Cells(Row, 6).Value = Trim(StrConv(Y_GLICSREC.HIN_NAME, vbUnicode))
    '数量
    excelSheet.Application.Cells(Row, 7).Value = CLng(StrConv(Y_GLICSREC.SURYO, vbUnicode))
    '棚番
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_GLICSREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_GLICSREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_GLICSREC.HIN_NO, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
        
            Call UniCode_Conv(ITEMREC.ST_SOKO, "")
            Call UniCode_Conv(ITEMREC.ST_RETU, "")
            Call UniCode_Conv(ITEMREC.ST_REN, "")
            Call UniCode_Conv(ITEMREC.ST_DAN, "")
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    End Select
    excelSheet.Cells(Row, 8).Value = Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) & _
                                        Trim(StrConv(ITEMREC.ST_RETU, vbUnicode)) & _
                                        Trim(StrConv(ITEMREC.ST_REN, vbUnicode)) & _
                                        Trim(StrConv(ITEMREC.ST_DAN, vbUnicode))
    '入庫区分
    
    INV_F = False
'    Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
'    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
'    Select Case sts
'        Case BtNoErr
        
            Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, RYOHEN)
            sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                
                    INV_F = True
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                    Exit Function
            End Select
        
'        Case BtErrKeyNotFound
'
'            INV_F = True
'
'        Case Else
'            Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
'            Exit Function
'
'    End Select
    
    
    If INV_F Then
        Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, INV_IO_TANKA_No)
        sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_IO_TANKA_No, "")
            
                Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_IN_TANKA, "00000000.00")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "入出庫単価設定マスタ")
                Exit Function
        End Select
    End If
    
    
    '入庫区分
    excelSheet.Application.Cells(Row, 9).Value = Trim(StrConv(SE_LOC_TANKA_M_REC.SE_IO_TANKA_No, vbUnicode))
    '入庫工料　単価
    If IsNumeric(StrConv(SE_LOC_TANKA_M_REC.SE_IN_TANKA, vbUnicode)) Then
        excelSheet.Application.Cells(Row, 10).Value = CDbl(StrConv(SE_LOC_TANKA_M_REC.SE_IN_TANKA, vbUnicode))
    Else
        excelSheet.Application.Cells(Row, 10).Value = 0
    End If
    '入庫工料　金額
    excelSheet.Application.Cells(Row, 11).Value = Int(CDbl(excelSheet.Application.Cells(Row, 10).Value) + 0.9)


    RYOHEN_Excel_Set_Proc = False
End Function


' ------------------------------------------------------------------------
'       指定した精度の数値に切り上げします。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に切り上げられた数値。
' ------------------------------------------------------------------------
Private Function ToRoundUp(ByVal dValue As Double, ByVal iDigits As Integer) As Double
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

