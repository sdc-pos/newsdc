VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PR000201 
   Caption         =   "注文残検索([PR00020] 2012.03.01 16:30) 大阪POS部材対応"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16110
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
   ScaleHeight     =   10305
   ScaleWidth      =   16110
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   10395
      MaxLength       =   10
      TabIndex        =   4
      Top             =   240
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   7980
      MaxLength       =   10
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   6195
      MaxLength       =   10
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   3270
      MaxLength       =   10
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1470
      MaxLength       =   10
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   1455
      MaxLength       =   5
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   2175
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   6
      Top             =   840
      Width           =   2775
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
      TabIndex        =   18
      Top             =   9720
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
      Left            =   9600
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   9720
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   9720
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
      TabIndex        =   15
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ﾃﾞｰﾀ"
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   9720
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9720
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
      Index           =   5
      Left            =   4920
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "検 索"
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
      Left            =   4080
      TabIndex        =   11
      Top             =   9720
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
      Left            =   2760
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9720
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9720
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9720
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
      Index           =   0
      Left            =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid2 
      Height          =   7935
      Left            =   210
      TabIndex        =   20
      Top             =   1320
      Width           =   15750
      _ExtentX        =   27781
      _ExtentY        =   13996
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "注文日"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "注文№"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "注文先名"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "資材品番"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "品名"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "注文数"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "注文残"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "在庫残"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "希望納期日"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "回答納期日"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "使用月"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2170"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2064"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2275"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2170"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=3836"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=3731"
      Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=0"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=2699"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2593"
      Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=0"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=3387"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=3281"
      Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=0"
      Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(25)=   "Column(5).Width=1773"
      Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=1667"
      Splits(0)._ColumnProps(28)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(30)=   "Column(6).Width=1773"
      Splits(0)._ColumnProps(31)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(6)._WidthInPix=1667"
      Splits(0)._ColumnProps(33)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(34)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(35)=   "Column(7).Width=1773"
      Splits(0)._ColumnProps(36)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(7)._WidthInPix=1667"
      Splits(0)._ColumnProps(38)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(39)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(40)=   "Column(8).Width=2699"
      Splits(0)._ColumnProps(41)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(8)._WidthInPix=2593"
      Splits(0)._ColumnProps(43)=   "Column(8)._ColStyle=0"
      Splits(0)._ColumnProps(44)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(45)=   "Column(9).Width=2699"
      Splits(0)._ColumnProps(46)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(9)._WidthInPix=2593"
      Splits(0)._ColumnProps(48)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(49)=   "Column(10).Width=1508"
      Splits(0)._ColumnProps(50)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(10)._WidthInPix=1402"
      Splits(0)._ColumnProps(52)=   "Column(10).Order=11"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFF&,.bold=0,.fontsize=1200"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=128"
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
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFF00&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=43,.parent=1,.bold=0,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(25)  =   ":id=43,.strikethrough=0,.charset=128"
      _StyleDefs(26)  =   ":id=43,.fontname=ＭＳ ゴシック"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=45,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=47,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=58,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(39)  =   ":id=58,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(40)  =   ":id=58,.fontname=ＭＳ ゴシック"
      _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=44"
      _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=45"
      _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=47"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=16,.parent=43"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=13,.parent=44"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=14,.parent=45"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=15,.parent=47"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=28,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(49)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(50)  =   ":id=28,.fontname=ＭＳ ゴシック"
      _StyleDefs(51)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=44"
      _StyleDefs(52)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=45"
      _StyleDefs(53)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=47"
      _StyleDefs(54)  =   "Splits(0).Columns(3).Style:id=66,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(55)  =   ":id=66,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(56)  =   ":id=66,.fontname=ＭＳ ゴシック"
      _StyleDefs(57)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=44"
      _StyleDefs(58)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=45"
      _StyleDefs(59)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=47"
      _StyleDefs(60)  =   "Splits(0).Columns(4).Style:id=32,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(61)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(62)  =   ":id=32,.fontname=ＭＳ ゴシック"
      _StyleDefs(63)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=44"
      _StyleDefs(64)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=45"
      _StyleDefs(65)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=47"
      _StyleDefs(66)  =   "Splits(0).Columns(5).Style:id=74,.parent=43,.alignment=1"
      _StyleDefs(67)  =   "Splits(0).Columns(5).HeadingStyle:id=71,.parent=44"
      _StyleDefs(68)  =   "Splits(0).Columns(5).FooterStyle:id=72,.parent=45"
      _StyleDefs(69)  =   "Splits(0).Columns(5).EditorStyle:id=73,.parent=47"
      _StyleDefs(70)  =   "Splits(0).Columns(6).Style:id=20,.parent=43,.alignment=1"
      _StyleDefs(71)  =   "Splits(0).Columns(6).HeadingStyle:id=17,.parent=44"
      _StyleDefs(72)  =   "Splits(0).Columns(6).FooterStyle:id=18,.parent=45"
      _StyleDefs(73)  =   "Splits(0).Columns(6).EditorStyle:id=19,.parent=47"
      _StyleDefs(74)  =   "Splits(0).Columns(7).Style:id=24,.parent=43,.alignment=1"
      _StyleDefs(75)  =   "Splits(0).Columns(7).HeadingStyle:id=21,.parent=44"
      _StyleDefs(76)  =   "Splits(0).Columns(7).FooterStyle:id=22,.parent=45"
      _StyleDefs(77)  =   "Splits(0).Columns(7).EditorStyle:id=23,.parent=47"
      _StyleDefs(78)  =   "Splits(0).Columns(8).Style:id=62,.parent=43,.alignment=0"
      _StyleDefs(79)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=44"
      _StyleDefs(80)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=45"
      _StyleDefs(81)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=47"
      _StyleDefs(82)  =   "Splits(0).Columns(9).Style:id=70,.parent=43"
      _StyleDefs(83)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=44"
      _StyleDefs(84)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=45"
      _StyleDefs(85)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=47"
      _StyleDefs(86)  =   "Splits(0).Columns(10).Style:id=78,.parent=43"
      _StyleDefs(87)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=44"
      _StyleDefs(88)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=45"
      _StyleDefs(89)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=47"
      _StyleDefs(90)  =   "Named:id=33:Normal"
      _StyleDefs(91)  =   ":id=33,.parent=0"
      _StyleDefs(92)  =   "Named:id=34:Heading"
      _StyleDefs(93)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(94)  =   ":id=34,.wraptext=-1"
      _StyleDefs(95)  =   "Named:id=35:Footing"
      _StyleDefs(96)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(97)  =   "Named:id=36:Selected"
      _StyleDefs(98)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(99)  =   "Named:id=37:Caption"
      _StyleDefs(100) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(101) =   "Named:id=38:HighlightRow"
      _StyleDefs(102) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(103) =   "Named:id=39:EvenRow"
      _StyleDefs(104) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(105) =   "Named:id=40:OddRow"
      _StyleDefs(106) =   ":id=40,.parent=33"
      _StyleDefs(107) =   "Named:id=41:RecordSelector"
      _StyleDefs(108) =   ":id=41,.parent=34"
      _StyleDefs(109) =   "Named:id=42:FilterBar"
      _StyleDefs(110) =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "使用月"
      Height          =   255
      Index           =   4
      Left            =   9555
      TabIndex        =   25
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "～"
      Height          =   255
      Index           =   3
      Left            =   7560
      TabIndex        =   24
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "希望納期日"
      Height          =   255
      Index           =   2
      Left            =   4830
      TabIndex        =   23
      Top             =   360
      Width           =   1260
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "注文日"
      Height          =   255
      Index           =   0
      Left            =   630
      TabIndex        =   22
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "～"
      Height          =   255
      Index           =   1
      Left            =   2790
      TabIndex        =   21
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "注文先"
      Height          =   255
      Index           =   6
      Left            =   615
      TabIndex        =   19
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "PR000201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'ラベル用添字
Private Const plblY_NOUKI1% = 2           '2007.12.05
Private Const plblY_NOUKI2% = 3           '2007.12.05

Private Const plblUSE_YM% = 4               '2007.12.05


'テキスト用添字
Private Const ptxS_ORDER_DT% = 0            '注文日　開始
Private Const ptxE_ORDER_DT% = 1            '注文日　終了

Private Const ptxS_Y_NOUKI_DT% = 2          '希望納期 開始　2008.01.10
Private Const ptxE_Y_NOUKI_DT% = 3          '希望納期 開始　2008.01.10

Private Const ptxUSE_YM% = 4                '使用月 2008.01.10


Private Const ptxORDER_CODE% = 5            '注文先ｺｰﾄﾞ

'コンボ用添字
Private Const pcmbORDER% = 0                '注文先


'---------------    注文残用    2007.07.27


Private Z_SHORDER  As New XArrayDB



Private Const Z_Min_Row% = 1                '最小行数
Private Const Z_Min_Col% = 0                '最小列数
Private Const Z_Max_Col% = 10               '最大列数   2007.12.05 8-->10
    
Private Const colZ_ORDER_DT% = 0            '注文日時
Private Const colZ_ORDER_NO% = 1          'CODE
Private Const colZ_ORDER_NAME% = 2          '注文先名
Private Const colZ_HIN_GAI% = 3             '資材品番
Private Const colZ_HIN_NAME% = 4            '品名
Private Const colZ_ORDER_QTY% = 5           '注文数
Private Const colZ_ZAN_QTY% = 6             '注文残
Private Const colZ_ZAIKO_QTY% = 7           '在庫数
Private Const colZ_Y_NOUKI_DT% = 8          '予定納期
                                            
Private Const colZ_ANS_NOUKI_DT% = 9        '回答納期日 2007.12.05
Private Const colZ_USE_YM% = 10             '使用月     2007.12.05
                                            
                                            
Private Z_Sort_Tbl(colZ_ORDER_DT To colZ_USE_YM) _
                As Integer                  'ｿｰﾄの制御 0:昇順 1:降順
Private Z_Tbl_Set_F   As Boolean
'---------------    注文残用    2007.07.27

Private P_SHORDER_CSV   As String           '2007.11.08




Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PR000201.MousePointer = vbHourglass

    Call Ctrl_Lock(PR000201)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PR000201)


    PR000201.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts         As Integer
Dim com         As Integer
Dim i           As Integer
    
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        
        
        Case ptxS_ORDER_DT      '注文日　開始
        
            If Trim(Text1(ptxS_ORDER_DT).Text) <> "" Then
                If IsDate(Text1(ptxS_ORDER_DT).Text) Then
                    Text1(ptxS_ORDER_DT).Text = Format(CDate(Text1(ptxS_ORDER_DT).Text), "YYYY/MM/DD")
                Else
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            Else
            End If
        
        Case ptxE_ORDER_DT      '注文日　終了
            
            If Trim(Text1(ptxE_ORDER_DT).Text) <> "" Then
                If IsDate(Text1(ptxE_ORDER_DT).Text) Then
                    Text1(ptxE_ORDER_DT).Text = Format(CDate(Text1(ptxE_ORDER_DT).Text), "YYYY/MM/DD")
                Else
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            
            
                If Text1(ptxS_ORDER_DT).Text > Text1(ptxE_ORDER_DT).Text Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(ptxS_ORDER_DT).SetFocus
                    Exit Function
                End If
            
            
            Else
            End If
            
            
        Case ptxS_Y_NOUKI_DT      '希望納期日　開始   2007.12.05
        
            If Trim(Text1(ptxS_Y_NOUKI_DT).Text) <> "" Then
                If IsDate(Text1(ptxS_Y_NOUKI_DT).Text) Then
                    Text1(ptxS_Y_NOUKI_DT).Text = Format(CDate(Text1(ptxS_Y_NOUKI_DT).Text), "YYYY/MM/DD")
                Else
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            Else
            End If
        
        Case ptxE_Y_NOUKI_DT      '希望納期日　終了   2007.12.05
            
            If Trim(Text1(ptxE_Y_NOUKI_DT).Text) <> "" Then
                If IsDate(Text1(ptxE_Y_NOUKI_DT).Text) Then
                    Text1(ptxE_Y_NOUKI_DT).Text = Format(CDate(Text1(ptxE_Y_NOUKI_DT).Text), "YYYY/MM/DD")
                Else
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            
            
                If Text1(ptxE_Y_NOUKI_DT).Text > Text1(ptxE_Y_NOUKI_DT).Text Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(ptxS_ORDER_DT).SetFocus
                    Exit Function
                End If
            
            
            Else
            End If
            
            
        
            If (Trim(Text1(ptxS_ORDER_DT).Text) <> "" Or Trim(Text1(ptxE_ORDER_DT).Text) <> "") Then
                If (Trim(Text1(ptxS_Y_NOUKI_DT).Text) <> "" Or Trim(Text1(ptxE_Y_NOUKI_DT).Text) <> "") Then
        
                    MsgBox "注文日、希望納期日の同時指定はできません。"
                    Text1(ptxS_ORDER_DT).SetFocus
                    Exit Function
            
                End If
            End If
        
        
        
        Case ptxUSE_YM          '使用月     2007.12.05
            
            If Trim(Text1(ptxUSE_YM).Text) <> "" Then
                If IsDate(Text1(ptxUSE_YM).Text & "/01") Then
                    Text1(ptxUSE_YM).Text = Left(Format(CDate(Text1(ptxUSE_YM).Text) & "/01", "YYYY/MM/DD"), 7)
                Else
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            
            
            
            Else
            End If
        
        
        Case ptxORDER_CODE      '注文先
            Combo1(pcmbORDER).ListIndex = -1
            For i = 0 To Combo1(pcmbORDER).ListCount - 1
                If Text1(ptxORDER_CODE).Text = Trim(Right(Combo1(pcmbORDER).List(i), 5)) Then
                    Combo1(pcmbORDER).ListIndex = i
                    Exit For
                End If
            
            Next i
        
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function
Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
    
    Select Case Index
        Case pcmbORDER           '注文先
            Text1(ptxORDER_CODE).Text = Trim(Right(Combo1(pcmbORDER).Text, 5))
    End Select
    
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    
    Select Case Index
        Case pcmbORDER           '注文先
            Text1(ptxORDER_CODE).Text = Trim(Right(Combo1(pcmbORDER).Text, 5))
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans         As Integer
Dim i           As Integer


Dim sts         As Integer

Dim rpt         As New PR00020F1
Dim f           As New PR000202


    Select Case Index
        Case P_CMD_Upd          '更新
        
        Case P_CMD_DEL          '削除
        
        Case P_CMD_DSP                      '検索/表示
        
            For i = ptxS_ORDER_DT To ptxORDER_CODE
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
            
            
            If Z_List_Disp_Proc() Then
                Unload Me
            End If
        
            Text1(ptxS_ORDER_DT).SetFocus
        
        
        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        
            For i = ptxS_ORDER_DT To ptxORDER_CODE
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
                
            ans = MsgBox("データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If CSV_Data_Out_Proc() Then
                    Unload Me
                End If
            End If
        
            Text1(ptxS_ORDER_DT).SetFocus
        
        
        Case P_CMD_PRT                      '印刷
            
            For i = ptxS_ORDER_DT To ptxORDER_CODE
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
                
            ans = MsgBox("印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                '注文残一覧表
                
                
                If Data_Make_Proc() Then    '2007.10.31
                    Unload Me
                End If
                
                Set rpt = New PR00020F1
            
                'レポートを印刷します。（true：印刷ダイアログあり false：なし）
                rpt.PrintReport False
            
                Set rpt = Nothing
                
                
'                f.RunReport rpt
'                f.Show
            End If
            
            Text1(ptxS_ORDER_DT).SetFocus
 
            
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

Dim c       As String * 128
Dim sts     As Integer
Dim i       As Integer

    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                
    '注文残データファイル名獲得   2007.11.09
    If GetIni("FILE", "P_SHORDER_CSV", "SYS", c) Then
        Command1(P_CMD_OUT).Enabled = False
    Else
        Command1(P_CMD_OUT).Enabled = True
        P_SHORDER_CSV = Trim(c)
    End If
                                
                                
                                '大阪ＰＣ対応版？　2008.01.10
    If GetIni(App.EXEName, "OSAKA_MODE", "P_SYS", c) Then
        OSAKA_MODE = False
    Else
        
        If Not IsNumeric(Trim(c)) Then
            OSAKA_MODE = False
        Else
                
            If Trim(c) = "1" Then
                OSAKA_MODE = True
            Else
                OSAKA_MODE = False
            End If
        End If
    End If
                                
                                
                                
                                '対象収支の獲得 2008.10.09
    If Trim(GLB_SYUSHI_F) = "" Then
    
    Else
    
        If GetIni(StrConv(App.EXEName, vbUpperCase), GLB_SYUSHI_F, "P_SYS", c) Then
            Beep
            MsgBox "対象収支の獲得に失敗しました。処理を中止して下さい。"
            End
        End If
    
        G_SYUSHI_TBL = Split(Trim(c), ",", -1)
    End If
                                
                                
                                
                                
                                
                                
                                
                                
    Label1(plblY_NOUKI1).Visible = OSAKA_MODE
    Label1(plblY_NOUKI2).Visible = OSAKA_MODE
    
    Text1(ptxS_Y_NOUKI_DT).Visible = OSAKA_MODE
    Text1(ptxS_Y_NOUKI_DT).TabStop = OSAKA_MODE
                                    
    Text1(ptxE_Y_NOUKI_DT).Visible = OSAKA_MODE
    Text1(ptxE_Y_NOUKI_DT).TabStop = OSAKA_MODE
                                    
                                    
                                    
    Label1(plblUSE_YM).Visible = OSAKA_MODE
    Text1(plblUSE_YM).Visible = OSAKA_MODE
    Text1(plblUSE_YM).TabStop = OSAKA_MODE
                                    
    TDBGrid2.Columns(colZ_ANS_NOUKI_DT).Visible = OSAKA_MODE
    TDBGrid2.Columns(colZ_USE_YM).Visible = OSAKA_MODE
                                
                                
                                
                                
                                
                                
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '管理マスタＯＰＥＮ
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '受払先マスタＯＰＥＮ
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '資材注文ﾃﾞｰﾀＯＰＥＮ
    If P_SHORDER_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫ﾃﾞｰﾀＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '資材注文ﾃﾞｰﾀ(一時ﾃﾞｰﾀ)ＯＰＥＮ 2007.10.31
    If tmpP_SHORDER_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    Load PR000202
        
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
        
        
        
    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc
    
    
    
    '得意先
    If Ukeharai_Set_Proc(pcmbORDER) Then
        Unload Me
    End If
    
    '画面初期設定
    If Init_Proc() Then
        Unload Me
    End If

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
                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "管理マスタ")
        End If
    End If
                                            '受払先マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "受払先マスタ")
        End If
    End If
                                            '資材注文ﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材注文ﾃﾞｰﾀ")
        End If
    End If
                                            '在庫ﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫ﾃﾞｰﾀ")
        End If
    End If
    
    
                                            '資材注文ﾃﾞｰﾀＣＬＯＳＥ 2007.10.31
    sts = BTRV(BtOpClose, tmpP_SHORDER_POS, tmpP_SHORDER_REC, Len(tmpP_SHORDER_REC), K0_tmpP_SHORDER, Len(K0_tmpP_SHORDER), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材注文ﾃﾞｰﾀ")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set PR000201 = Nothing
    Set PR000202 = Nothing


    End
End Sub







Private Sub TDBGrid2_HeadClick(ByVal ColIndex As Integer)
    If Z_Sort_Tbl(ColIndex) = 0 Then
        Z_Sort_Tbl(ColIndex) = 1
    Else
        If Z_Sort_Tbl(ColIndex) = 1 Then
            Z_Sort_Tbl(ColIndex) = 0
        End If
    
    End If
    
    If Z_Sort_Tbl(ColIndex) = 0 Or Z_Sort_Tbl(ColIndex) = 1 Then
                    
        Z_SHORDER.QuickSort Z_Min_Row, Z_SHORDER.UpperBound(1), ColIndex, Z_Sort_Tbl(ColIndex), XTYPE_STRING
        
        Set TDBGrid2.Array = Z_SHORDER
        
        TDBGrid2.ReBind
        TDBGrid2.Update
        TDBGrid2.MoveFirst


    End If

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
        
        
    If Error_Check_Proc(Index) Then    'エラーチェック
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動
End Sub
Private Function Init_Proc() As Integer
'----------------------------------------------------------------------------
'                   入力画面の初期設定
'----------------------------------------------------------------------------
Dim i       As Integer
Dim sts     As Integer


    Init_Proc = True
    
    
    
    For i = ptxS_ORDER_DT To ptxORDER_CODE
        Text1(i).Text = ""
    Next i

    For i = pcmbORDER To pcmbORDER
        
        Combo1(i).ListIndex = -1
    
    Next i
    'ｿｰﾄ情報の初期化
    
    'ｿｰﾄ情報の初期化
    For i = 0 To UBound(Z_Sort_Tbl)
        Z_Sort_Tbl(i) = 0                 'ﾃﾞﾌｫﾙﾄ昇順
    Next i
    Z_Sort_Tbl(colZ_HIN_NAME) = 9           'ｿｰﾄ除外

    Init_Proc = False

End Function
Private Function Ukeharai_Set_Proc(Index As Integer) As Integer
'----------------------------------------------------------------------------
'                   受払先マスタをコンボにセットする。
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer




Dim i           As Integer
    
    Ukeharai_Set_Proc = True
    
    Combo1(Index).Clear
    
    
    Combo1(Index).AddItem Space(5)

    
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

        
        
        Combo1(Index).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode) & "            " & _
                                StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        
        com = BtOpGetNext
    
    Loop

    Ukeharai_Set_Proc = False
    



End Function

Private Function Z_List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'           資材注文残の表示    2007.07.27
'----------------------------------------------------------------------------
Dim sts                 As Integer
Dim com                 As Integer

Dim Row                 As Long

Dim SKIP_Flg            As Boolean

Dim i                   As Integer


Dim Mode                As Integer  '2007.12.05

    Z_List_Disp_Proc = True
    
    PR000201.MousePointer = vbHourglass
    
    'ｿｰﾄ情報の初期化
    For i = 0 To UBound(Z_Sort_Tbl)
        Z_Sort_Tbl(i) = 0           'ﾃﾞﾌｫﾙﾄ昇順
    Next i

    Z_Sort_Tbl(colZ_HIN_NAME) = 9   'ｿｰﾄ除外
    
    
    
    
    Set Z_SHORDER = Nothing
    
    Row = Z_Min_Row - 1
       
    
    
    '2007.12.05
    If (Trim(Text1(ptxS_Y_NOUKI_DT).Text) <> "" Or Trim(Text1(ptxE_Y_NOUKI_DT).Text) <> "") Then
        Mode = 1
    Else
        Mode = 0
    End If
    '2007.12.05
    
    
    
    Select Case Mode    '2007.12.05
        Case 0  '注文日指定
    
            Call UniCode_Conv(K3_P_SHORDER.KAN_F, P_KAN_OFF)
            
            If Len(Trim(Text1(ptxS_ORDER_DT).Text)) >= 10 Then
            
                Call UniCode_Conv(K3_P_SHORDER.ORDER_DT, Mid(Text1(ptxS_ORDER_DT).Text, 1, 4) & _
                                                                            Mid(Text1(ptxS_ORDER_DT).Text, 6, 2) & _
                                                                            Mid(Text1(ptxS_ORDER_DT).Text, 9, 2))
            
            
            Else
            
                Call UniCode_Conv(K3_P_SHORDER.ORDER_DT, "")
            End If
            
            Call UniCode_Conv(K3_P_SHORDER.ORDER_CODE, Text1(ptxORDER_CODE).Text)
    
        Case 1  '希望納期指定
    
            Call UniCode_Conv(K5_P_SHORDER.KAN_F, P_KAN_OFF)
            
            If Len(Trim(Text1(ptxS_Y_NOUKI_DT).Text)) >= 10 Then
            
                Call UniCode_Conv(K5_P_SHORDER.Y_NOUKI_DT, Mid(Text1(ptxS_Y_NOUKI_DT).Text, 1, 4) & _
                                                                            Mid(Text1(ptxS_Y_NOUKI_DT).Text, 6, 2) & _
                                                                            Mid(Text1(ptxS_Y_NOUKI_DT).Text, 9, 2))
            
            
            Else
            
                Call UniCode_Conv(K5_P_SHORDER.Y_NOUKI_DT, "")
            End If
            
            Call UniCode_Conv(K5_P_SHORDER.ORDER_CODE, Text1(ptxORDER_CODE).Text)
    
    
    
    
    End Select  '2007.12.05
    
    
    
    com = BtOpGetGreaterEqual
    
       
       
    Do
    
        DoEvents
    
        Select Case Mode    '2007.12.05
    
            Case 0
    
               sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K3_P_SHORDER, Len(K3_P_SHORDER), 3)
            
            Case 1
            
               sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K5_P_SHORDER, Len(K5_P_SHORDER), 5)
        
        
        End Select          '2007.12.05
        
        Select Case sts
            Case BtNoErr
                
            
                If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) <> P_KAN_OFF Then
                    Exit Do
                End If
            
                Select Case Mode    '2007.12.05
            
                    Case 0
            
            
                        If Trim(Text1(ptxE_ORDER_DT).Text) <> "" Then
                            If StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode) > (Mid(Text1(ptxE_ORDER_DT).Text, 1, 4) & _
                                                                        Mid(Text1(ptxE_ORDER_DT).Text, 6, 2) & _
                                                                        Mid(Text1(ptxE_ORDER_DT).Text, 9, 2)) Then
                    
                                Exit Do
                            End If
                        End If
            
            
                    Case 1
                
                
                        If Trim(Text1(ptxE_Y_NOUKI_DT).Text) <> "" Then
                            If StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode) > (Mid(Text1(ptxE_Y_NOUKI_DT).Text, 1, 4) & _
                                                                        Mid(Text1(ptxE_Y_NOUKI_DT).Text, 6, 2) & _
                                                                        Mid(Text1(ptxE_Y_NOUKI_DT).Text, 9, 2)) Then
                    
                                Exit Do
                            End If
                        End If
                
                
                End Select          '2007.12.05
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "資材注文ﾃﾞｰﾀ")
                Exit Function
        End Select
    
    
        SKIP_Flg = False
    
        
        If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
            SKIP_Flg = True
        End If
        
        
        
        If Trim(Text1(ptxUSE_YM).Text) <> "" Then   '2007.12.05
            If Left(Format(CDate(Text1(ptxUSE_YM).Text & "/01"), "YYYYMMDD"), 6) <> StrConv(P_SHORDER_REC.USE_YM, vbUnicode) Then
                SKIP_Flg = True
            End If
        End If
        
        If Trim(Text1(ptxORDER_CODE).Text) <> "" Then
        
            If Trim(Text1(ptxORDER_CODE).Text) <> Trim(StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode)) Then
                SKIP_Flg = True
            End If
        
        End If
        
        
        
        
        If GLB_SYUSHI_F = "" Then       '2008.10.09
        Else
            
            For i = 0 To UBound(G_SYUSHI_TBL)
            
                If Trim(StrConv(P_SHORDER_REC.G_SYUSHI, vbUnicode)) = G_SYUSHI_TBL(i) Then
                    Exit For
                End If
            
            
            Next i
        
        
            If i > UBound(G_SYUSHI_TBL) Then
                SKIP_Flg = True
            End If
        End If

    
        
        
        
        
        If Not SKIP_Flg Then
    
            Row = Row + 1
            If Z_Grid_Set_Proc(Row) Then
                Exit Function
            End If
        
        
        
        End If
        
        com = BtOpGetNext
    
    Loop
    
    
    
    Set TDBGrid2.Array = Z_SHORDER
    TDBGrid2.ReBind
    TDBGrid2.Update
    TDBGrid2.MoveFirst
    
    PR000201.MousePointer = vbDefault
    Z_List_Disp_Proc = False
    


End Function


Private Function Z_Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           資材注文残の内容をｸﾞﾘｯﾄﾞにｾｯﾄする   2007.07.27
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim i           As Integer

Dim Mi_QTY      As Long
Dim Sumi_QTY    As Long

    Z_Grid_Set_Proc = True
    
    Z_SHORDER.ReDim Z_Min_Row, Row, Z_Min_Col, Z_Max_Col
    
    Z_SHORDER(Row, colZ_ORDER_DT) = Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 7, 2)
    
    
    '注文№
    Z_SHORDER(Row, colZ_ORDER_NO) = Trim(StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode))
    '注文名
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
            Exit Function
    End Select
    Z_SHORDER(Row, colZ_ORDER_NAME) = Trim(StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode)) & " " & StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode)
    '資材品番
    Z_SHORDER(Row, colZ_HIN_GAI) = StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)
    '品名
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    End Select
    Z_SHORDER(Row, colZ_HIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    '手配数
    Z_SHORDER(Row, colZ_ORDER_QTY) = Format(CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)), "#,##0")
    '注文残
    Z_SHORDER(Row, colZ_ZAN_QTY) = Format(CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) - CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)), "#,##0")
    '現在庫
    If Zaiko_Syukei_Proc(Sumi_QTY, Mi_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                            StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
        Exit Function
    End If
    Z_SHORDER(Row, colZ_ZAIKO_QTY) = Format(Mi_QTY + Sumi_QTY, "#,##0")
    '納期予定日
    Z_SHORDER(Row, colZ_Y_NOUKI_DT) = Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 7, 2)
    
    
    '回答納期日 2007.12.05
    Z_SHORDER(Row, colZ_ANS_NOUKI_DT) = Mid(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 7, 2)
    
    '使用月 2007.12.05
    Z_SHORDER(Row, colZ_USE_YM) = Mid(StrConv(P_SHORDER_REC.USE_YM, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SHORDER_REC.USE_YM, vbUnicode), 5, 2)
    
    
    Z_Grid_Set_Proc = False

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
Private Function Data_Make_Proc() As Integer
'----------------------------------------------------------------------------
'                   資材注文残ﾃﾞｰﾀ作成  2007.10.31
'----------------------------------------------------------------------------
Dim sts                     As Integer
Dim com                     As Integer

Dim upd_com                 As Integer


    
Dim SKIP_Flg                As Boolean
    
Dim Mode                    As Integer  '2007.12.05
    
    
Dim i                       As Integer  '2008.10.09
    
    
    Data_Make_Proc = True
    PR000201.MousePointer = vbHourglass

    com = BtOpGetFirst

    Do
    
    
        sts = BTRV(com, tmpP_SHORDER_POS, tmpP_SHORDER_REC, Len(tmpP_SHORDER_REC), K0_tmpP_SHORDER, Len(K0_tmpP_SHORDER), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
            
                Exit Do
            
            
            Case Else
                Call File_Error(sts, com, "資材注文残ﾃﾞｰﾀ")
                Exit Function
        End Select

        sts = BTRV(BtOpDelete, tmpP_SHORDER_POS, tmpP_SHORDER_REC, Len(tmpP_SHORDER_REC), K0_tmpP_SHORDER, Len(K0_tmpP_SHORDER), 0)
            
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "資材注文残ﾃﾞｰﾀ")
        End Select

    
        com = BtOpGetNext
    
    Loop
    
                                                                    
                                                                    
                                                                    
    '2008.01.10
    If (Trim(Text1(ptxE_Y_NOUKI_DT).Text) <> "" Or Trim(Text1(ptxE_Y_NOUKI_DT).Text) <> "") Then
        Mode = 1
    Else
        Mode = 0
    End If
    '2008.01.10
                                                                    
                                                                    
                                                                    
    Select Case Mode        '2008.01.10
        Case 0
                                                                    
                                                                    '仕入先
            Call UniCode_Conv(K3_P_SHORDER.KAN_F, P_KAN_OFF)
            
            If Len(Trim(Text1(ptxS_ORDER_DT).Text)) >= 10 Then
            
                Call UniCode_Conv(K3_P_SHORDER.ORDER_DT, Mid(Text1(ptxS_ORDER_DT).Text, 1, 4) & _
                                                                            Mid(Text1(ptxS_ORDER_DT).Text, 6, 2) & _
                                                                            Mid(Text1(ptxS_ORDER_DT).Text, 9, 2))
            
            
            Else
            
                Call UniCode_Conv(K3_P_SHORDER.ORDER_DT, "")
            End If
                
            
            Call UniCode_Conv(K3_P_SHORDER.ORDER_CODE, Text1(ptxORDER_CODE).Text)
    
        Case 1
    
            Call UniCode_Conv(K5_P_SHORDER.KAN_F, P_KAN_OFF)
    
    
            If Len(Trim(Text1(ptxS_Y_NOUKI_DT).Text)) >= 10 Then
            
                Call UniCode_Conv(K5_P_SHORDER.Y_NOUKI_DT, Mid(Text1(ptxS_Y_NOUKI_DT).Text, 1, 4) & _
                                                                            Mid(Text1(ptxS_Y_NOUKI_DT).Text, 6, 2) & _
                                                                            Mid(Text1(ptxS_Y_NOUKI_DT).Text, 9, 2))
            
            
            Else
            
                Call UniCode_Conv(K5_P_SHORDER.Y_NOUKI_DT, "")
            End If
                
            
            Call UniCode_Conv(K5_P_SHORDER.ORDER_CODE, Text1(ptxORDER_CODE).Text)
    
    
    End Select              '2008.01.10
    
    
    
    com = BtOpGetGreaterEqual
    
       
       
    Do
    
        DoEvents
    
        Select Case Mode    '2008.01.10
    
            Case 0
    
               sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K3_P_SHORDER, Len(K3_P_SHORDER), 3)
            
            Case 1
            
               sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K5_P_SHORDER, Len(K5_P_SHORDER), 5)
        
        
        End Select          '2008.01.10
            
        Select Case sts
            Case BtNoErr
                
            
                If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) <> P_KAN_OFF Then
                    Exit Do
                End If
            
                Select Case Mode    '2008.01.10
                    
                    Case 0
                
                        If Trim(Text1(ptxE_ORDER_DT).Text) <> "" Then
                            If StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode) > (Mid(Text1(ptxE_ORDER_DT).Text, 1, 4) & _
                                                                        Mid(Text1(ptxE_ORDER_DT).Text, 6, 2) & _
                                                                        Mid(Text1(ptxE_ORDER_DT).Text, 9, 2)) Then
                    
                                Exit Do
                            
                            End If
                        End If
            
                    Case 1
                
                        If Trim(Text1(ptxE_Y_NOUKI_DT).Text) <> "" Then
                            If StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode) > (Mid(Text1(ptxE_Y_NOUKI_DT).Text, 1, 4) & _
                                                                        Mid(Text1(ptxE_Y_NOUKI_DT).Text, 6, 2) & _
                                                                        Mid(Text1(ptxE_Y_NOUKI_DT).Text, 9, 2)) Then
                    
                                Exit Do
                            
                            End If
                        End If
                
                
                End Select          '2008.01.10
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "資材注文ﾃﾞｰﾀ")
                Exit Function
        End Select
    
    
        SKIP_Flg = False
    
        
        If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
            SKIP_Flg = True
        End If
        
        
        
        If Trim(Text1(ptxUSE_YM).Text) <> "" Then   '2008.01.10
            If Left(Format(CDate(Text1(ptxUSE_YM).Text & "/01"), "YYYYMMDD"), 6) <> StrConv(P_SHORDER_REC.USE_YM, vbUnicode) Then
                SKIP_Flg = True
            End If
        End If
        
        
        
        If Trim(Text1(ptxORDER_CODE).Text) <> "" Then
        
            If Trim(Text1(ptxORDER_CODE).Text) <> Trim(StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode)) Then
                SKIP_Flg = True
            End If
        
        End If
        
        
        
        If GLB_SYUSHI_F = "" Then       '2008.10.09
        Else
            
            For i = 0 To UBound(G_SYUSHI_TBL)
            
                If Trim(StrConv(P_SHORDER_REC.G_SYUSHI, vbUnicode)) = G_SYUSHI_TBL(i) Then
                    Exit For
                End If
            
            
            Next i
        
        
            If i > UBound(G_SYUSHI_TBL) Then
                SKIP_Flg = True
            End If
        End If
        
        If Not SKIP_Flg Then
    
        
           sts = BTRV(BtOpInsert, tmpP_SHORDER_POS, P_SHORDER_REC, Len(tmpP_SHORDER_REC), K0_tmpP_SHORDER, Len(K0_tmpP_SHORDER), 0)
                           
            Select Case sts
                Case BtNoErr
                Case Else
                    Call File_Error(sts, com, "資材注文ﾃﾞｰﾀ")
                    Exit Function
            End Select
        
        
        
        
        End If
        
        com = BtOpGetNext
    
    Loop

    PR000201.MousePointer = vbDefault

   Data_Make_Proc = False

End Function


Private Function CSV_Data_Out_Proc() As Integer
'----------------------------------------------------------------------------
'           資材注文残のCSV出力    2007.11.09
'----------------------------------------------------------------------------
Dim FileNo          As Integer
Dim fileName        As String
Dim Ret             As Integer



Dim sts             As Integer
Dim com             As Integer

Dim Sumi_QTY        As Long
Dim Mi_QTY          As Long


Dim SKIP_Flg        As Boolean

Dim i               As Integer

Dim Mode            As Integer      '2008.01.10




    CSV_Data_Out_Proc = True
    
    PR000201.MousePointer = vbHourglass
    Call Input_Lock
    
    FileNo = FreeFile
    fileName = P_SHORDER_CSV
    
'    Ret = InStr(1, Trim(fileName), ".") - 1
    
    
    Ret = InStrRev(Trim(fileName), ".") - 1
    
    fileName = Left(Trim(fileName), Ret) & Right(Trim(fileName), Len(Trim(fileName)) - Ret)

    On Error GoTo Error_Proc

    Open (fileName) For Output As FileNo
    
'    Write #FileNo, "注文日", "注文№", "注文先名", "資材品番", "品名", "注文数", "注文残", "在庫残", "予定納期"                            2008.01.10
    Write #FileNo, "注文日", "注文№", "注文先名", "資材品番", "品名", "注文数", "注文残", "在庫残", "希望納期日", "回答納期日", "使用月"   '2008.01.10
       
    
    '2008.01.10
    If (Trim(Text1(ptxS_Y_NOUKI_DT).Text) <> "" Or Trim(Text1(ptxE_Y_NOUKI_DT).Text) <> "") Then
        Mode = 1
    Else
        Mode = 0
    End If
    '2008.01.10
    
    
    
    Select Case Mode        '2008.01.10
        Case 0
                                                                    
                                                                    '仕入先
            Call UniCode_Conv(K3_P_SHORDER.KAN_F, P_KAN_OFF)
            
            If Len(Trim(Text1(ptxS_ORDER_DT).Text)) >= 10 Then
            
                Call UniCode_Conv(K3_P_SHORDER.ORDER_DT, Mid(Text1(ptxS_ORDER_DT).Text, 1, 4) & _
                                                                            Mid(Text1(ptxS_ORDER_DT).Text, 6, 2) & _
                                                                            Mid(Text1(ptxS_ORDER_DT).Text, 9, 2))
            
            
            Else
            
                Call UniCode_Conv(K3_P_SHORDER.ORDER_DT, "")
            End If
                
            
            Call UniCode_Conv(K3_P_SHORDER.ORDER_CODE, Text1(ptxORDER_CODE).Text)
    
        Case 1
    
            Call UniCode_Conv(K5_P_SHORDER.KAN_F, P_KAN_OFF)
    
    
            If Len(Trim(Text1(ptxS_Y_NOUKI_DT).Text)) >= 10 Then
            
                Call UniCode_Conv(K5_P_SHORDER.Y_NOUKI_DT, Mid(Text1(ptxS_Y_NOUKI_DT).Text, 1, 4) & _
                                                                            Mid(Text1(ptxS_Y_NOUKI_DT).Text, 6, 2) & _
                                                                            Mid(Text1(ptxS_Y_NOUKI_DT).Text, 9, 2))
            
            
            Else
            
                Call UniCode_Conv(K5_P_SHORDER.Y_NOUKI_DT, "")
            End If
                
            
            Call UniCode_Conv(K5_P_SHORDER.ORDER_CODE, Text1(ptxORDER_CODE).Text)
    
    
    End Select              '2008.01.10
    
    
    com = BtOpGetGreaterEqual
    
       
       
    Do
    
        DoEvents
    
        Select Case Mode    '2008.01.10
    
            Case 0
    
               sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K3_P_SHORDER, Len(K3_P_SHORDER), 3)
            
            Case 1
            
               sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K5_P_SHORDER, Len(K5_P_SHORDER), 5)
        
        
        End Select          '2008.01.10
            
        Select Case sts
            Case BtNoErr
                
            
                If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) <> P_KAN_OFF Then
                    Exit Do
                End If
            
                Select Case Mode    '2008.01.10
                    
                    Case 0
                
                        If Trim(Text1(ptxE_ORDER_DT).Text) <> "" Then
                            If StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode) > (Mid(Text1(ptxE_ORDER_DT).Text, 1, 4) & _
                                                                        Mid(Text1(ptxE_ORDER_DT).Text, 6, 2) & _
                                                                        Mid(Text1(ptxE_ORDER_DT).Text, 9, 2)) Then
                    
                                Exit Do
                            
                            End If
                        End If
            
                    Case 1
                
                        If Trim(Text1(ptxE_Y_NOUKI_DT).Text) <> "" Then
                            If StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode) > (Mid(Text1(ptxE_Y_NOUKI_DT).Text, 1, 4) & _
                                                                        Mid(Text1(ptxE_Y_NOUKI_DT).Text, 6, 2) & _
                                                                        Mid(Text1(ptxE_Y_NOUKI_DT).Text, 9, 2)) Then
                    
                                Exit Do
                            
                            End If
                        End If
                
                
                End Select          '2008.01.10
            
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "資材注文ﾃﾞｰﾀ")
                Exit Function
        End Select
    
    
        SKIP_Flg = False
    
        
        If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
            SKIP_Flg = True
        End If
        
        
        
        If Trim(Text1(ptxUSE_YM).Text) <> "" Then   '2008.01.10
            If Left(Format(CDate(Text1(ptxUSE_YM).Text & "/01"), "YYYYMMDD"), 6) <> StrConv(P_SHORDER_REC.USE_YM, vbUnicode) Then
                SKIP_Flg = True
            End If
        End If
        
        
        If Trim(Text1(ptxORDER_CODE).Text) <> "" Then
        
            If Trim(Text1(ptxORDER_CODE).Text) <> Trim(StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode)) Then
                SKIP_Flg = True
            End If
        
        End If
        
        
        If Not SKIP_Flg Then
    
        
            '注文日
            Write #FileNo, Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 7, 2),
        
            '注文№
            Write #FileNo, Trim(StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode)),
            '注文先名
            Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
                    Exit Function
            End Select
            Write #FileNo, Trim(StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode)) & " " & StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode),
            '資材品番
            Write #FileNo, Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)),
            '品名
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
            Write #FileNo, StrConv(ITEMREC.HIN_NAME, vbUnicode),
            '注文数
            Write #FileNo, Format(CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)), "#,##0"),
            '注文残
            Write #FileNo, Format(CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) - CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)), "#,##0"),
            '現在庫
            If Zaiko_Syukei_Proc(Sumi_QTY, Mi_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                            StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                Exit Function
            End If
            Write #FileNo, Format(Mi_QTY + Sumi_QTY, "#,##0"),
            '納期予定日
            Write #FileNo, Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 7, 2),
            '回答納期日 2008.01.10
            If Trim(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode)) <> "" Then
                Write #FileNo, Mid(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 7, 2),
            Else
                Write #FileNo, ,
            End If
        
            '使用月 2008.01.10
            If Trim(StrConv(P_SHORDER_REC.USE_YM, vbUnicode)) <> "" Then
                Write #FileNo, Mid(StrConv(P_SHORDER_REC.USE_YM, vbUnicode), 1, 4) & "年" & _
                                        Mid(StrConv(P_SHORDER_REC.USE_YM, vbUnicode), 5, 2) & "月"
            Else
                Write #FileNo,
            End If
        
        
        End If
        
        com = BtOpGetNext
    
    Loop
    
    
    
    
    PR000201.MousePointer = vbDefault
    CSV_Data_Out_Proc = False
    
    Close #FileNo
    
    Call Input_UnLock
    
    Beep
    MsgBox "「" & fileName & "」は正常に出力されました。"


    Exit Function
Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "が使用中です。"
        CSV_Data_Out_Proc = False
    Else
        MsgBox "Err.Number= " & Err.Number
        CSV_Data_Out_Proc = True
    End If

    Call Input_UnLock


End Function


