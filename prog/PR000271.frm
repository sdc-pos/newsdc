VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PR000271 
   Caption         =   "仕入実績検索・仕入集計表発行(PR00027 2016.02.26 09:15)"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17100
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
   ScaleWidth      =   17100
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   2
      Left            =   12960
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   4320
      MaxLength       =   5
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "出力対象"
      Height          =   855
      Left            =   8640
      TabIndex        =   21
      Top             =   120
      Width           =   5295
      Begin VB.CheckBox Check1 
         Caption         =   "仕入明細表"
         Height          =   375
         Index           =   1
         Left            =   3240
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "仕入先別仕入集計表"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   5040
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   7455
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   16245
      _ExtentX        =   28654
      _ExtentY        =   13150
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "注文��"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "受入月日"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "仕入先"
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
      Columns(5).Caption=   "仕入区分"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "収支単位"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "数量"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "単価"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "金額"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "消費税額"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1826"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1720"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2090"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1984"
      Splits(0)._ColumnProps(8)=   "Column(1)._ColStyle=0"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=3493"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=3387"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=2064"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=1958"
      Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=0"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=4075"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=3969"
      Splits(0)._ColumnProps(22)=   "Column(4)._ColStyle=0"
      Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(24)=   "Column(5).Width=1879"
      Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=1773"
      Splits(0)._ColumnProps(27)=   "Column(5)._ColStyle=0"
      Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(29)=   "Column(6).Width=1879"
      Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=1773"
      Splits(0)._ColumnProps(32)=   "Column(6)._ColStyle=0"
      Splits(0)._ColumnProps(33)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(34)=   "Column(7).Width=2011"
      Splits(0)._ColumnProps(35)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(7)._WidthInPix=1905"
      Splits(0)._ColumnProps(37)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(38)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(39)=   "Column(8).Width=2514"
      Splits(0)._ColumnProps(40)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(8)._WidthInPix=2408"
      Splits(0)._ColumnProps(42)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(43)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(44)=   "Column(9).Width=2699"
      Splits(0)._ColumnProps(45)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(9)._WidthInPix=2593"
      Splits(0)._ColumnProps(47)=   "Column(9)._ColStyle=2"
      Splits(0)._ColumnProps(48)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(49)=   "Column(10).Width=2699"
      Splits(0)._ColumnProps(50)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(10)._WidthInPix=2593"
      Splits(0)._ColumnProps(52)=   "Column(10)._ColStyle=2"
      Splits(0)._ColumnProps(53)=   "Column(10).Order=11"
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
      Caption         =   "仕入明細表"
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
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=62,.parent=43"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=44"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=45"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=47"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=58,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(43)  =   ":id=58,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(44)  =   ":id=58,.fontname=ＭＳ ゴシック"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=44"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=45"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=47"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=16,.parent=43"
      _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=13,.parent=44"
      _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=14,.parent=45"
      _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=15,.parent=47"
      _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=28,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(53)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(54)  =   ":id=28,.fontname=ＭＳ ゴシック"
      _StyleDefs(55)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=44"
      _StyleDefs(56)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=45"
      _StyleDefs(57)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=47"
      _StyleDefs(58)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(59)  =   ":id=66,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(60)  =   ":id=66,.fontname=ＭＳ ゴシック"
      _StyleDefs(61)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44"
      _StyleDefs(62)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
      _StyleDefs(63)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
      _StyleDefs(64)  =   "Splits(0).Columns(5).Style:id=32,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(65)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(66)  =   ":id=32,.fontname=ＭＳ ゴシック"
      _StyleDefs(67)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=44"
      _StyleDefs(68)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=45"
      _StyleDefs(69)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=47"
      _StyleDefs(70)  =   "Splits(0).Columns(6).Style:id=74,.parent=43,.alignment=0"
      _StyleDefs(71)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=44"
      _StyleDefs(72)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=45"
      _StyleDefs(73)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=47"
      _StyleDefs(74)  =   "Splits(0).Columns(7).Style:id=20,.parent=43,.alignment=1"
      _StyleDefs(75)  =   "Splits(0).Columns(7).HeadingStyle:id=17,.parent=44"
      _StyleDefs(76)  =   "Splits(0).Columns(7).FooterStyle:id=18,.parent=45"
      _StyleDefs(77)  =   "Splits(0).Columns(7).EditorStyle:id=19,.parent=47"
      _StyleDefs(78)  =   "Splits(0).Columns(8).Style:id=24,.parent=43,.alignment=1"
      _StyleDefs(79)  =   "Splits(0).Columns(8).HeadingStyle:id=21,.parent=44"
      _StyleDefs(80)  =   "Splits(0).Columns(8).FooterStyle:id=22,.parent=45"
      _StyleDefs(81)  =   "Splits(0).Columns(8).EditorStyle:id=23,.parent=47"
      _StyleDefs(82)  =   "Splits(0).Columns(9).Style:id=78,.parent=43,.alignment=1"
      _StyleDefs(83)  =   "Splits(0).Columns(9).HeadingStyle:id=75,.parent=44"
      _StyleDefs(84)  =   "Splits(0).Columns(9).FooterStyle:id=76,.parent=45"
      _StyleDefs(85)  =   "Splits(0).Columns(9).EditorStyle:id=77,.parent=47"
      _StyleDefs(86)  =   "Splits(0).Columns(10).Style:id=70,.parent=43,.alignment=1"
      _StyleDefs(87)  =   "Splits(0).Columns(10).HeadingStyle:id=67,.parent=44"
      _StyleDefs(88)  =   "Splits(0).Columns(10).FooterStyle:id=68,.parent=45"
      _StyleDefs(89)  =   "Splits(0).Columns(10).EditorStyle:id=69,.parent=47"
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
      Caption         =   "ﾌﾟﾚﾋﾞｭｰ"
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
      Caption         =   "再計算"
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
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "合計金額"
      Height          =   255
      Index           =   0
      Left            =   11760
      TabIndex        =   22
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "対象年月度"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   20
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "仕入先"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   19
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "PR000271"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'テキスト用添字
Private Const ptxKEIJYO_YM% = 0             '対象年月
Private Const ptxSHIIRE_CODE% = 1           '仕入先ｺｰﾄﾞ
Private Const ptxTOTAL% = 2                 '合計金額


'コンボ用添字
Private Const pcmbSHIIRE% = 0               '仕入先

'チェックボックス用添字
Private Const pchkG_SHUKEIRE% = 0           '仕入先別仕入集計表
Private Const pchkD_SHUKEIRE% = 1           '仕入明細表

'Glid用環境---------------------------------

'仕入明細
Private Const pGridDETAIL% = 0


Private SHUKEIRE      As New XArrayDB


Private Const Min_Row% = 1                  '最小行数
Private Const Min_Col% = 0                  '最小列数
Private Const Max_Col% = 10                 '最大列数           '2007.08.01

Private Const colORDER_NO% = 0              '注文��             '2007.06.29
Private Const colUKEIRE_DT% = 1             '年月日（受入）
Private Const colSHIIRE% = 2                '仕入先
Private Const colHIN_GAI% = 3               '品番
Private Const colHIN_NAME% = 4              '品名
Private Const colSHIIRE_KBN% = 5            '販売区分
Private Const colSYUSHI% = 6                '収支
Private Const colUKEIRE_QTY% = 7             '数量
Private Const colUKEIRE_TANKA% = 8          '単価
Private Const colUKEIRE_KINGAKU% = 9        '金額
Private Const colZEI_KIN% = 10              '消費税額           '2007.08.01


Private Sort_Tbl(Min_Col To Max_Col) _
                As Integer                  'ｿｰﾄの制御 0:昇順 1:降順
Private Tbl_Set_F   As Boolean

Private P_SHUKEIRE_CSV As String            'データ出力用   2007.08.03


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PR000271.MousePointer = vbHourglass

    Call Ctrl_Lock(PR000271)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PR000271)


    PR000271.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts         As Integer
Dim com         As Integer
    
Dim wkdate      As String * 10

Dim i           As Integer
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        
        Case ptxKEIJYO_YM       '対象年月
        
            wkdate = Text1(ptxKEIJYO_YM).Text & "/01"
            
            If Not IsDate(wkdate) Then
                MsgBox "入力した項目はエラーです。"
                Text1(Mode).SetFocus
                Exit Function
            Else
                
                wkdate = Format(CDate(Text1(ptxKEIJYO_YM).Text), "YYYY/MM/DD")
                
                Text1(ptxKEIJYO_YM).Text = Mid(wkdate, 1, 7)
            End If
        
        Case ptxSHIIRE_CODE     '仕入先
            
            Text1(ptxSHIIRE_CODE).Text = StrConv(Text1(ptxSHIIRE_CODE).Text, vbUpperCase)       '2016.01.19
            
            Combo1(pcmbSHIIRE).ListIndex = -1
            For i = 0 To Combo1(pcmbSHIIRE).ListCount - 1
                If Text1(ptxSHIIRE_CODE).Text = Trim(Right(Combo1(pcmbSHIIRE).List(i), 5)) Then
                    Combo1(pcmbSHIIRE).ListIndex = i
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
        Case pcmbSHIIRE             '仕入先
            Text1(ptxSHIIRE_CODE).Text = Trim(Right(Combo1(pcmbSHIIRE).Text, 5))
    End Select
    
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    
    Select Case Index
        Case pcmbSHIIRE             '仕入先
            Text1(ptxSHIIRE_CODE).Text = Trim(Right(Combo1(pcmbSHIIRE).Text, 5))
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans         As Integer
Dim i           As Integer

Dim rpt             As New PR00027F2
Dim f               As New PR000272


Dim sts         As Integer

Dim yn          As Integer


    Select Case Index
        Case P_CMD_Upd          '更新
        
                    
            If Kingaku_Proc() Then
                Unload Me
            End If
        
        
        
        
        
        
        Case P_CMD_DEL          '削除
        
        Case P_CMD_DSP                      '検索/表示
        
            For i = ptxKEIJYO_YM To ptxSHIIRE_CODE
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
        
        
            If List_Disp_Proc() Then
                Unload Me
            End If
        
            Text1(ptxKEIJYO_YM).SetFocus
        
        
        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        
            Beep
            yn = MsgBox("データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                If Data_Proc() Then
                    Unload Me
                End If
            End If
        
        
        
        
        
        Case P_CMD_PRT                      '印刷
 
            For i = ptxKEIJYO_YM To ptxSHIIRE_CODE
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
                
            ans = MsgBox("印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                
                If Check1(pchkG_SHUKEIRE).Value = vbChecked Then
                    '仕入集計表
                    If G_Print_Proc(0) Then
                        Unload Me
                    End If
                End If
            
                If Check1(pchkD_SHUKEIRE).Value = vbChecked Then
                    '仕入明細表
                    Set rpt = New PR00027F2
                
                    'レポートを印刷します。（true：印刷ダイアログあり false：なし）
                    rpt.PrintReport False
                
                    Set rpt = Nothing
                    
                    
'                    f.RunReport rpt
'                    f.Show
                End If
            
            End If
            
            Text1(ptxKEIJYO_YM).SetFocus
            
        Case 9                          'ﾌﾟﾚﾋﾞｭｰ 2007.10.01
 
            For i = ptxKEIJYO_YM To ptxSHIIRE_CODE
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
                
            ans = MsgBox("ﾌﾟﾚﾋﾞｭｰ表示しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                
                If Check1(pchkG_SHUKEIRE).Value = vbChecked Then
                    '仕入集計表
                    If G_Print_Proc(1) Then
                        Unload Me
                    End If
                End If
            
                If Check1(pchkD_SHUKEIRE).Value = vbChecked Then
                    
                    
                    f.ARViewer1.Zoom = -2
                    
                    f.RunReport rpt
                    f.Show vbModal
                End If
            
            End If
            
            Text1(ptxKEIJYO_YM).SetFocus
            
            
            
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
                                '資材受入ﾃﾞｰﾀＯＰＥＮ
    If P_SHUKEIRE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '資材仕入集計ﾃﾞｰﾀＯＰＥＮ
    If P_SHSHI_SUM_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    Load PR000272
    
    
    
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
    
    '仕入実績データファイル名獲得   2007.08.03
    If GetIni("FILE", "P_SHUKEIRE_CSV", "SYS", c) Then
        Command1(P_CMD_OUT).Enabled = False
    Else
        Command1(P_CMD_OUT).Enabled = True
        P_SHUKEIRE_CSV = Trim(c)
    End If
    
    
    '仕入先
    If Ukeharai_Set_Proc(pcmbSHIIRE) Then
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
                                            '資材受入データCLOSE
    sts = BTRV(BtOpClose, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K0_P_SHUKEIRE, Len(K0_P_SHUKEIRE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材受入ﾃﾞｰﾀ")
        End If
    End If
                                            '資材仕入データCLOSE
    sts = BTRV(BtOpClose, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材受入ﾃﾞｰﾀ")
        End If
    End If
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PR000271 = Nothing
    Set PR000272 = Nothing


    End
End Sub





Private Sub TDBGrid1_HeadClick(Index As Integer, ByVal ColIndex As Integer)



    Select Case Index
        
        Case pGridDETAIL        '仕入明細
    
    
            If Sort_Tbl(ColIndex) = 0 Then
                Sort_Tbl(ColIndex) = 1
            Else
                If Sort_Tbl(ColIndex) = 1 Then
                    Sort_Tbl(ColIndex) = 0
                End If
            
            End If
            
            If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then
                            
                SHUKEIRE.QuickSort Min_Row, SHUKEIRE.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
                
                Set TDBGrid1(Index).Array = SHUKEIRE
                
                TDBGrid1(Index).ReBind
                TDBGrid1(Index).Update
                TDBGrid1(Index).MoveFirst
        
        
            End If
    
    
    
    End Select




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
    
    
    
    For i = ptxKEIJYO_YM To ptxTOTAL
        Text1(i).Text = ""
    Next i
    '処理年月＝当月
    Text1(ptxKEIJYO_YM).Text = Mid(Format(Now, "YYYY/MM/DD"), 1, 7)



    For i = pcmbSHIIRE To pcmbSHIIRE
        
        Combo1(i).ListIndex = -1
    
    Next i
    'ｿｰﾄ情報の初期化
    
    'ｿｰﾄ情報の初期化
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0               'ﾃﾞﾌｫﾙﾄ昇順
    Next i
    Sort_Tbl(colHIN_NAME) = 9       'ｿｰﾄ除外

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
Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'           資材受入データの表示
'----------------------------------------------------------------------------
Dim sts                 As Integer
Dim com                 As Integer

Dim Row                 As Long


Dim wKEIJYO_YM          As String * 6
Dim SKIP_Flg            As Boolean

Dim i                   As Integer

Dim TOTAL               As Long

    List_Disp_Proc = True
    PR000271.MousePointer = vbHourglass
    
    Set SHUKEIRE = Nothing
    
    Row = Min_Row - 1
       
    TOTAL = 0
    
    wKEIJYO_YM = Mid(Format(CDate(Text1(ptxKEIJYO_YM).Text & "/01"), "YYYYMMDD"), 1, 6)
    
    Call UniCode_Conv(K1_P_SHUKEIRE.KEIJYO_YM, wKEIJYO_YM)                      '計上年月
    Call UniCode_Conv(K1_P_SHUKEIRE.ORDER_CODE, Text1(ptxSHIIRE_CODE).Text)     '仕入先
    Call UniCode_Conv(K1_P_SHUKEIRE.UKEIRE_DT, "")                              '受入年月日
    
    
    com = BtOpGetGreaterEqual
    
       
       
    Do
    
        DoEvents
    
        sts = BTRV(com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K1_P_SHUKEIRE, Len(K1_P_SHUKEIRE), 1)
            
        Select Case sts
            Case BtNoErr
            
                '計上年月のﾌﾞﾚｰｸ
                If StrConv(P_SHUKEIRE_REC.KEIJYO_YM, vbUnicode) <> wKEIJYO_YM Then
                    Exit Do
                End If
            
                '仕入先のﾌﾞﾚｰｸ
                If Trim(Text1(ptxSHIIRE_CODE).Text) = "" Then
                Else
                    If Trim(Text1(ptxSHIIRE_CODE).Text) <> Trim(StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode)) Then
                        Exit Do
                    End If
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "資材受入ﾃﾞｰﾀ")
                Exit Function
        End Select
    
            '注文ﾃﾞｰﾀ読み込み
        SKIP_Flg = False
        Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                '異常データ
'                SKIP_Flg = True
            
                Call UniCode_Conv(P_SHORDER_REC.HIN_GAI, "***")
                Call UniCode_Conv(P_SHORDER_REC.G_SYUSHI, "***")
                Call UniCode_Conv(P_SHORDER_REC.G_SHIIRE_KBN, "**")
            
            
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "資材注文ﾃﾞｰﾀ")
                Exit Function
        End Select
    
        If Not SKIP_Flg Then
            Row = Row + 1
            If Grid_Set_Proc(Row) Then
                Exit Function
            
            End If
        
        
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN01_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SHORDER_REC.G_SHIIRE_KBN, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                
                    If Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)) <> P_SH_ZEI Then
                        TOTAL = TOTAL + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                    End If
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_CODEREC.C_RNAME, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "コードマスタ")
                    Exit Function
            End Select
        
        
        End If
        
        
        com = BtOpGetNext
    
    Loop
    
    
    
    Set TDBGrid1(pGridDETAIL).Array = SHUKEIRE
    TDBGrid1(pGridDETAIL).ReBind
    TDBGrid1(pGridDETAIL).Update
    TDBGrid1(pGridDETAIL).MoveFirst
    
    Text1(ptxTOTAL).Text = Format(TOTAL, "#,##0")
    
    PR000271.MousePointer = vbDefault
    List_Disp_Proc = False
    


End Function


Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           資材受入ﾃﾞｰﾀ（仕入明細表）の内容をｸﾞﾘｯﾄﾞにｾｯﾄする
'----------------------------------------------------------------------------
Dim sts As Integer
Dim i   As Integer


    Grid_Set_Proc = True
    
    SHUKEIRE.ReDim Min_Row, Row, Min_Col, Max_Col


    'ｵｰﾀﾞｰ��
    SHUKEIRE(Row, colORDER_NO) = StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode) & "-" & StrConv(P_SHUKEIRE_REC.SEQNO, vbUnicode)

    '月日
    SHUKEIRE(Row, colUKEIRE_DT) = Mid(StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode), 7, 2)

    '仕入先
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
            Exit Function
    End Select
    SHUKEIRE(Row, colSHIIRE) = StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode) & " " & StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode)
    
    '品番
    SHUKEIRE(Row, colHIN_GAI) = StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)
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
    SHUKEIRE(Row, colHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    
    
    '収支区分
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN03_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SHORDER_REC.G_SYUSHI, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "コードマスタ")
            Exit Function
    End Select
    SHUKEIRE(Row, colSYUSHI) = Trim(StrConv(P_SHORDER_REC.G_SYUSHI, vbUnicode)) & " " & _
                StrConv(P_CODEREC.C_RNAME, vbUnicode)
    
    
    '仕入区分
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN01_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SHORDER_REC.G_SHIIRE_KBN, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "コードマスタ")
            Exit Function
    End Select
    SHUKEIRE(Row, colSHIIRE_KBN) = Trim(StrConv(P_SHORDER_REC.G_SHIIRE_KBN, vbUnicode)) & " " & _
                StrConv(P_CODEREC.C_RNAME, vbUnicode)
    
    '数量
    SHUKEIRE(Row, colUKEIRE_QTY) = Format(CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode)), "#,##0.00")
    '単価
    SHUKEIRE(Row, colUKEIRE_TANKA) = Format(CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_TANKA, vbUnicode)), "#,##0.00")
    '金額
    
    If Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)) <> P_SH_ZEI Then
        SHUKEIRE(Row, colUKEIRE_KINGAKU) = Format(CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)), "#,##0")
    
        '消費税額   2007.08.01
        If IsNumeric(StrConv(P_SHUKEIRE_REC.ZEI_KIN, vbUnicode)) Then
            SHUKEIRE(Row, colZEI_KIN) = Format(CDbl(StrConv(P_SHUKEIRE_REC.ZEI_KIN, vbUnicode)), "#,##0")
        Else
            SHUKEIRE(Row, colZEI_KIN) = "0"
        End If
    
    Else
        SHUKEIRE(Row, colUKEIRE_KINGAKU) = 0
        If IsNumeric(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)) Then
            SHUKEIRE(Row, colZEI_KIN) = Format(CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)), "#,##0")
        Else
            SHUKEIRE(Row, colZEI_KIN) = "0"
        End If
    
    
    End If
    
    
    Grid_Set_Proc = False

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

Private Function SUM_Make_Proc(Data_Flg As Boolean) As Integer
'----------------------------------------------------------------------------
'                   資材仕入集計ﾃﾞｰﾀ作成
'----------------------------------------------------------------------------
Dim sts                     As Integer
Dim com                     As Integer

Dim upd_com                 As Integer

Dim wKEIJYO_YM              As String * 6

Dim i                       As Integer
    
    
Dim GENERAL_SUM(0 To 6)     As Long
Dim NAISYOKU_SUM(0 To 6)    As Long
Dim GENKIN_SUM(0 To 6)      As Long
Dim SYANAI_SUM(0 To 6)      As Long
Dim TACENTER_SUM(0 To 6)    As Long
    
Dim SKIP_Flg                As Boolean
    
Dim YMD                     As String * 8
Dim ZEI                     As Long
    
Dim wkKINGAKU               As Long
    
    
    SUM_Make_Proc = True
    PR000271.MousePointer = vbHourglass

    com = BtOpGetFirst

    Do
    
    
        sts = BTRV(com, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
            
                Exit Do
            
            
            Case Else
                Call File_Error(sts, com, "資材仕入集計ﾃﾞｰﾀ")
                Exit Function
        End Select

        sts = BTRV(BtOpDelete, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
            
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "資材仕入集計ﾃﾞｰﾀ")
        End Select

    
        com = BtOpGetNext
    
    Loop
    
    For i = 0 To UBound(GENERAL_SUM)
        GENERAL_SUM(i) = 0
        NAISYOKU_SUM(i) = 0
        GENKIN_SUM(i) = 0
        SYANAI_SUM(i) = 0
        TACENTER_SUM(i) = 0
    Next i
    
    
    wKEIJYO_YM = Mid(Format(CDate(Text1(ptxKEIJYO_YM).Text & "/01"), "YYYYMMDD"), 1, 6)
    
    Call UniCode_Conv(K1_P_SHUKEIRE.KEIJYO_YM, wKEIJYO_YM)          '計上年月
                                                                    '仕入先
    Call UniCode_Conv(K1_P_SHUKEIRE.ORDER_CODE, Text1(ptxSHIIRE_CODE).Text)
    Call UniCode_Conv(K1_P_SHUKEIRE.UKEIRE_DT, "")                  '受入日

    
    com = BtOpGetGreaterEqual
    
       
       
    Do
    
        DoEvents
    
        sts = BTRV(com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K1_P_SHUKEIRE, Len(K1_P_SHUKEIRE), 1)
            
        Select Case sts
            Case BtNoErr
            
                '計上年月のﾌﾞﾚｰｸ
                If StrConv(P_SHUKEIRE_REC.KEIJYO_YM, vbUnicode) <> wKEIJYO_YM Then
                    Exit Do
                End If
            
                '仕入先のﾌﾞﾚｰｸ
                If Trim(Text1(ptxSHIIRE_CODE).Text) = "" Then
                Else
                    If Trim(Text1(ptxSHIIRE_CODE).Text) <> Trim(StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode)) Then
                        Exit Do
                    End If
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "資材受入ﾃﾞｰﾀ")
                Exit Function
        End Select
        
        SKIP_Flg = False
        Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                SKIP_Flg = True
            Case Else
                Call File_Error(sts, BtOpGetEqual, "資材注文ﾃﾞｰﾀ")
                Exit Function
        End Select
            
        If Not SKIP_Flg Then
            
            '対象ﾃﾞｰﾀ
            Data_Flg = True
                
            'ｺｰﾄﾞﾏｽﾀ読み込み
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN01_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SHORDER_REC.G_SHIIRE_KBN, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    '未登録はその他
                    Call UniCode_Conv(P_CODEREC.OPTION1, P_HN_ETC)
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "ｺｰﾄﾞﾏｽﾀ")
                    Exit Function
            End Select
    
            Select Case Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
                
                Case P_SH_SHIIRE            '一般仕入
                    
                    Select Case StrConv(P_SHORDER_REC.TORI_KBN, vbUnicode)
                        
                        Case P_TORI_GENERAL
                            GENERAL_SUM(0) = GENERAL_SUM(0) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_NAISYOKU, P_TORI_JIKYU
                            NAISYOKU_SUM(0) = NAISYOKU_SUM(0) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_GENKIN
                            GENKIN_SUM(0) = GENKIN_SUM(0) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_SYANAI
                            SYANAI_SUM(0) = SYANAI_SUM(0) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_ANOTHER
                            TACENTER_SUM(0) = TACENTER_SUM(0) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                    End Select
                
                Case P_SH_SEIZOU            '製造
                    
                    Select Case StrConv(P_SHORDER_REC.TORI_KBN, vbUnicode)
                        
                        Case P_TORI_GENERAL
                            GENERAL_SUM(1) = GENERAL_SUM(1) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_NAISYOKU, P_TORI_JIKYU
                            NAISYOKU_SUM(1) = NAISYOKU_SUM(1) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_GENKIN
                            GENKIN_SUM(1) = GENKIN_SUM(1) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_SYANAI
                            SYANAI_SUM(1) = SYANAI_SUM(1) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_ANOTHER
                            TACENTER_SUM(1) = TACENTER_SUM(1) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                    End Select
                    
                Case P_SH_YATIN             '家賃
                    
                    Select Case StrConv(P_SHORDER_REC.TORI_KBN, vbUnicode)
                        
                        Case P_TORI_GENERAL
                            GENERAL_SUM(2) = GENERAL_SUM(2) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_NAISYOKU, P_TORI_JIKYU
                            NAISYOKU_SUM(2) = NAISYOKU_SUM(2) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_GENKIN
                            GENKIN_SUM(2) = GENKIN_SUM(2) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_SYANAI
                            SYANAI_SUM(2) = SYANAI_SUM(2) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_ANOTHER
                            TACENTER_SUM(2) = TACENTER_SUM(2) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                    End Select
                    
                Case P_SH_ETC               'その他
                    
                    Select Case StrConv(P_SHORDER_REC.TORI_KBN, vbUnicode)
                        
                        Case P_TORI_GENERAL
                            GENERAL_SUM(3) = GENERAL_SUM(3) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_NAISYOKU, P_TORI_JIKYU
                            NAISYOKU_SUM(3) = NAISYOKU_SUM(3) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_GENKIN
                            GENKIN_SUM(3) = GENKIN_SUM(3) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_SYANAI
                            SYANAI_SUM(3) = SYANAI_SUM(3) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_ANOTHER
                            TACENTER_SUM(3) = TACENTER_SUM(3) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                    End Select
                
                Case P_SH_HAKEN             '派遣
                    
                    Select Case StrConv(P_SHORDER_REC.TORI_KBN, vbUnicode)
                        
                        Case P_TORI_GENERAL
                            GENERAL_SUM(4) = GENERAL_SUM(4) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_NAISYOKU, P_TORI_JIKYU
                            NAISYOKU_SUM(4) = NAISYOKU_SUM(4) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_GENKIN
                            GENKIN_SUM(4) = GENKIN_SUM(4) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_SYANAI
                            SYANAI_SUM(4) = SYANAI_SUM(4) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_ANOTHER
                            TACENTER_SUM(4) = TACENTER_SUM(4) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                    End Select
                    
                Case P_SH_KEIHI             '経費
                    
                    Select Case StrConv(P_SHORDER_REC.TORI_KBN, vbUnicode)
                        
                        Case P_TORI_GENERAL
                            GENERAL_SUM(5) = GENERAL_SUM(5) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_NAISYOKU, P_TORI_JIKYU
                            NAISYOKU_SUM(5) = NAISYOKU_SUM(5) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_GENKIN
                            GENKIN_SUM(5) = GENKIN_SUM(5) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_SYANAI
                            SYANAI_SUM(5) = SYANAI_SUM(5) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_ANOTHER
                            TACENTER_SUM(5) = TACENTER_SUM(5) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                    End Select
                
                Case P_SH_ZEI               '消費税
                    '何もしない
                
                Case Else

                    Select Case StrConv(P_SHORDER_REC.TORI_KBN, vbUnicode)
                        
                        Case P_TORI_GENERAL
                            GENERAL_SUM(3) = GENERAL_SUM(3) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_NAISYOKU, P_TORI_JIKYU
                            NAISYOKU_SUM(3) = NAISYOKU_SUM(3) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_GENKIN
                            GENKIN_SUM(3) = GENKIN_SUM(3) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_SYANAI
                            SYANAI_SUM(3) = SYANAI_SUM(3) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_ANOTHER
                            TACENTER_SUM(3) = TACENTER_SUM(3) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                    End Select
            
            End Select
            '消費税分
            
            ZEI = 0
            If Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)) = P_SH_ZEI Then
                '消費税はなにもしない
            Else
                
                If StrConv(P_SHORDER_REC.TORI_KBN, vbUnicode) = P_TORI_JIKYU Then
                '時給は何もしない
                Else
'                    YMD = StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode)
'
'                    If CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)) >= 0 Then
'                        If YMD < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
'                            ZEI = Int(CDbl(CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)) * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
'                                    CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
'                        Else
'                            ZEI = Int(CDbl(CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
'                                    CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
'                        End If
'                    Else
'
'                        wkKINGAKU = CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)) * -1
'
'                        If YMD < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
'                            ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
'                                    CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
'                        Else
'                            ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
'                                    CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
'                        End If
'                        ZEI = ZEI * -1
'                    End If
'
'
'                    Select Case StrConv(P_SHORDER_REC.TORI_KBN, vbUnicode)
'
'                        Case P_TORI_GENERAL
'                            GENERAL_SUM(6) = GENERAL_SUM(6) + ZEI
'                        Case P_TORI_NAISYOKU
'                            NAISYOKU_SUM(6) = NAISYOKU_SUM(6) + ZEI
'                        Case P_TORI_GENKIN
'                            GENKIN_SUM(6) = GENKIN_SUM(6) + ZEI
'                        Case P_TORI_SYANAI
'                            SYANAI_SUM(6) = SYANAI_SUM(6) + ZEI
'                        Case P_TORI_ANOTHER
'                            TACENTER_SUM(6) = 0
'                    End Select
                
                
                    If IsNumeric(StrConv(P_SHUKEIRE_REC.ZEI_KIN, vbUnicode)) Then
                        Select Case StrConv(P_SHORDER_REC.TORI_KBN, vbUnicode)
                            
                            Case P_TORI_GENERAL
                                GENERAL_SUM(6) = GENERAL_SUM(6) + CLng(StrConv(P_SHUKEIRE_REC.ZEI_KIN, vbUnicode))
                            Case P_TORI_NAISYOKU
                                        
                                NAISYOKU_SUM(6) = NAISYOKU_SUM(6) + CLng(StrConv(P_SHUKEIRE_REC.ZEI_KIN, vbUnicode))
                            Case P_TORI_GENKIN
                                GENKIN_SUM(6) = GENKIN_SUM(6) + CLng(StrConv(P_SHUKEIRE_REC.ZEI_KIN, vbUnicode))
                            Case P_TORI_SYANAI
                                SYANAI_SUM(6) = SYANAI_SUM(6) + CLng(StrConv(P_SHUKEIRE_REC.ZEI_KIN, vbUnicode))
                            Case P_TORI_ANOTHER
                                TACENTER_SUM(6) = 0
                        End Select
                    End If
                
                
                End If
                            
            End If
                
            '資材仕入集計ﾃﾞｰﾀ読み込み
                
            Call UniCode_Conv(K0_P_SHSHI_SUM.SHIIRE_CODE, StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode))
            Call UniCode_Conv(K0_P_SHSHI_SUM.TORI_KBN, "")
            sts = BTRV(BtOpGetEqual, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
            Select Case sts
                Case BtNoErr
                    upd_com = BtOpUpdate
                Case BtErrKeyNotFound
                    upd_com = BtOpInsert
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "資材仕入集計ﾃﾞｰﾀ")
                    Exit Function
            End Select
            
            
            If upd_com = BtOpInsert Then
            
                Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_CODE, StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode))
            
                Call UniCode_Conv(P_SHSHI_SUM_REC.TORI_KBN, "")
            
            
                For i = 0 To 6
                    Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(i).SHIIRE, "00000000")
                Next i
            
            End If
            
            
            
            
            Select Case Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
                Case P_SH_SHIIRE            '仕入
                    
                    Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(0).SHIIRE, _
                                    Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(0).SHIIRE, vbUnicode)) + _
                                    CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)), "00000000"))
                
                Case P_SH_SEIZOU            '製造
                    Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(1).SHIIRE, _
                                    Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(1).SHIIRE, vbUnicode)) + _
                                    CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)), "00000000"))
                Case P_SH_YATIN             '家賃
                    Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(2).SHIIRE, _
                                    Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(2).SHIIRE, vbUnicode)) + _
                                    CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)), "00000000"))
                Case P_SH_ETC               'その他
                    Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(3).SHIIRE, _
                                    Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(3).SHIIRE, vbUnicode)) + _
                                    CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)), "00000000"))
                Case P_SH_HAKEN             '派遣
                    Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(4).SHIIRE, _
                                    Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(4).SHIIRE, vbUnicode)) + _
                                    CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)), "00000000"))
                Case P_SH_KEIHI             '一般経費
                    Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(5).SHIIRE, _
                                    Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(5).SHIIRE, vbUnicode)) + _
                                    CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)), "00000000"))
                Case P_SH_ZEI               '消費税
                    '何もしない
                Case Else
                    Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(3).SHIIRE, _
                                    Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(3).SHIIRE, vbUnicode)) + _
                                    CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)), "00000000"))
            End Select
'            Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(6).SHIIRE, _
'                    Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(6).SHIIRE, vbUnicode)) + _
'                    ZEI, "00000000"))

            If IsNumeric(StrConv(P_SHUKEIRE_REC.ZEI_KIN, vbUnicode)) Then
                Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(6).SHIIRE, _
                        Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(6).SHIIRE, vbUnicode)) + _
                        CLng(StrConv(P_SHUKEIRE_REC.ZEI_KIN, vbUnicode)), "00000000"))
            End If



'''----------- 2006.04.23 廃止 ↓
''''----------- kubota
''''            ' 2006.03.24 kubota
'''            If Trim(StrConv(P_SHSHI_SUM_REC.SHIIRE_CODE, vbUnicode)) = "D421" Or _
'''                Trim(StrConv(P_SHSHI_SUM_REC.SHIIRE_CODE, vbUnicode)) = "F777" Or _
'''                Trim(StrConv(P_SHSHI_SUM_REC.SHIIRE_CODE, vbUnicode)) = "S414" Then
'''                Dim lngTotal As Long
'''                Dim intC As Integer
'''                lngTotal = 0
'''                For intC = 0 To 5
'''                    lngTotal = lngTotal + CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(intC).SHIIRE, vbUnicode))
'''                Next
'''                Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(6).SHIIRE, _
'''                        Format(lngTotal * 0.05, "00000000"))
''''                Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(6).SHIIRE, _
''''                        Format(1110, "00000000"))
'''            End If
''''----------- kubota
''''----------- 2006.04.23 廃止 ↑
            
            
            
            sts = BTRV(upd_com, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
            Select Case sts
                Case BtNoErr
                Case Else
                    Call File_Error(sts, upd_com, "資材仕入集計ﾃﾞｰﾀ")
                    Exit Function
            End Select
        
        End If
        
        
        
        com = BtOpGetNext
    
    Loop






'*--------------------------    消費税ﾃﾞｰﾀの集計


    wKEIJYO_YM = Mid(Format(CDate(Text1(ptxKEIJYO_YM).Text & "/01"), "YYYYMMDD"), 1, 6)
    
    Call UniCode_Conv(K1_P_SHUKEIRE.KEIJYO_YM, wKEIJYO_YM)                          '計上年月
    Call UniCode_Conv(K1_P_SHUKEIRE.ORDER_CODE, Text1(ptxSHIIRE_CODE).Text)         '仕入先 仕入先セット    2007.10.24
    Call UniCode_Conv(K1_P_SHUKEIRE.UKEIRE_DT, "")                                  '受入日

    
    com = BtOpGetGreaterEqual
    
       
       
    Do
    
        DoEvents
    
        sts = BTRV(com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K1_P_SHUKEIRE, Len(K1_P_SHUKEIRE), 1)
            
        Select Case sts
            Case BtNoErr
            
                '計上年月のﾌﾞﾚｰｸ
                If StrConv(P_SHUKEIRE_REC.KEIJYO_YM, vbUnicode) <> wKEIJYO_YM Then
                    Exit Do
                End If
            
                '仕入先のﾌﾞﾚｰｸ
                If Trim(Text1(ptxSHIIRE_CODE).Text) = "" Then
                Else
                    If Trim(Text1(ptxSHIIRE_CODE).Text) <> Trim(StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode)) Then
                        Exit Do
                    End If
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "資材受入ﾃﾞｰﾀ")
                Exit Function
        End Select
        
        SKIP_Flg = False
        Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                SKIP_Flg = True
            Case Else
                Call File_Error(sts, BtOpGetEqual, "資材注文ﾃﾞｰﾀ")
                Exit Function
        End Select
            
        If Not SKIP_Flg Then
            
            '対象ﾃﾞｰﾀ
            Data_Flg = True
                
            'ｺｰﾄﾞﾏｽﾀ読み込み
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN01_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SHORDER_REC.G_SHIIRE_KBN, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    '未登録はその他
                    Call UniCode_Conv(P_CODEREC.OPTION1, P_HN_ETC)
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "ｺｰﾄﾞﾏｽﾀ")
                    Exit Function
            End Select
    
            Select Case Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
                
                Case P_SH_SHIIRE            '一般仕入
                
                Case P_SH_SEIZOU            '製造
                    
                    
                Case P_SH_YATIN             '家賃
                    
                    
                Case P_SH_ETC               'その他
                    
                
                Case P_SH_HAKEN             '派遣
                    
                    
                Case P_SH_KEIHI             '経費
                    
                
                Case P_SH_ZEI               '消費税
                
                
                    Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(6).SHIIRE, _
                                    Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(6).SHIIRE, vbUnicode)) + _
                                    CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)), "00000000"))
                
                
                    Select Case StrConv(P_SHORDER_REC.TORI_KBN, vbUnicode)
                        
                        Case P_TORI_GENERAL
                            GENERAL_SUM(6) = GENERAL_SUM(6) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_NAISYOKU
                            NAISYOKU_SUM(6) = NAISYOKU_SUM(6) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_GENKIN
                            GENKIN_SUM(6) = GENKIN_SUM(6) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_SYANAI
                            SYANAI_SUM(6) = SYANAI_SUM(6) + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))
                        Case P_TORI_ANOTHER
                            TACENTER_SUM(6) = 0
                    End Select
                
                Case Else
            
            
            End Select
                
                
            '資材仕入集計ﾃﾞｰﾀ読み込み
                
            Call UniCode_Conv(K0_P_SHSHI_SUM.SHIIRE_CODE, StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode))
            Call UniCode_Conv(K0_P_SHSHI_SUM.TORI_KBN, "")
            sts = BTRV(BtOpGetEqual, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
            Select Case sts
                Case BtNoErr
                    upd_com = BtOpUpdate
                Case BtErrKeyNotFound
                    upd_com = BtOpInsert
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "資材仕入集計ﾃﾞｰﾀ")
                    Exit Function
            End Select
            
            
            If upd_com = BtOpInsert Then
            
                Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_CODE, StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode))
            
                Call UniCode_Conv(P_SHSHI_SUM_REC.TORI_KBN, "")
            
            
                For i = 0 To 6
                    Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(i).SHIIRE, "00000000")
                Next i
            
            End If
            
            
            
            
            Select Case Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
                Case P_SH_SHIIRE            '仕入
                
                Case P_SH_SEIZOU            '製造
                Case P_SH_YATIN             '家賃
                Case P_SH_ETC               'その他
                Case P_SH_HAKEN             '派遣
                Case P_SH_KEIHI             '一般経費
                Case P_SH_ZEI               '消費税
                    Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(6).SHIIRE, _
                                    Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(6).SHIIRE, vbUnicode)) + _
                                    CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)), "00000000"))
                Case Else
            End Select

'''----------- 2006.04.23 廃止 ↓
''''----------- kubota
''''            ' 2006.03.24 kubota
'''            If Trim(StrConv(P_SHSHI_SUM_REC.SHIIRE_CODE, vbUnicode)) = "D421" Or _
'''                Trim(StrConv(P_SHSHI_SUM_REC.SHIIRE_CODE, vbUnicode)) = "F777" Or _
'''                Trim(StrConv(P_SHSHI_SUM_REC.SHIIRE_CODE, vbUnicode)) = "S414" Then
'''                Dim lngTotal As Long
'''                Dim intC As Integer
'''                lngTotal = 0
'''                For intC = 0 To 5
'''                    lngTotal = lngTotal + CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(intC).SHIIRE, vbUnicode))
'''                Next
'''                Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(6).SHIIRE, _
'''                        Format(lngTotal * 0.05, "00000000"))
''''                Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(6).SHIIRE, _
''''                        Format(1110, "00000000"))
'''            End If
''''----------- kubota
''''----------- 2006.04.23 廃止 ↑
            
            
            
            sts = BTRV(upd_com, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
            Select Case sts
                Case BtNoErr
                Case Else
                    Call File_Error(sts, upd_com, "資材仕入集計ﾃﾞｰﾀ")
                    Exit Function
            End Select
        
        End If
        
        
        
        com = BtOpGetNext
    
    Loop



    


    If Data_Flg Then
        '合計ﾚｺｰﾄﾞ（一般）



'''----------- 2006.04.23 廃止 ↓
'----------- kubota
'        ' 2006.03.24 kubota
'        com = BtOpGetFirst
'        GENERAL_SUM(6) = 0
'        Do
'            sts = BTRV(com, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
'            If sts <> BtNoErr Then
'                Exit Do
'            End If
'            'ｺｰﾄﾞﾏｽﾀ読み込み
'            Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHSHI_SUM_REC.SHIIRE_CODE, vbUnicode))
'            sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
'            If sts = BtNoErr Then
'                Select Case Trim(StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode))
'                Case P_TORI_GENERAL$            '一般仕入
'                    GENERAL_SUM(6) = GENERAL_SUM(6) + CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(6).SHIIRE, vbUnicode))
'                End Select
'            End If


'            com = BtOpGetNext
'        Loop
'----------- kubota
'----------- 2006.04.23 廃止 ↑
        
        Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_CODE, "")
        Call UniCode_Conv(P_SHSHI_SUM_REC.TORI_KBN, P_TORI_GENERAL)
    
        For i = 0 To 6
            Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(i).SHIIRE, Format(GENERAL_SUM(i)))
        Next i
    
        sts = BTRV(BtOpInsert, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
        Select Case sts
            Case BtNoErr
            Case Else
                Call File_Error(sts, BtOpInsert, "資材仕入集計ﾃﾞｰﾀ")
                Exit Function
        End Select
        
        '合計ﾚｺｰﾄﾞ（内職）
        Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_CODE, "")
        Call UniCode_Conv(P_SHSHI_SUM_REC.TORI_KBN, P_TORI_NAISYOKU)
    
        For i = 0 To 6
            Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(i).SHIIRE, Format(NAISYOKU_SUM(i)))
        Next i
    
        sts = BTRV(BtOpInsert, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
        Select Case sts
            Case BtNoErr
            Case Else
                Call File_Error(sts, BtOpInsert, "資材仕入集計ﾃﾞｰﾀ")
                Exit Function
        End Select
        '合計ﾚｺｰﾄﾞ（現金）
        Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_CODE, "")
        Call UniCode_Conv(P_SHSHI_SUM_REC.TORI_KBN, P_TORI_GENKIN)
    
        For i = 0 To 6
            Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(i).SHIIRE, Format(GENKIN_SUM(i)))
        Next i
    
        sts = BTRV(BtOpInsert, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
        Select Case sts
            Case BtNoErr
            Case Else
                Call File_Error(sts, BtOpInsert, "資材仕入集計ﾃﾞｰﾀ")
                Exit Function
        End Select
        
        
        '合計ﾚｺｰﾄﾞ（他ｾﾝﾀｰ）
        Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_CODE, "")
        Call UniCode_Conv(P_SHSHI_SUM_REC.TORI_KBN, P_TORI_ANOTHER)
    
        For i = 0 To 6
            Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(i).SHIIRE, Format(TACENTER_SUM(i)))
        Next i
    
        sts = BTRV(BtOpInsert, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
        Select Case sts
            Case BtNoErr
            Case Else
                Call File_Error(sts, BtOpInsert, "資材仕入集計ﾃﾞｰﾀ")
                Exit Function
        End Select
        
        
        
        '合計ﾚｺｰﾄﾞ（社内）
        Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_CODE, "")
        Call UniCode_Conv(P_SHSHI_SUM_REC.TORI_KBN, P_TORI_SYANAI)
    
        For i = 0 To 6
            Call UniCode_Conv(P_SHSHI_SUM_REC.SHIIRE_TBL(i).SHIIRE, Format(SYANAI_SUM(i)))
        Next i
    
        sts = BTRV(BtOpInsert, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
        Select Case sts
            Case BtNoErr
            Case Else
                Call File_Error(sts, BtOpInsert, "資材仕入集計ﾃﾞｰﾀ")
                Exit Function
        End Select
    
    End If

    PR000271.MousePointer = vbDefault

   SUM_Make_Proc = False

End Function




Private Function G_Print_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'           印刷処理
'----------------------------------------------------------------------------

Dim Data_Flg        As Boolean


Dim rpt             As New PR00027F1
Dim f               As New PR000272
            
    
    G_Print_Proc = True
    '集計ﾃﾞｰﾀ作成
    If SUM_Make_Proc(Data_Flg) Then
        Exit Function
    End If
    
    If Data_Flg Then
       
       
        Select Case Mode
            Case 0
               Set rpt = New PR00027F1
            
                'レポートを印刷します。（true：印刷ダイアログあり false：なし）
               rpt.PrintReport False
            
               Set rpt = Nothing

            Case 1
                
                f.ARViewer1.Zoom = -2
                
                f.RunReport rpt
                f.Show vbModal
    
    
        End Select
    End If

    G_Print_Proc = False


End Function

Private Function Kingaku_Proc() As Integer
'----------------------------------------------------------------------------
'                   金額集計
'----------------------------------------------------------------------------
Dim sts As Integer
Dim com As Integer
'Dim Kin As Double

Dim kin As Currency

    
    com = BtOpGetFirst
    Do
    
        DoEvents
    
        sts = BTRV(com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K0_P_SHUKEIRE, Len(K0_P_SHUKEIRE), 0)
            
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "資材仕入ﾃﾞｰﾀ")
                Exit Function
        End Select

        '2009.11.02
'        Kin = (CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_TANKA, vbUnicode)) * CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode)))
'        If Kin < 0 Then
'            Kin = Kin * -1
'            Kin = Int(Kin + 0.5)
'
'            Kin = Kin * -1
'        Else
'
'            Kin = Int(Kin + 0.5)
'        End If



        Select Case StrConv(P_KANRIREC.SHI_MARUME, vbUnicode)
            Case "0"    '切捨て
                kin = ToRoundDown(CCur(CCur(StrConv(P_SHUKEIRE_REC.UKEIRE_TANKA, vbUnicode)) * _
                                        CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode))), 0)
            
    
            Case "5"    '四捨五入
            
                kin = ToHalfAdjust(CCur(CCur(StrConv(P_SHUKEIRE_REC.UKEIRE_TANKA, vbUnicode)) * _
                                        CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode))), 0)

    
            
            
            
            
            Case "9"    '切り上げ
        
        
                kin = ToRoundUp(CCur(CCur(StrConv(P_SHUKEIRE_REC.UKEIRE_TANKA, vbUnicode)) * _
                                        CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode))), 0)

    
    
        
        
            Case Else    '四捨五入
            
                kin = ToHalfAdjust(CCur(CCur(StrConv(P_SHUKEIRE_REC.UKEIRE_TANKA, vbUnicode)) * _
                                        CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode))), 0)
        
        
        End Select



        If kin < 0 Then
            Call UniCode_Conv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, Format(kin, "00000000"))
        Else
            Call UniCode_Conv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, Format(kin, "000000000"))
        End If

        sts = BTRV(BtOpUpdate, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K0_P_SHUKEIRE, Len(K0_P_SHUKEIRE), 0)
            
        Select Case sts
            Case BtNoErr
            Case Else
                Call File_Error(sts, com, "資材売上集計ﾃﾞｰﾀ")
                Exit Function
        End Select
    


        com = BtOpGetNext
    Loop
End Function

Private Function Data_Proc() As Integer
'----------------------------------------------------------------------------
'                   データ出力
'----------------------------------------------------------------------------

Dim FileNo          As Integer
Dim FileName        As String
Dim Ret             As Integer

Dim com             As Integer
Dim sts             As Integer

Dim wKEIJYO_YM      As String * 6
    
Dim SKIP_Flg        As Boolean
    
    
Dim wkSHIIRE_KBN    As String
    
    Call Input_Lock

'    P_SHUKEIRE_CSV = "c:\sdc_siga\work\shukeire.csv"

    FileNo = FreeFile
    FileName = P_SHUKEIRE_CSV
    
    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & Right(Trim(FileName), Len(Trim(FileName)) - Ret)

    On Error GoTo Error_Proc

    Open (FileName) For Output As FileNo
    
    Write #FileNo, "受入月日", "仕入先", "資材品番", "品名", "仕入区分", "収支単位", "数量", "単価", "金額", "消費税"

    wKEIJYO_YM = Mid(Format(CDate(Text1(ptxKEIJYO_YM).Text & "/01"), "YYYYMMDD"), 1, 6)
    
    Call UniCode_Conv(K1_P_SHUKEIRE.KEIJYO_YM, wKEIJYO_YM)                      '計上年月
    Call UniCode_Conv(K1_P_SHUKEIRE.ORDER_CODE, "")     '仕入先
    Call UniCode_Conv(K1_P_SHUKEIRE.UKEIRE_DT, "")                              '受入年月日
    
    
    com = BtOpGetGreaterEqual
    
       
       
    Do
    
        DoEvents
    
        sts = BTRV(com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K1_P_SHUKEIRE, Len(K1_P_SHUKEIRE), 1)
            
        Select Case sts
            Case BtNoErr
            
                '計上年月のﾌﾞﾚｰｸ
                If StrConv(P_SHUKEIRE_REC.KEIJYO_YM, vbUnicode) <> wKEIJYO_YM Then
                    Exit Do
                End If
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "資材受入ﾃﾞｰﾀ")
                Exit Function
        End Select
    
            '注文ﾃﾞｰﾀ読み込み
        SKIP_Flg = False
        Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                '異常データ
                SKIP_Flg = True
            Case Else
                Call File_Error(sts, BtOpGetEqual, "資材注文ﾃﾞｰﾀ")
                Exit Function
        End Select
    
        If Not SKIP_Flg Then
    
            '月日
            Write #FileNo, Mid(StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode), 5, 2) & "/" & _
                                                Mid(StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode), 7, 2),

            '仕入先
            Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
                    Exit Function
            End Select
            Write #FileNo, StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode) & " " & StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode),
            
            '品番
            Write #FileNo, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode),
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
    
            '仕入区分
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN01_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SHORDER_REC.G_SHIIRE_KBN, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                
                    wkSHIIRE_KBN = Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_CODEREC.C_RNAME, "")
                    wkSHIIRE_KBN = ""
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "コードマスタ")
                    Exit Function
            End Select
            Write #FileNo, Trim(StrConv(P_SHORDER_REC.G_SHIIRE_KBN, vbUnicode)) & " " & _
                        StrConv(P_CODEREC.C_RNAME, vbUnicode),
    
            '収支区分
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN03_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SHORDER_REC.G_SYUSHI, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_CODEREC.C_RNAME, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "コードマスタ")
                    Exit Function
            End Select
            Write #FileNo, Trim(StrConv(P_SHORDER_REC.G_SYUSHI, vbUnicode)) & " " & _
                        StrConv(P_CODEREC.C_RNAME, vbUnicode),
            '数量
            Write #FileNo, Format(CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode)), "#,##0.00"),
            '単価
            Write #FileNo, Format(CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_TANKA, vbUnicode)), "#,##0.00"),
            
            
            
            If Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)) <> P_SH_ZEI Then
                Write #FileNo, Format(CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)), "#,##0"),
            
                '消費税額   2007.08.01
                If IsNumeric(StrConv(P_SHUKEIRE_REC.ZEI_KIN, vbUnicode)) Then
                    Write #FileNo, Format(CDbl(StrConv(P_SHUKEIRE_REC.ZEI_KIN, vbUnicode)), "#,##0")
                Else
                    Write #FileNo, Format(0, "#,##0")
                End If
            
            Else
                If IsNumeric(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)) Then
                    Write #FileNo, Format(0, "#,##0"),
                    Write #FileNo, Format(CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)), "#,##0")
                Else
                    Write #FileNo, Format(0, "#,##0"),
                    Write #FileNo, Format(0, "#,##0")
                End If
            
            
            End If
            
            
            
        
        End If
    
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
        Data_Proc = False
    Else
        MsgBox "Err.Number= " & Err.Number
        Data_Proc = True
    End If

    Call Input_UnLock



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



