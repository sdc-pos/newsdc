VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PR000601 
   Caption         =   "生産実績明細書発行"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15150
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
   ScaleWidth      =   15150
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox txSEL_KEY 
      Height          =   375
      Left            =   10680
      TabIndex        =   26
      Text            =   "Text2"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "出力対象"
      Height          =   855
      Left            =   6960
      TabIndex        =   25
      Top             =   480
      Width           =   3015
      Begin VB.CheckBox Check1 
         Caption         =   "明細表"
         Height          =   375
         Index           =   3
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "集計表"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   6
      Top             =   1200
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   1
      Left            =   2520
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   7
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Caption         =   "内職"
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "外注"
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   2520
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   4
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   7455
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   2040
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   13150
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "CODE"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "手配先"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "合計"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "消費税"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "支払合計"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   15
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=15"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1561"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1455"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3201"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3096"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2381"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2275"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2381"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2275"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2381"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2275"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2381"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2275"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=2381"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2275"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=2381"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2275"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=2381"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2275"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(9).Width=2381"
      Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2275"
      Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=2"
      Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(51)=   "Column(10).Width=2381"
      Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=2275"
      Splits(0)._ColumnProps(54)=   "Column(10)._ColStyle=2"
      Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(56)=   "Column(11).Width=2381"
      Splits(0)._ColumnProps(57)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(11)._WidthInPix=2275"
      Splits(0)._ColumnProps(59)=   "Column(11)._ColStyle=2"
      Splits(0)._ColumnProps(60)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(61)=   "Column(12).Width=2381"
      Splits(0)._ColumnProps(62)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(12)._WidthInPix=2275"
      Splits(0)._ColumnProps(64)=   "Column(12)._ColStyle=2"
      Splits(0)._ColumnProps(65)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(66)=   "Column(13).Width=2381"
      Splits(0)._ColumnProps(67)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(13)._WidthInPix=2275"
      Splits(0)._ColumnProps(69)=   "Column(13)._ColStyle=2"
      Splits(0)._ColumnProps(70)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(71)=   "Column(14).Width=2381"
      Splits(0)._ColumnProps(72)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(14)._WidthInPix=2275"
      Splits(0)._ColumnProps(74)=   "Column(14)._ColStyle=2"
      Splits(0)._ColumnProps(75)=   "Column(14).Order=15"
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
      Caption         =   "生産集計明細"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFF&,.bold=0,.fontsize=975"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=128"
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
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=110,.parent=43,.alignment=0"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=107,.parent=44"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=108,.parent=45"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=109,.parent=47"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=16,.parent=43,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=13,.parent=44"
      _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=14,.parent=45"
      _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=15,.parent=47"
      _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=28,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(53)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(54)  =   ":id=28,.fontname=ＭＳ ゴシック"
      _StyleDefs(55)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=44"
      _StyleDefs(56)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=45"
      _StyleDefs(57)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=47"
      _StyleDefs(58)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(59)  =   ":id=66,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(60)  =   ":id=66,.fontname=ＭＳ ゴシック"
      _StyleDefs(61)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44"
      _StyleDefs(62)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
      _StyleDefs(63)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
      _StyleDefs(64)  =   "Splits(0).Columns(5).Style:id=32,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(65)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(66)  =   ":id=32,.fontname=ＭＳ ゴシック"
      _StyleDefs(67)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=44"
      _StyleDefs(68)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=45"
      _StyleDefs(69)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=47"
      _StyleDefs(70)  =   "Splits(0).Columns(6).Style:id=74,.parent=43,.alignment=1"
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
      _StyleDefs(86)  =   "Splits(0).Columns(10).Style:id=62,.parent=43,.alignment=1"
      _StyleDefs(87)  =   "Splits(0).Columns(10).HeadingStyle:id=59,.parent=44"
      _StyleDefs(88)  =   "Splits(0).Columns(10).FooterStyle:id=60,.parent=45"
      _StyleDefs(89)  =   "Splits(0).Columns(10).EditorStyle:id=61,.parent=47"
      _StyleDefs(90)  =   "Splits(0).Columns(11).Style:id=70,.parent=43,.alignment=1"
      _StyleDefs(91)  =   "Splits(0).Columns(11).HeadingStyle:id=67,.parent=44"
      _StyleDefs(92)  =   "Splits(0).Columns(11).FooterStyle:id=68,.parent=45"
      _StyleDefs(93)  =   "Splits(0).Columns(11).EditorStyle:id=69,.parent=47"
      _StyleDefs(94)  =   "Splits(0).Columns(12).Style:id=82,.parent=43,.alignment=1"
      _StyleDefs(95)  =   "Splits(0).Columns(12).HeadingStyle:id=79,.parent=44"
      _StyleDefs(96)  =   "Splits(0).Columns(12).FooterStyle:id=80,.parent=45"
      _StyleDefs(97)  =   "Splits(0).Columns(12).EditorStyle:id=81,.parent=47"
      _StyleDefs(98)  =   "Splits(0).Columns(13).Style:id=86,.parent=43,.alignment=1"
      _StyleDefs(99)  =   "Splits(0).Columns(13).HeadingStyle:id=83,.parent=44"
      _StyleDefs(100) =   "Splits(0).Columns(13).FooterStyle:id=84,.parent=45"
      _StyleDefs(101) =   "Splits(0).Columns(13).EditorStyle:id=85,.parent=47"
      _StyleDefs(102) =   "Splits(0).Columns(14).Style:id=90,.parent=43,.alignment=1"
      _StyleDefs(103) =   "Splits(0).Columns(14).HeadingStyle:id=87,.parent=44"
      _StyleDefs(104) =   "Splits(0).Columns(14).FooterStyle:id=88,.parent=45"
      _StyleDefs(105) =   "Splits(0).Columns(14).EditorStyle:id=89,.parent=47"
      _StyleDefs(106) =   "Named:id=33:Normal"
      _StyleDefs(107) =   ":id=33,.parent=0"
      _StyleDefs(108) =   "Named:id=34:Heading"
      _StyleDefs(109) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(110) =   ":id=34,.wraptext=-1"
      _StyleDefs(111) =   "Named:id=35:Footing"
      _StyleDefs(112) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(113) =   "Named:id=36:Selected"
      _StyleDefs(114) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(115) =   "Named:id=37:Caption"
      _StyleDefs(116) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(117) =   "Named:id=38:HighlightRow"
      _StyleDefs(118) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(119) =   "Named:id=39:EvenRow"
      _StyleDefs(120) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(121) =   "Named:id=40:OddRow"
      _StyleDefs(122) =   ":id=40,.parent=33"
      _StyleDefs(123) =   "Named:id=41:RecordSelector"
      _StyleDefs(124) =   ":id=41,.parent=34"
      _StyleDefs(125) =   "Named:id=42:FilterBar"
      _StyleDefs(126) =   ":id=42,.parent=33"
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      Index           =   7
      Left            =   6600
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      Index           =   2
      Left            =   1920
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "～"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   24
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "対象年月日"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   23
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "PR000601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'テキスト用添字
Private Const ptxS_YMD% = 0                 '開始　対象年月日
Private Const ptxE_YMD% = 1                 '終了　対象年月日

Private Const ptxGENERAL% = 2               '外注
Private Const ptxNAISYOKU% = 3              '内職



'コンボ用添字
Private Const pcmbGENERAL% = 0              '外注
Private Const pcmbNAISYOKU% = 1             '内職


'チェックボックス用添字
Private Const pchkGENERAL% = 0              '外注
Private Const pchkNAISYOKU% = 1             '内職

Private Const pchkGK% = 2                   '集計表
Private Const pchkDET% = 3                  '明細表


'Glid用環境---------------------------------

'仕入明細
Private Const pGridDETAIL% = 0


Private SEISAN      As New XArrayDB


Private Const Min_Row% = 1                  '最小行数
Private Const Min_Col% = 0                  '最小列数
Private Const Max_Col% = 14                 '最大列数

Private Const colTORI_CODE% = 0             '取引先ｺｰﾄﾞ
Private Const colTORI_NAME% = 1             '取引先名称
Private Const colSHUMUKE01_KIN% = 2         '仕向け先1
Private Const colSHUMUKE021_KIN% = 3        '仕向け先2
Private Const colSHUMUKE03_KIN% = 4         '仕向け先3
Private Const colSHUMUKE04_KIN% = 5         '仕向け先4
Private Const colSHUMUKE05_KIN% = 6         '仕向け先5
Private Const colSHUMUKE06_KIN% = 7         '仕向け先6
Private Const colSHUMUKE07_KIN% = 8         '仕向け先7
Private Const colSHUMUKE08_KIN% = 9         '仕向け先8
Private Const colSHUMUKE09_KIN% = 10        '仕向け先9
Private Const colSHUMUKE10_KIN% = 11        '仕向け先10
Private Const colTOTAL% = 12                '合計
Private Const colZEI% = 13                  '消費税額
Private Const colSHIHARAI% = 14             '支払い額




Private Sort_Tbl(Min_Col To Max_Col) _
                As Integer                  'ｿｰﾄの制御 0:昇順 1:降順
Private Tbl_Set_F   As Boolean




Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PR000601.MousePointer = vbHourglass

    Call Ctrl_Lock(PR000601)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PR000601)


    PR000601.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts     As Integer
Dim com     As Integer
    
Dim i       As Integer
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        
        Case ptxS_YMD           '対象年月日
        
            
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0000/01/01"
            End If
            
            If Not IsDate(Text1(Mode).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(Mode).SetFocus
                Exit Function
            Else
                
                Text1(Mode).Text = Format(CDate(Text1(Mode).Text), "YYYY/MM/DD")
            
            End If
        
        Case ptxE_YMD           '対象年月日
        
            
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "9999/12/31"
            End If
            
            If Not IsDate(Text1(Mode).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(Mode).SetFocus
                Exit Function
            Else
                
                Text1(Mode).Text = Format(CDate(Text1(Mode).Text), "YYYY/MM/DD")
            
            End If
        
        
        
        
        Case ptxGENERAL     '外注ｺｰﾄﾞ
           
           
            Combo1(pcmbGENERAL).ListIndex = -1
            For i = 0 To Combo1(pcmbGENERAL).ListCount - 1
                If Trim(Text1(ptxGENERAL).Text) = Trim(Right(Combo1(pcmbGENERAL).List(i), 5)) Then
                    Combo1(pcmbGENERAL).ListIndex = i
                    Exit For
                End If
            
            Next i
        
        Case ptxNAISYOKU    '内職ｺｰﾄﾞ
           
           
            Combo1(pcmbNAISYOKU).ListIndex = -1
            For i = 0 To Combo1(pcmbNAISYOKU).ListCount - 1
                If Trim(Text1(ptxNAISYOKU).Text) = Trim(Right(Combo1(pcmbNAISYOKU).List(i), 5)) Then
                    Combo1(pcmbNAISYOKU).ListIndex = i
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
        Case pcmbGENERAL        '外注
        
            Text1(ptxGENERAL).Text = Trim(Right(Combo1(pcmbGENERAL).Text, 5))
        Case pcmbNAISYOKU       '内職
        
            Text1(ptxNAISYOKU).Text = Trim(Right(Combo1(pcmbNAISYOKU).Text, 5))
    End Select
    
    
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    Select Case Index
        Case pcmbGENERAL        '外注
        
            Text1(ptxGENERAL).Text = Trim(Right(Combo1(pcmbGENERAL).Text, 5))
        Case pcmbNAISYOKU       '内職
        
            Text1(ptxNAISYOKU).Text = Trim(Right(Combo1(pcmbNAISYOKU).Text, 5))
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans         As Integer
Dim i           As Integer

Dim Data_Flg    As Boolean

Dim rpt             As New PR00060F1
Dim f               As New PR000603


Dim sts         As Integer

    Select Case Index
        Case P_CMD_Upd          '更新
        
        Case P_CMD_DEL          '削除
        
        Case P_CMD_DSP                      '検索/表示
        
            For i = ptxS_YMD To ptxNAISYOKU
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
            
            If SUM_Make_Proc(Data_Flg) Then
                Exit Sub
            End If
            
            
            If Not Data_Flg Then
                MsgBox "対象ﾃﾞｰﾀがありません"
                Exit Sub
            End If
        
        
            If List_Disp_Proc() Then
                Unload Me
            End If
        
            Text1(ptxS_YMD).SetFocus
        
        
        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        
        Case P_CMD_PRT                      '印刷
 
            For i = ptxS_YMD To ptxNAISYOKU
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
            
            If SUM_Make_Proc(Data_Flg) Then
                Exit Sub
            End If
            
            
            If Not Data_Flg Then
                MsgBox "対象ﾃﾞｰﾀがありません"
                Exit Sub
            End If
                
            ans = MsgBox("印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                
            
                If Check1(pchkGK).Value = vbChecked Then
            
                    
                    Set rpt = New PR00060F1
                
                    'レポートを印刷します。（true：印刷ダイアログあり false：なし）
                    rpt.PrintReport False
                
                    Set rpt = Nothing
                    
'                    f.RunReport rpt
'                    f.Show
                
                End If
            
                If Check1(pchkDET).Value = vbChecked Then
                    '明細表
                    If D_Print_Proc() Then
                        Unload Me
                    End If
                End If
            
            
            End If
            
            Text1(ptxS_YMD).SetFocus
            
            
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
                                'クラスマスタＯＰＥＮ
    If P_Class_Open(BtOpenNomal) Then
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
                                '管理マスタＯＰＥＮ
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '生産実績明細ﾃﾞｰﾀＯＰＥＮ
    If P_SEISAN_DET_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '生産実績明細集計ﾃﾞｰﾀＯＰＥＮ
    If P_SEISAN_GK_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '商品化指示(親)ﾃﾞｰﾀＯＰＥＮ
    If P_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '商品化指示(子)ﾃﾞｰﾀＯＰＥＮ
    If P_SSHIJI_K_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '商品化指示受入履歴ﾃﾞｰﾀＯＰＥＮ
    If P_SUKEIRE_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    Load PR000602
    Load PR000603
    
    
    
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
    
    '仕向け先設定
    If SHIMUKE_TBL_Proc(i, P_KBN04_CD) Then
        Unload Me
    End If
            
            
            
            
            
    If i = -1 Then
        MsgBox "仕向け先が設定されていません"
        Unload Me
    End If
    
    
            '外注単価変更ﾌﾗｸﾞ設定   2007.07.13
    If GetIni(App.EXEName, "GAICYU", "P_SYS", c) Then
        GAICYU_F = False
    Else
        If Trim(c) = "1" Then
            GAICYU_F = True
        Else
            GAICYU_F = False
        End If
    End If
    
    
    
    '外注先
    If UKEHARAI_Set_Proc(pcmbGENERAL, P_TORI_GENERAL) Then
        Unload Me
    End If
    '内職
    If UKEHARAI_Set_Proc(pcmbNAISYOKU, P_TORI_NAISYOKU) Then
        Unload Me
    End If
    
            
    '画面初期設定
    If Init_Proc() Then
        Unload Me
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
                                            'クラスマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "クラスマスタ")
        End If
    End If
                                            '受払先マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "受払先マスタ")
        End If
    End If
    
                                            '管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "クラスマスタ")
        End If
    End If
                                            '生産実績明細集計ﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SEISAN_GK_POS, P_SEISAN_GK_REC, Len(P_SEISAN_GK_REC), K0_P_SEISAN_GK, Len(K0_P_SEISAN_GK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "生産実績明細ﾃﾞｰﾀ")
        End If
    End If
                                            '生産実績明細データCLOSE
    sts = BTRV(BtOpClose, P_SEISAN_DET_POS, P_SEISAN_DET_REC, Len(P_SEISAN_DET_REC), K0_P_SEISAN_DET, Len(K0_P_SEISAN_DET), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "生産実績明細ﾃﾞｰﾀ")
        End If
    End If
                                            '商品化指示（親）ﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化指示（親）ﾃﾞｰﾀ")
        End If
    End If
                                            '商品化指示（子）ﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化指示（親）ﾃﾞｰﾀ")
        End If
    End If
                                            '商品化受入履歴ﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化受入履歴ﾃﾞｰﾀ")
        End If
    End If
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PR000601 = Nothing
    Set PR000602 = Nothing
    Set PR000603 = Nothing


    End
End Sub





Private Sub TDBGrid1_DblClick(Index As Integer)
    
    txSEL_KEY.Text = SEISAN(TDBGrid1(Index).Bookmark, colTORI_CODE)
    If Item_Input_Proc() Then           '明細入力
        Unload Me
    End If

End Sub

Private Sub TDBGrid1_HeadClick(Index As Integer, ByVal ColIndex As Integer)



    Select Case Index
        
        Case pGridDETAIL        '生産実績明細
    
    
            If Sort_Tbl(ColIndex) = 0 Then
                Sort_Tbl(ColIndex) = 1
            Else
                If Sort_Tbl(ColIndex) = 1 Then
                    Sort_Tbl(ColIndex) = 0
                End If
            
            End If
            
            If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then
                            
                SEISAN.QuickSort Min_Row, SEISAN.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
                
                Set TDBGrid1(Index).Array = SEISAN
                
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
    
    
    
    For i = ptxS_YMD To ptxNAISYOKU
        Text1(i).Text = ""
    Next i
    
    '処理年月日＝当日
    Text1(ptxS_YMD).Text = Format(Now, "YYYY/MM/DD")
    Text1(ptxE_YMD).Text = Format(Now, "YYYY/MM/DD")
    
    For i = pcmbGENERAL To pcmbNAISYOKU
        
        Combo1(i).ListIndex = -1
    
    Next i


    For i = pchkGENERAL To pchkDET
    
        Check1(i).Value = vbUnchecked
    Next i
    
    
    
    'ｿｰﾄ情報の初期化
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0               'ﾃﾞﾌｫﾙﾄ昇順
    Next i

    Init_Proc = False

End Function



Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'           資材受入データの表示
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim Row             As Long





    List_Disp_Proc = True
    PR000601.MousePointer = vbHourglass
        
    
    '-------------------------------------  '実績明細のｾｯﾄ
    Set SEISAN = Nothing
    
    Row = Min_Row - 1
    
    
    
    com = BtOpGetFirst
    
    
    Do
    
        DoEvents
    
        sts = BTRV(com, P_SEISAN_GK_POS, P_SEISAN_GK_REC, Len(P_SEISAN_GK_REC), K0_P_SEISAN_GK, Len(K0_P_SEISAN_GK), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "生産実績明細集計ﾃﾞｰﾀ")
                Exit Function
        End Select
    
    
        If Trim(StrConv(P_SEISAN_GK_REC.TORI_CODE, vbUnicode)) = "" Then
        Else
            Row = Row + 1
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
        End If
        
        com = BtOpGetNext
    
    
    
    Loop
    
    
    Set TDBGrid1(pGridDETAIL).Array = SEISAN
    TDBGrid1(pGridDETAIL).ReBind
    TDBGrid1(pGridDETAIL).Update
    TDBGrid1(pGridDETAIL).MoveFirst
    
    
    PR000601.MousePointer = vbDefault
    
    
    List_Disp_Proc = False
    


End Function


Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           生産実績の内容をｸﾞﾘｯﾄﾞにｾｯﾄする
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim i           As Integer
Dim j           As Integer

Dim TOTAL       As Long

Dim ZEI         As Long


    Grid_Set_Proc = True
    
    
    SEISAN.ReDim Min_Row, Row, Min_Col, Max_Col


    '取引先ｺｰﾄﾞ
    SEISAN(Row, colTORI_CODE) = StrConv(P_SEISAN_GK_REC.TORI_CODE, vbUnicode)
    '取引先名称
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SEISAN_GK_REC.TORI_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        
    Select Case sts
        Case BtNoErr
        
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
            Call UniCode_Conv(P_UKEHARAIREC.TORI_KBN, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "受払先ﾏｽﾀ")
            Exit Function
    End Select
    SEISAN(Row, colTORI_NAME) = StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode)
        
    j = colSHUMUKE01_KIN
    TOTAL = 0
    For i = 0 To UBound(SHIMUKE_TBL)
        
        '仕向け先別
        SEISAN(Row, j) = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(i).KIN, vbUnicode)), "#,##0")
        TOTAL = TOTAL + CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(i).KIN, vbUnicode))
        j = j + 1
    
    Next i
    SEISAN(Row, colTOTAL) = Format(TOTAL, "#,##0")
    
    Select Case StrConv(P_SEISAN_GK_REC.TORI_KBN, vbUnicode)
        Case P_TORI_GENERAL
            
            If GAICYU_F Then        '2007.07.17
            
                SEISAN(Row, colZEI) = ""
                
            Else
                
                ZEI = Int(Int(TOTAL * CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode) / 100)) + CInt(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode) / 10))
                SEISAN(Row, colZEI) = Format(ZEI, "#,##0")
            
            End If                  '2007.07.17
            
            
            SEISAN(Row, colSHIHARAI) = Format(TOTAL + ZEI, "#,##0")
        Case Else
            SEISAN(Row, colZEI) = ""
            SEISAN(Row, colSHIHARAI) = Format(TOTAL, "#,##0")
    End Select
    
    
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


Private Function UKEHARAI_Set_Proc(Index As Integer, KBN As String) As Integer
'----------------------------------------------------------------------------
'                   受払先マスタをコンボにセットする。
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
    
    UKEHARAI_Set_Proc = True
    
    Combo1(Index).Clear
    
    Combo1(Index).AddItem Space(5)
    
    Call UniCode_Conv(K1_P_UKEHARAI.TORI_KBN, KBN)
    Call UniCode_Conv(K1_P_UKEHARAI.UKEHARAI_CODE, "")

    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K1_P_UKEHARAI, Len(K1_P_UKEHARAI), 1)
            
        Select Case sts
            Case BtNoErr
            
                                
                If StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) <> KBN Then
                    
                    Exit Do
                
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "取引先マスタ")
                Exit Function
        
        End Select

        
        
        
        Combo1(Index).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode) & "            " & _
                                StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        
        
        com = BtOpGetNext
    
    Loop

    UKEHARAI_Set_Proc = False
    



End Function


Private Function SUM_Make_Proc(Data_Flg As Boolean) As Integer
'----------------------------------------------------------------------------
'                   生産実績集計ﾃﾞｰﾀ作成
'----------------------------------------------------------------------------
Dim sts                     As Integer
Dim com                     As Integer

Dim SKIP_Flg                As Boolean
    
    
Dim i                       As Integer
    
Dim SAVE_TORI_KBN           As String * 1
Dim SAVE_TORI_CODE          As String * 5


Dim ALL_KIN(0 To 9)         As Long
Dim ALL_CNT                 As Integer
Dim ALL_QTY                 As Double
Dim KAZEI                   As Long


Dim TOTAL_KIN(0 To 9)       As Long
Dim TOTAL_CNT               As Integer
Dim TOTAL_QTY               As Double
    
Dim wkTANKA                 As Double
    
    
    SUM_Make_Proc = True
    PR000601.MousePointer = vbHourglass

    '-----------------------------------------  集計ﾃﾞｰﾀ全件削除


    com = BtOpGetFirst



    Do
    
    
        sts = BTRV(com, P_SEISAN_GK_POS, P_SEISAN_GK_REC, Len(P_SEISAN_GK_REC), K0_P_SEISAN_GK, Len(K0_P_SEISAN_GK), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
            
                Exit Do
            
            
            Case Else
                Call File_Error(sts, com, "生産実績明細集計ﾃﾞｰﾀ")
                Exit Function
        End Select

        sts = BTRV(BtOpDelete, P_SEISAN_GK_POS, P_SEISAN_GK_REC, Len(P_SEISAN_GK_REC), K0_P_SEISAN_GK, Len(K0_P_SEISAN_GK), 0)
            
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "生産実績明細集計ﾃﾞｰﾀ")
        End Select

    
        com = BtOpGetNext
    
    Loop
    
    com = BtOpGetFirst



    Do
    
    
        sts = BTRV(com, P_SEISAN_DET_POS, P_SEISAN_DET_REC, Len(P_SEISAN_DET_REC), K0_P_SEISAN_DET, Len(K0_P_SEISAN_DET), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
            
                Exit Do
            
            
            Case Else
                Call File_Error(sts, com, "生産実績明細ﾃﾞｰﾀ")
                Exit Function
        End Select

        sts = BTRV(BtOpDelete, P_SEISAN_DET_POS, P_SEISAN_DET_REC, Len(P_SEISAN_DET_REC), K0_P_SEISAN_DET, Len(K0_P_SEISAN_DET), 0)
            
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "生産実績明細集計ﾃﾞｰﾀ")
        End Select

    
        com = BtOpGetNext
    
    Loop
    
        
    '-----------------------------------------  集計処理開始
    
    Data_Flg = False
        
    
    '----------------   外注
    If Check1(pchkGENERAL).Value = vbChecked Then
    
        Call UniCode_Conv(K2_P_SUKEIRE.TORI_CODE, Text1(ptxGENERAL).Text)
        
        If Trim(Text1(ptxGENERAL).Text) = "" Then
            Call UniCode_Conv(K2_P_SUKEIRE.UKEIRE_DT, "")
        Else
            Call UniCode_Conv(K2_P_SUKEIRE.UKEIRE_DT, Format(Text1(ptxS_YMD).Text, "YYYYMMDD"))
        End If
    
        com = BtOpGetGreaterEqual
        
        Do
        
            DoEvents
        
            sts = BTRV(com, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K2_P_SUKEIRE, Len(K2_P_SUKEIRE), 2)
                
            Select Case sts
                Case BtNoErr
                    If Trim(Text1(ptxGENERAL).Text) = "" Then
                    Else
                        If Trim(Text1(ptxGENERAL).Text) <> Trim(StrConv(P_SUKEIRE_REC.TORI_CODE, vbUnicode)) Then
                            Exit Do
                        End If
                    End If
                
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "商品化指示受入履歴")
                    Exit Function
            End Select
    
    
    
            SKIP_Flg = False
            
        
        
                    
            '受入年月日のﾌﾞﾚｰｸ
            If StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode) < Format(CDate(Text1(ptxS_YMD).Text), "YYYYMMDD") Or _
                StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode) > Format(CDate(Text1(ptxE_YMD).Text), "YYYYMMDD") Then
                SKIP_Flg = True
            End If
        
        
        
            '指示ﾃﾞｰﾀ読み込み
            Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_NO, StrConv(P_SUKEIRE_REC.SHIJI_NO, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            Select Case sts
                Case BtNoErr
                
                    If Trim(Text1(ptxGENERAL).Text) <> "" Then
                        If Trim(StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode)) <> Trim(Text1(ptxGENERAL).Text) Then
                            SKIP_Flg = True
                        End If
                    End If
                
                
                Case BtErrKeyNotFound
                    SKIP_Flg = True
                    Call UniCode_Conv(P_SSHIJI_O_REC.TORI_KBN, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "商品化指図(親)ﾃﾞｰﾀ")
                    Exit Function
            End Select
            
            If StrConv(P_SSHIJI_O_REC.TORI_KBN, vbUnicode) <> P_TORI_GENERAL Then
                SKIP_Flg = True
            End If
    
If Trim(StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode)) = "27" Then
    Debug.Print
End If
    
            If Not SKIP_Flg Then
                Data_Flg = True
                                                '取引先区分
                Call UniCode_Conv(P_SEISAN_DET_REC.TORI_KBN, StrConv(P_SSHIJI_O_REC.TORI_KBN, vbUnicode))
                                                '取引先ｺｰﾄﾞ
                Call UniCode_Conv(P_SEISAN_DET_REC.TORI_CODE, StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode))
                                                '受入日
                Call UniCode_Conv(P_SEISAN_DET_REC.UKEIRE_DT, StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode))
                                                '指図票№
                Call UniCode_Conv(P_SEISAN_DET_REC.SHIJI_NO, StrConv(P_SUKEIRE_REC.SHIJI_NO, vbUnicode))
                                                '仕向け先
                Call UniCode_Conv(P_SEISAN_DET_REC.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                                                '品番
                Call UniCode_Conv(P_SEISAN_DET_REC.HIN_GAI, StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))
                                                '数量
                Call UniCode_Conv(P_SEISAN_DET_REC.UKEIRE_QTY, StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode))
                                                
                                                '商品化ｸﾗｽ
                Call UniCode_Conv(P_SEISAN_DET_REC.S_CLASS_CODE, StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode))
                                                '付加ｸﾗｽ
                Call UniCode_Conv(P_SEISAN_DET_REC.F_CLASS_CODE, StrConv(P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode))
                                                '内職ｸﾗｽ
                Call UniCode_Conv(P_SEISAN_DET_REC.N_CLASS_CODE, StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode))
            
                wkTANKA = 0
                
                If Not GAICYU_F Then        ''2007.07.13
                
                    If Trim(StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode)) <> "" Then
                                                    '商品化単価
                        Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                        Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode))
                        sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                        Select Case sts
                            Case BtNoErr
                                wkTANKA = CDbl(StrConv(P_CLASSREC.KOURYOU, vbUnicode))
                            Case BtErrKeyNotFound
                                wkTANKA = 0
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "ｸﾗｽﾏｽﾀ")
                                Exit Function
                        End Select
                
                    End If
                
                    If Trim(StrConv(P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode)) <> "" Then
                                                    '付加単価
                        Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                        Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, StrConv(P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode))
                        sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                        Select Case sts
                            Case BtNoErr
                                wkTANKA = wkTANKA + CDbl(StrConv(P_CLASSREC.KOURYOU, vbUnicode))
                            Case BtErrKeyNotFound
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "ｸﾗｽﾏｽﾀ")
                                Exit Function
                        End Select
                
                    End If
                End If
            
                If Trim(StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode)) <> "" Then
                                                '内職単価
                    Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                    Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode))
                    sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                    Select Case sts
                        Case BtNoErr
                            wkTANKA = wkTANKA + CDbl(StrConv(P_CLASSREC.KOURYOU, vbUnicode))
                        Case BtErrKeyNotFound
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "ｸﾗｽﾏｽﾀ")
                            Exit Function
                    End Select
            
                End If
            
            
            
            
                                                '単価
                Call UniCode_Conv(P_SEISAN_DET_REC.KOURYOU, Format(wkTANKA, "00000000.00"))
                                                '金額
                Call UniCode_Conv(P_SEISAN_DET_REC.KIN, Format(wkTANKA * CDbl(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode)), "00000000000"))
                                                            
            
            
                sts = BTRV(BtOpInsert, P_SEISAN_DET_POS, P_SEISAN_DET_REC, Len(P_SEISAN_DET_REC), K0_P_SEISAN_DET, Len(K0_P_SEISAN_DET), 0)
                Select Case sts
                    Case BtNoErr
                    Case Else
                        Call File_Error(sts, BtOpInsert, "生産実績明細ﾃﾞｰﾀ")
                        Exit Function
                End Select
            
            
            End If
    
            com = BtOpGetNext
    
        Loop
    End If
    
    '----------------   内職
    If Check1(pchkNAISYOKU).Value = vbChecked Then
    
        Call UniCode_Conv(K2_P_SUKEIRE.TORI_CODE, Text1(ptxNAISYOKU).Text)
        Call UniCode_Conv(K2_P_SUKEIRE.UKEIRE_DT, "")
    
    
        com = BtOpGetGreaterEqual
        
        Do
        
            DoEvents
        
            sts = BTRV(com, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K2_P_SUKEIRE, Len(K2_P_SUKEIRE), 2)
                
            Select Case sts
                Case BtNoErr
                    
                
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "商品化指示受入履歴")
                    Exit Function
            End Select
    
    
    
            SKIP_Flg = False
    
    
    
    
            '受入年月日のﾌﾞﾚｰｸ
            If StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode) < Format(CDate(Text1(ptxS_YMD).Text), "YYYYMMDD") Or _
                StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode) > Format(CDate(Text1(ptxE_YMD).Text), "YYYYMMDD") Then
                SKIP_Flg = True
            End If
        
            '指示ﾃﾞｰﾀ読み込み
            Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_NO, StrConv(P_SUKEIRE_REC.SHIJI_NO, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            Select Case sts
                Case BtNoErr
                
                
If Trim(StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode)) = "02" Then
Debug.Print
End If
                
                    If Trim(Text1(ptxNAISYOKU).Text) <> "" Then
                        If Trim(StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode)) <> Trim(Text1(ptxNAISYOKU).Text) Then
                            SKIP_Flg = True
                        End If
                    End If
                
                
                
                Case BtErrKeyNotFound
                    SKIP_Flg = True
                    Call UniCode_Conv(P_SSHIJI_O_REC.TORI_KBN, "")
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "商品化指図(親)ﾃﾞｰﾀ")
                    Exit Function
            End Select
            
            If StrConv(P_SSHIJI_O_REC.TORI_KBN, vbUnicode) <> P_TORI_NAISYOKU Then
                SKIP_Flg = True
            End If
    
If Trim(StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode)) = "27" Then
    Debug.Print
End If
    
    
            If Not SKIP_Flg Then
                Data_Flg = True
                                                '取引先区分
                Call UniCode_Conv(P_SEISAN_DET_REC.TORI_KBN, StrConv(P_SSHIJI_O_REC.TORI_KBN, vbUnicode))
                                                '取引先ｺｰﾄﾞ
                Call UniCode_Conv(P_SEISAN_DET_REC.TORI_CODE, StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode))
                                                '受入日
                Call UniCode_Conv(P_SEISAN_DET_REC.UKEIRE_DT, StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode))
                                                '指図票№
                Call UniCode_Conv(P_SEISAN_DET_REC.SHIJI_NO, StrConv(P_SUKEIRE_REC.SHIJI_NO, vbUnicode))
                                                '仕向け先
                Call UniCode_Conv(P_SEISAN_DET_REC.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                                                '品番
                Call UniCode_Conv(P_SEISAN_DET_REC.HIN_GAI, StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))
                                                '数量
                Call UniCode_Conv(P_SEISAN_DET_REC.UKEIRE_QTY, StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode))
                                                
                                                '商品化ｸﾗｽ
                Call UniCode_Conv(P_SEISAN_DET_REC.S_CLASS_CODE, StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode))
                                                '付加ｸﾗｽ
                Call UniCode_Conv(P_SEISAN_DET_REC.F_CLASS_CODE, StrConv(P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode))
                                                '内職ｸﾗｽ
                Call UniCode_Conv(P_SEISAN_DET_REC.N_CLASS_CODE, StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode))
            
                wkTANKA = 0
            
'                If Trim(StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode)) <> "" Then
'                                                '商品化単価
'                    Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
'                    Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode))
'                    sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
'                    Select Case sts
'                        Case BtNoErr
'                            wkTANKA = CDbl(StrConv(P_CLASSREC.KOURYOU, vbUnicode))
'                        Case BtErrKeyNotFound
'                            wkTANKA = 0
'                        Case Else
'                            Call File_Error(sts, BtOpGetEqual, "ｸﾗｽﾏｽﾀ")
'                            Exit Function
'                    End Select
'
'                End If
            
'                If Trim(StrConv(P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode)) <> "" Then
'                                                '付加単価
'                    Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
'                    Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, StrConv(P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode))
'                    sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
'                    Select Case sts
'                        Case BtNoErr
'                            wkTANKA = wkTANKA + CDbl(StrConv(P_CLASSREC.KOURYOU, vbUnicode))
'                        Case BtErrKeyNotFound
'                        Case Else
'                            Call File_Error(sts, BtOpGetEqual, "ｸﾗｽﾏｽﾀ")
'                            Exit Function
'                    End Select
'
'                End If
            
                
If StrConv(P_SSHIJI_O_REC.SHIJI_NO, vbUnicode) = "00233" Then
    Debug.Print
End If
                
                If Trim(StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode)) <> "" Then
                                                '内職単価
                    Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                    Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode))
                    sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                    Select Case sts
                        Case BtNoErr
                            wkTANKA = wkTANKA + CDbl(StrConv(P_CLASSREC.KOURYOU, vbUnicode))
                        Case BtErrKeyNotFound
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "ｸﾗｽﾏｽﾀ")
                            Exit Function
                    End Select
            
                End If
            
                                                '単価
                Call UniCode_Conv(P_SEISAN_DET_REC.KOURYOU, Format(wkTANKA, "00000000.00"))
                                                '金額
                Call UniCode_Conv(P_SEISAN_DET_REC.KIN, Format(wkTANKA * CDbl(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode)), "00000000000"))
                                                            
            
            
                sts = BTRV(BtOpInsert, P_SEISAN_DET_POS, P_SEISAN_DET_REC, Len(P_SEISAN_DET_REC), K0_P_SEISAN_DET, Len(K0_P_SEISAN_DET), 0)
                Select Case sts
                    Case BtNoErr
                    Case Else
                        Call File_Error(sts, BtOpInsert, "生産実績明細ﾃﾞｰﾀ")
                        Exit Function
                End Select
            
            
            End If
    
            com = BtOpGetNext
    
        Loop
    End If
    
    
    
    
    SAVE_TORI_CODE = ""
    
    ALL_CNT = 0
    ALL_QTY = 0
    For i = 0 To UBound(ALL_KIN)
        ALL_KIN(i) = 0
    Next i
    KAZEI = 0
    
    
    TOTAL_CNT = 0
    TOTAL_QTY = 0
    For i = 0 To UBound(TOTAL_KIN)
        TOTAL_KIN(i) = 0
    Next i
        
    
    com = BtOpGetFirst
    
    
    
    Do
    
    
        sts = BTRV(com, P_SEISAN_DET_POS, P_SEISAN_DET_REC, Len(P_SEISAN_DET_REC), K0_P_SEISAN_DET, Len(K0_P_SEISAN_DET), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
            
                Exit Do
            
            
            Case Else
                Call File_Error(sts, com, "生産実績明細ﾃﾞｰﾀ")
                Exit Function
        End Select

    
        If com = BtOpGetFirst Then
            SAVE_TORI_KBN = StrConv(P_SEISAN_DET_REC.TORI_KBN, vbUnicode)
            SAVE_TORI_CODE = StrConv(P_SEISAN_DET_REC.TORI_CODE, vbUnicode)
        End If
    
    
        If SAVE_TORI_CODE <> StrConv(P_SEISAN_DET_REC.TORI_CODE, vbUnicode) Then
    
            If Sum_Total_Make_Proc(SAVE_TORI_KBN, SAVE_TORI_CODE, TOTAL_KIN(), TOTAL_CNT, TOTAL_QTY, 0) Then
                Exit Function
            End If
        
            ALL_CNT = ALL_CNT + TOTAL_CNT
            ALL_QTY = ALL_QTY + TOTAL_CNT
            
            For i = 0 To UBound(TOTAL_KIN)
                ALL_KIN(i) = ALL_KIN(i) + TOTAL_KIN(i)
            Next i
        
            If SAVE_TORI_KBN = P_TORI_NAISYOKU Then
            Else
                For i = 0 To UBound(TOTAL_KIN)
                    KAZEI = KAZEI + TOTAL_KIN(i)
                Next i
            End If
        
            TOTAL_CNT = 0
            TOTAL_QTY = 0
            For i = 0 To UBound(TOTAL_KIN)
                TOTAL_KIN(i) = 0
            Next i
        
            SAVE_TORI_KBN = StrConv(P_SEISAN_DET_REC.TORI_KBN, vbUnicode)
            SAVE_TORI_CODE = StrConv(P_SEISAN_DET_REC.TORI_CODE, vbUnicode)
        
        
        End If
        
        For i = 0 To UBound(SHIMUKE_TBL)
        
        
            If StrConv(P_SEISAN_DET_REC.SHIMUKE_CODE, vbUnicode) = SHIMUKE_TBL(i) Then
                TOTAL_KIN(i) = TOTAL_KIN(i) + CLng(StrConv(P_SEISAN_DET_REC.KIN, vbUnicode))
                Exit For
            End If
        
        Next i
        TOTAL_CNT = TOTAL_CNT + 1
        TOTAL_QTY = TOTAL_QTY + CDbl(StrConv(P_SEISAN_DET_REC.UKEIRE_QTY, vbUnicode))
        
        
        com = BtOpGetNext
    
    Loop
    
    If com <> BtOpGetFirst Then
        If Sum_Total_Make_Proc(SAVE_TORI_KBN, SAVE_TORI_CODE, TOTAL_KIN(), TOTAL_CNT, TOTAL_QTY, 0) Then
            Exit Function
        End If
    
        ALL_CNT = ALL_CNT + TOTAL_CNT
        ALL_QTY = ALL_QTY + TOTAL_CNT
        
        For i = 0 To UBound(TOTAL_KIN)
            ALL_KIN(i) = ALL_KIN(i) + TOTAL_KIN(i)
        Next i
    
        If SAVE_TORI_KBN = P_TORI_NAISYOKU Then
        Else
            For i = 0 To UBound(TOTAL_KIN)
                KAZEI = KAZEI + TOTAL_KIN(i)
            Next i
        End If
    
        If Sum_Total_Make_Proc("", "", ALL_KIN(), ALL_CNT, ALL_QTY, KAZEI) Then
            Exit Function
        End If
    
    
    
    
    End If
    
    
    
    

    PR000601.MousePointer = vbDefault

   SUM_Make_Proc = False

End Function






Private Function Sum_Total_Make_Proc(TORI_KBN As String, TORI_CODE As String, TOTAL_KIN() As Long, CNT As Integer, QTY As Double, KAZEI As Long) As Integer
'----------------------------------------------------------------------------
'           合計ﾚｺｰﾄﾞ出力
'----------------------------------------------------------------------------
Dim i   As Integer
Dim sts As Integer
    
    Sum_Total_Make_Proc = True

    Call UniCode_Conv(P_SEISAN_GK_REC.TORI_KBN, TORI_KBN)       '取引先区分
    Call UniCode_Conv(P_SEISAN_GK_REC.TORI_CODE, TORI_CODE)     '取引先ｺｰﾄﾞ
                                                                
    For i = 0 To 9
        Call UniCode_Conv(P_SEISAN_GK_REC.UCHIWAKE_TBL(i).KIN, "00000000000")
    Next i
                                                                
                                                                
                                                                '仕向け先別金額
    For i = 0 To UBound(TOTAL_KIN)
    
        Call UniCode_Conv(P_SEISAN_GK_REC.UCHIWAKE_TBL(i).KIN, Format(TOTAL_KIN(i), "00000000000"))
    
    Next i
                                                                '件数
    Call UniCode_Conv(P_SEISAN_GK_REC.CNT, Format(CNT, "00000000000"))
                                                                '数量
    Call UniCode_Conv(P_SEISAN_GK_REC.QTY, Format(QTY, "00000000.00"))
                                                                
                                                                
                                                                    '課税対象
    Call UniCode_Conv(P_SEISAN_GK_REC.KAZEI, Format(KAZEI, "00000000000"))


    sts = BTRV(BtOpInsert, P_SEISAN_GK_POS, P_SEISAN_GK_REC, Len(P_SEISAN_GK_REC), K0_P_SEISAN_GK, Len(K0_P_SEISAN_GK), 0)
    Select Case sts
        Case BtNoErr
        Case Else
            Call File_Error(sts, BtOpInsert, "生産実績明細集計ﾃﾞｰﾀ")
            Exit Function
    End Select

    Sum_Total_Make_Proc = False

End Function



Private Function SHIMUKE_TBL_Proc(i As Integer, KBN As String) As Integer
'----------------------------------------------------------------------------
'           仕向け先のﾃｰﾌﾞﾙ
'----------------------------------------------------------------------------

Dim com     As Integer
Dim sts     As Integer
Dim j       As Integer

    SHIMUKE_TBL_Proc = True

    ReDim Preserve SHIMUKE_TBL(0 To 9)

    For j = 0 To UBound(SHIMUKE_TBL)
        SHIMUKE_TBL(j) = ""
    Next j


    Call UniCode_Conv(K0_P_CODE.DATA_KBN, KBN)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")

    com = BtOpGetGreaterEqual
    i = -1

    Do
        DoEvents
    
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
                
                If Trim(StrConv(P_CODEREC.DATA_KBN, vbUnicode)) <> KBN Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "ｺｰﾄﾞﾏｽﾀ")
                Exit Function
        End Select
    
        i = i + 1
        
        SHIMUKE_TBL(i) = Left(StrConv(P_CODEREC.C_Code, vbUnicode), 2)
            
    
    
        com = BtOpGetNext
    
    Loop

    If i = -1 Then
        SHIMUKE_TBL_Proc = False
        Exit Function
    End If
    
    
    j = colSHUMUKE01_KIN
    For i = 0 To UBound(SHIMUKE_TBL)
        
        If Trim(SHIMUKE_TBL(i)) = "" Then
            TDBGrid1(pGridDETAIL).Columns(j).Visible = False
        Else
        
            TDBGrid1(pGridDETAIL).Columns(j).Visible = True
            TDBGrid1(pGridDETAIL).Columns(j).Caption = SHIMUKE_TBL(i)
        End If
        j = j + 1
    Next i

    SHIMUKE_TBL_Proc = False

End Function
Private Function Item_Input_Proc() As Integer
'----------------------------------------------------------------------------
'                   作業管理明細入力画面　表示
'----------------------------------------------------------------------------
    Item_Input_Proc = True

    
    PR000602.Show vbModal                       '明細入力フォーム表示
    If G_SCREEN_FLG = SYS_ERR Then
        Exit Function
    End If

    

    Item_Input_Proc = False

End Function

Private Function D_Print_Proc() As Integer
'----------------------------------------------------------------------------
'           印刷処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer




Dim rpt             As New PR00060F2
Dim f               As New PR000603
            
    
    D_Print_Proc = True
            
        
    com = BtOpGetFirst
    
    Do
        
        DoEvents
        
        sts = BTRV(com, P_SEISAN_GK_POS, P_SEISAN_GK_REC, Len(P_SEISAN_GK_REC), K0_P_SEISAN_GK, Len(K0_P_SEISAN_GK), 0)
            
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "生産実績明細集計ﾃﾞｰﾀ")
                Exit Function
        End Select
    
        
        If Trim(StrConv(P_SEISAN_GK_REC.TORI_CODE, vbUnicode)) = "" Then
        Else
    
            Set rpt = New PR00060F2
        
            'レポートを印刷します。（true：印刷ダイアログあり false：なし）
            rpt.PrintReport False
        
            Set rpt = Nothing
            
            
'            f.RunReport rpt
'            f.Show
        End If
    
        com = BtOpGetNext
    
    Loop
        
        
 
 
 
 
 
    D_Print_Proc = False



End Function

