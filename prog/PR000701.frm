VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PR000701 
   Caption         =   "資材発注検討処理"
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
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   2
      Left            =   7035
      MaxLength       =   10
      TabIndex        =   18
      Top             =   240
      Width           =   1335
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
      TabIndex        =   2
      Top             =   1200
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
      Columns(1).Caption=   "品名"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "9999/99"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "9999/99"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "9999/99"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "平均/月"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "LT"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "収支"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "基準在庫"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "理論在庫"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "ﾛｯﾄ"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "発注先"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "残　数量"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "理論発注数　"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "確定発注数　"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "単価"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "金額"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   17
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=17"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1005"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=900"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3201"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3096"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1614"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1508"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=1614"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1508"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=1614"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1508"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=1614"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1508"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=926"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=820"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=900"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=794"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=1773"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=1667"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(9).Width=1693"
      Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=1588"
      Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=2"
      Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(51)=   "Column(10).Width=1693"
      Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=1588"
      Splits(0)._ColumnProps(54)=   "Column(10)._ColStyle=2"
      Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(56)=   "Column(11).Width=1508"
      Splits(0)._ColumnProps(57)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(11)._WidthInPix=1402"
      Splits(0)._ColumnProps(59)=   "Column(11)._ColStyle=2"
      Splits(0)._ColumnProps(60)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(61)=   "Column(12).Width=1693"
      Splits(0)._ColumnProps(62)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(12)._WidthInPix=1588"
      Splits(0)._ColumnProps(64)=   "Column(12)._ColStyle=2"
      Splits(0)._ColumnProps(65)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(66)=   "Column(13).Width=1693"
      Splits(0)._ColumnProps(67)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(13)._WidthInPix=1588"
      Splits(0)._ColumnProps(69)=   "Column(13)._ColStyle=2"
      Splits(0)._ColumnProps(70)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(71)=   "Column(14).Width=1693"
      Splits(0)._ColumnProps(72)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(14)._WidthInPix=1588"
      Splits(0)._ColumnProps(74)=   "Column(14)._ColStyle=2"
      Splits(0)._ColumnProps(75)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(76)=   "Column(15).Width=1826"
      Splits(0)._ColumnProps(77)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(78)=   "Column(15)._WidthInPix=1720"
      Splits(0)._ColumnProps(79)=   "Column(15)._ColStyle=2"
      Splits(0)._ColumnProps(80)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(81)=   "Column(16).Width=1826"
      Splits(0)._ColumnProps(82)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(83)=   "Column(16)._WidthInPix=1720"
      Splits(0)._ColumnProps(84)=   "Column(16)._ColStyle=2"
      Splits(0)._ColumnProps(85)=   "Column(16).Order=17"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
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
      _StyleDefs(24)  =   "Splits(0).Style:id=17,.parent=1,.bold=0,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(25)  =   ":id=17,.strikethrough=0,.charset=128"
      _StyleDefs(26)  =   ":id=17,.fontname=ＭＳ ゴシック"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=26,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=18,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=19,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=20,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=22,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=21,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=23,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=24,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=25,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=27,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=28,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=32,.parent=17,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(39)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(40)  =   ":id=32,.fontname=ＭＳ ゴシック"
      _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=18"
      _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=19"
      _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=21"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=58,.parent=17,.alignment=0"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=18"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=19"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=21"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=62,.parent=17,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=18"
      _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=19"
      _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=21"
      _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=66,.parent=17,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(53)  =   ":id=66,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(54)  =   ":id=66,.fontname=ＭＳ ゴシック"
      _StyleDefs(55)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=18"
      _StyleDefs(56)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=19"
      _StyleDefs(57)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=21"
      _StyleDefs(58)  =   "Splits(0).Columns(4).Style:id=70,.parent=17,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(59)  =   ":id=70,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(60)  =   ":id=70,.fontname=ＭＳ ゴシック"
      _StyleDefs(61)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=18"
      _StyleDefs(62)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=19"
      _StyleDefs(63)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=21"
      _StyleDefs(64)  =   "Splits(0).Columns(5).Style:id=74,.parent=17,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(65)  =   ":id=74,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(66)  =   ":id=74,.fontname=ＭＳ ゴシック"
      _StyleDefs(67)  =   "Splits(0).Columns(5).HeadingStyle:id=71,.parent=18"
      _StyleDefs(68)  =   "Splits(0).Columns(5).FooterStyle:id=72,.parent=19"
      _StyleDefs(69)  =   "Splits(0).Columns(5).EditorStyle:id=73,.parent=21"
      _StyleDefs(70)  =   "Splits(0).Columns(6).Style:id=82,.parent=17,.alignment=1"
      _StyleDefs(71)  =   "Splits(0).Columns(6).HeadingStyle:id=79,.parent=18"
      _StyleDefs(72)  =   "Splits(0).Columns(6).FooterStyle:id=80,.parent=19"
      _StyleDefs(73)  =   "Splits(0).Columns(6).EditorStyle:id=81,.parent=21"
      _StyleDefs(74)  =   "Splits(0).Columns(7).Style:id=86,.parent=17,.alignment=1"
      _StyleDefs(75)  =   "Splits(0).Columns(7).HeadingStyle:id=83,.parent=18"
      _StyleDefs(76)  =   "Splits(0).Columns(7).FooterStyle:id=84,.parent=19"
      _StyleDefs(77)  =   "Splits(0).Columns(7).EditorStyle:id=85,.parent=21"
      _StyleDefs(78)  =   "Splits(0).Columns(8).Style:id=90,.parent=17,.alignment=1"
      _StyleDefs(79)  =   "Splits(0).Columns(8).HeadingStyle:id=87,.parent=18"
      _StyleDefs(80)  =   "Splits(0).Columns(8).FooterStyle:id=88,.parent=19"
      _StyleDefs(81)  =   "Splits(0).Columns(8).EditorStyle:id=89,.parent=21"
      _StyleDefs(82)  =   "Splits(0).Columns(9).Style:id=110,.parent=17,.alignment=1"
      _StyleDefs(83)  =   "Splits(0).Columns(9).HeadingStyle:id=107,.parent=18"
      _StyleDefs(84)  =   "Splits(0).Columns(9).FooterStyle:id=108,.parent=19"
      _StyleDefs(85)  =   "Splits(0).Columns(9).EditorStyle:id=109,.parent=21"
      _StyleDefs(86)  =   "Splits(0).Columns(10).Style:id=170,.parent=17,.alignment=1"
      _StyleDefs(87)  =   "Splits(0).Columns(10).HeadingStyle:id=167,.parent=18"
      _StyleDefs(88)  =   "Splits(0).Columns(10).FooterStyle:id=168,.parent=19"
      _StyleDefs(89)  =   "Splits(0).Columns(10).EditorStyle:id=169,.parent=21"
      _StyleDefs(90)  =   "Splits(0).Columns(11).Style:id=174,.parent=17,.alignment=1"
      _StyleDefs(91)  =   "Splits(0).Columns(11).HeadingStyle:id=171,.parent=18"
      _StyleDefs(92)  =   "Splits(0).Columns(11).FooterStyle:id=172,.parent=19"
      _StyleDefs(93)  =   "Splits(0).Columns(11).EditorStyle:id=173,.parent=21"
      _StyleDefs(94)  =   "Splits(0).Columns(12).Style:id=178,.parent=17,.alignment=1"
      _StyleDefs(95)  =   "Splits(0).Columns(12).HeadingStyle:id=175,.parent=18"
      _StyleDefs(96)  =   "Splits(0).Columns(12).FooterStyle:id=176,.parent=19"
      _StyleDefs(97)  =   "Splits(0).Columns(12).EditorStyle:id=177,.parent=21"
      _StyleDefs(98)  =   "Splits(0).Columns(13).Style:id=186,.parent=17,.alignment=1"
      _StyleDefs(99)  =   "Splits(0).Columns(13).HeadingStyle:id=183,.parent=18"
      _StyleDefs(100) =   "Splits(0).Columns(13).FooterStyle:id=184,.parent=19"
      _StyleDefs(101) =   "Splits(0).Columns(13).EditorStyle:id=185,.parent=21"
      _StyleDefs(102) =   "Splits(0).Columns(14).Style:id=16,.parent=17,.alignment=1"
      _StyleDefs(103) =   "Splits(0).Columns(14).HeadingStyle:id=13,.parent=18"
      _StyleDefs(104) =   "Splits(0).Columns(14).FooterStyle:id=14,.parent=19"
      _StyleDefs(105) =   "Splits(0).Columns(14).EditorStyle:id=15,.parent=21"
      _StyleDefs(106) =   "Splits(0).Columns(15).Style:id=114,.parent=17,.alignment=1"
      _StyleDefs(107) =   "Splits(0).Columns(15).HeadingStyle:id=111,.parent=18"
      _StyleDefs(108) =   "Splits(0).Columns(15).FooterStyle:id=112,.parent=19"
      _StyleDefs(109) =   "Splits(0).Columns(15).EditorStyle:id=113,.parent=21"
      _StyleDefs(110) =   "Splits(0).Columns(16).Style:id=118,.parent=17,.alignment=1"
      _StyleDefs(111) =   "Splits(0).Columns(16).HeadingStyle:id=115,.parent=18"
      _StyleDefs(112) =   "Splits(0).Columns(16).FooterStyle:id=116,.parent=19"
      _StyleDefs(113) =   "Splits(0).Columns(16).EditorStyle:id=117,.parent=21"
      _StyleDefs(114) =   "Named:id=33:Normal"
      _StyleDefs(115) =   ":id=33,.parent=0"
      _StyleDefs(116) =   "Named:id=34:Heading"
      _StyleDefs(117) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(118) =   ":id=34,.wraptext=-1"
      _StyleDefs(119) =   "Named:id=35:Footing"
      _StyleDefs(120) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(121) =   "Named:id=36:Selected"
      _StyleDefs(122) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(123) =   "Named:id=37:Caption"
      _StyleDefs(124) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(125) =   "Named:id=38:HighlightRow"
      _StyleDefs(126) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(127) =   "Named:id=39:EvenRow"
      _StyleDefs(128) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(129) =   "Named:id=40:OddRow"
      _StyleDefs(130) =   ":id=40,.parent=33"
      _StyleDefs(131) =   "Named:id=41:RecordSelector"
      _StyleDefs(132) =   ":id=41,.parent=34"
      _StyleDefs(133) =   "Named:id=42:FilterBar"
      _StyleDefs(134) =   ":id=42,.parent=33"
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
      TabIndex        =   14
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
      Index           =   9
      Left            =   8760
      TabIndex        =   12
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
      Index           =   7
      Left            =   6600
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
      Index           =   6
      Left            =   5760
      TabIndex        =   9
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "新 規"
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
      TabIndex        =   7
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "再計算"
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
      Left            =   1920
      TabIndex        =   5
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   9720
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
      Left            =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "合計金額"
      Height          =   255
      Index           =   0
      Left            =   5775
      TabIndex        =   17
      Top             =   360
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "～"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   16
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "対象年月日"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   15
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "PR000701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim YOIN_TBL        As Variant              '対象要因
Dim REC_DAYS        As Integer              '基準日数

'テキスト用添字
Private Const ptxS_YMD% = 0                 '開始　対象年月日
Private Const ptxE_YMD% = 1                 '終了　対象年月日

Private Const ptxTOTAL% = 2                 '合計




'Glid用環境---------------------------------

Private Const pGridDETAIL% = 0


Private SEISAN      As New XArrayDB


Private Const Min_Row% = 1                  '最小行数
Private Const Min_Col% = 0                  '最小列数
Private Const Max_Col% = 16                 '最大列数

Private Const colHIN_GAI% = 0               '品番(外部)
Private Const colHIN_NAME% = 1              '品名
Private Const colJITU_QTY1% = 2             '前々月
Private Const colJITU_QTY2% = 3             '前月
Private Const colJITU_QTY3% = 4             '当月
Private Const colJITU_AVE% = 5              '平均

Private Const colLT_DAYS% = 6               'LT 日数

Private Const colSYUSHI_CODE% = 7           '収支

Private Const colZAIKO_STANDARD% = 8        '基準在庫
Private Const colZAIKO_QTY% = 9             '理論在庫

Private Const colLOT% = 10                  'ﾛｯﾄ

Private Const colORDER_CODE% = 11           '発注先

Private Const colSHIJI_Z_QTY% = 12          '発注残　数量

Private Const colSHIJI_QTY_R% = 13          '論理発注数
Private Const colSHIJI_QTY_K% = 14          '確定発注数

Private Const colTANKA% = 15                '単価
Private Const colKINGAKU% = 16              '金額




Private Sort_Tbl(Min_Col To Max_Col) _
                As Integer                  'ｿｰﾄの制御 0:昇順 1:降順
Private Tbl_Set_F   As Boolean

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PR000701.MousePointer = vbHourglass

    Call Ctrl_Lock(PR000701)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PR000701)


    PR000701.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts     As Integer
Dim com     As Integer
    
Dim i       As Integer
    
    Error_Check_Proc = True
    
    Select Case mode
    
        
        Case ptxS_YMD           '対象年月日
        
            
            If Text1(mode).Text = "" Then
                Text1(mode).Text = "0000/01/01"
            End If
            
            If Not IsDate(Text1(mode).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(mode).SetFocus
                Exit Function
            Else
                
                Text1(mode).Text = Format(CDate(Text1(mode).Text), "YYYY/MM/DD")
            
            End If
        
        Case ptxE_YMD           '対象年月日
        
            
            If Text1(mode).Text = "" Then
                Text1(mode).Text = "9999/12/31"
            End If
            
            If Not IsDate(Text1(mode).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(mode).SetFocus
                Exit Function
            Else
                
                Text1(mode).Text = Format(CDate(Text1(mode).Text), "YYYY/MM/DD")
            
            End If
        
    End Select
        
        
    Error_Check_Proc = False
    

End Function

Private Sub Command1_Click(Index As Integer)

Dim ans         As Integer
Dim i           As Integer

Dim Data_Flg    As Boolean

Dim rpt             As New PR00070F1
Dim f               As New PR000702


Dim sts         As Integer

    Select Case Index
        Case P_CMD_Upd          '更新
        
            For i = ptxS_YMD To ptxTOTAL
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
        
            Call RE_SUM_PROC
        
            If Update_Proc() Then
                Exit Sub
            End If
                    
            If List_Disp_Proc() Then
                Unload Me
            End If
        
            Text1(ptxS_YMD).SetFocus
        
        Case 2                  '再計算
        
            Call RE_SUM_PROC
        
        
        
        
        Case P_CMD_DEL          '削除
        
        
        Case P_CMD_DSP                      '検索/表示
        
            For i = ptxS_YMD To ptxTOTAL
            
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
 
            For i = ptxS_YMD To ptxTOTAL
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
            
                
                
            ans = MsgBox("印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                
                
                
                Call RE_SUM_PROC
            
                If Update_Proc() Then
                    Exit Sub
                End If
                        
                
                GLB_S_YMD = Text1(ptxS_YMD).Text
                GLB_E_YMD = Text1(ptxE_YMD).Text
            
                GLB_TOTAL_KINGAKU = CLng(Text1(ptxTOTAL).Text)
                    
                Set rpt = New PR00070F1
        
            'レポートを印刷します。（true：印刷ダイアログあり false：なし）
                rpt.PrintReport False
        
                Set rpt = Nothing
            
'                    f.RunReport rpt
'                    f.Show
                
            
            
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
    
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '管理マスタＯＰＥＮ
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '在庫移動歴ＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫ﾃﾞｰﾀＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '資材発注ﾃﾞｰﾀＯＰＥＮ
    If P_SHORDER_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '資材発注検討ﾌｧｲﾙＯＰＥＮ
    If P_SHKENTO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    
    Load PR000702
    
    
                                '対象要因取り込み
    If GetIni(App.EXEName, "YOIN", "P_SYS", c) Then
        c = " "
    End If
    YOIN_TBL = Split(Trim(c), ",", -1)
                                '基準日数
    If GetIni(App.EXEName, "DAYS", "P_SYS", c) Then
        REC_DAYS = 0
    Else
        If IsNumeric(Trim(c)) Then
            REC_DAYS = CInt(Trim(c))
        Else
            REC_DAYS = 0
        End If
    End If
                        
    
    
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
    
    
    
            
    '画面初期設定
    If Init_Proc() Then
        Unload Me
    End If

    If List_Disp_Proc() Then
        Unload Me
    End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '在庫移動歴Close
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
    
                                            '管理ﾏｽﾀClose
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "管理ﾏｽﾀ")
        End If
    End If
    
    
                                            '品目ﾏｽﾀClose
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目ﾏｽﾀ")
        End If
    End If
                                            '資材発注検討ﾌｧｲﾙClose
    sts = BTRV(BtOpClose, P_SHKENTO_POS, P_SHKENTO_REC, Len(P_SHKENTO_REC), K0_P_SHKENTO, Len(K0_P_SHKENTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材発注検討ﾌｧｲﾙ")
        End If
    End If
                                            '資材注文ﾃﾞｰﾀClose
    sts = BTRV(BtOpClose, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材注文ﾃﾞｰﾀ")
        End If
    End If
                                            '在庫ﾃﾞｰﾀClose
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫ﾃﾞｰﾀ")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PR000701 = Nothing
    Set PR000702 = Nothing


    End
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
    
    
    
    For i = ptxS_YMD To ptxTOTAL
        Text1(i).Text = ""
    Next i
    
    
    sts = BTRV(BtOpGetFirst, P_SHKENTO_POS, P_SHKENTO_REC, Len(P_SHKENTO_REC), K1_P_SHKENTO, Len(K1_P_SHKENTO), 1)
        
    Select Case sts
        Case BtNoErr
        
            Text1(ptxS_YMD).Text = Mid(StrConv(P_SHKENTO_REC.S_YMD, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SHKENTO_REC.S_YMD, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_SHKENTO_REC.S_YMD, vbUnicode), 7, 2)

            Text1(ptxE_YMD).Text = Mid(StrConv(P_SHKENTO_REC.E_YMD, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SHKENTO_REC.E_YMD, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_SHKENTO_REC.E_YMD, vbUnicode), 7, 2)


        Case BtErrEOF
        
            '処理年月日＝当日
            Text1(ptxS_YMD).Text = Format(DateAdd("m", -1, Format(Now, "YYYY/MM/DD")), "YYYY/MM/DD")
            Text1(ptxS_YMD).Text = Format(DateAdd("d", 1, Text1(ptxS_YMD).Text), "YYYY/MM/DD")
            
            
            Text1(ptxE_YMD).Text = Format(Now, "YYYY/MM/DD")
        
        
        Case Else
            Call File_Error(sts, BtOpGetFirst, "資材発注検討ﾌｧｲﾙ")
            Exit Function
    End Select
    
    
    
    

    
    
    
    'ｿｰﾄ情報の初期化
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0               'ﾃﾞﾌｫﾙﾄ昇順
    Next i

    Init_Proc = False

End Function



Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'           資材発注検討ﾌｧｲﾙの表示
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim Row             As Long


Dim TOTAL           As Long


    List_Disp_Proc = True
    PR000701.MousePointer = vbHourglass
        
    
    '-------------------------------------  '実績明細のｾｯﾄ
    Set SEISAN = Nothing
    
    Row = Min_Row - 1
    
    TOTAL = 0
    
    com = BtOpGetLast
    
    
    Do
    
        DoEvents
    
        sts = BTRV(com, P_SHKENTO_POS, P_SHKENTO_REC, Len(P_SHKENTO_REC), K1_P_SHKENTO, Len(K1_P_SHKENTO), 1)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "資材発注検討ﾌｧｲﾙ")
                Exit Function
        End Select
    
    
        Row = Row + 1
        If Grid_Set_Proc(Row) Then
            Exit Function
        End If
        
        TOTAL = TOTAL + CLng(StrConv(P_SHKENTO_REC.KINGAKU, vbUnicode))
        
        com = BtOpGetPrev
    
    
    
    Loop
    
    
    Text1(ptxTOTAL).Text = Format(TOTAL, "#,##0")
    
    
    Set TDBGrid1(pGridDETAIL).Array = SEISAN
    TDBGrid1(pGridDETAIL).ReBind
    TDBGrid1(pGridDETAIL).Update
    TDBGrid1(pGridDETAIL).MoveFirst
    
    
    PR000701.MousePointer = vbDefault
    
    
    List_Disp_Proc = False
    


End Function


Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           資材発注検討ﾌｧｲﾙの内容をｸﾞﾘｯﾄﾞにｾｯﾄする
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim i           As Integer
Dim j           As Integer


Dim wkAVE       As Long


    Grid_Set_Proc = True
    
    
    SEISAN.ReDim Min_Row, Row, Min_Col, Max_Col

    If Row = 1 Then
        TDBGrid1(pGridDETAIL).Columns(colJITU_QTY1).Caption = StrConv(P_SHKENTO_REC.JITU_TBL(2).JITU_YM, vbUnicode)
        TDBGrid1(pGridDETAIL).Columns(colJITU_QTY2).Caption = StrConv(P_SHKENTO_REC.JITU_TBL(1).JITU_YM, vbUnicode)
        TDBGrid1(pGridDETAIL).Columns(colJITU_QTY3).Caption = StrConv(P_SHKENTO_REC.JITU_TBL(0).JITU_YM, vbUnicode)
    End If
    '品目ｺｰﾄﾞ
    SEISAN(Row, colHIN_GAI) = StrConv(P_SHKENTO_REC.HIN_GAI, vbUnicode)
    '品名
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHKENTO_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHKENTO_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SHKENTO_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Select Case sts
        Case BtNoErr
        
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
            Exit Function
    End Select
    SEISAN(Row, colHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        
    '実績
    wkAVE = CLng(StrConv(P_SHKENTO_REC.JITU_TBL(0).JITU_QTY, vbUnicode)) + _
                CLng(StrConv(P_SHKENTO_REC.JITU_TBL(1).JITU_QTY, vbUnicode)) + _
                CLng(StrConv(P_SHKENTO_REC.JITU_TBL(2).JITU_QTY, vbUnicode))
    SEISAN(Row, colJITU_QTY1) = Format(CLng(StrConv(P_SHKENTO_REC.JITU_TBL(2).JITU_QTY, vbUnicode)), "#,##0")
    SEISAN(Row, colJITU_QTY2) = Format(CLng(StrConv(P_SHKENTO_REC.JITU_TBL(1).JITU_QTY, vbUnicode)), "#,##0")
    SEISAN(Row, colJITU_QTY3) = Format(CLng(StrConv(P_SHKENTO_REC.JITU_TBL(0).JITU_QTY, vbUnicode)), "#,##0")
    '平均
    SEISAN(Row, colJITU_AVE) = Format(Round(wkAVE / 3, 1), "#0.0")
    'LT
    If IsNumeric(StrConv(P_SHKENTO_REC.LT_DAYS, vbUnicode)) Then
        SEISAN(Row, colLT_DAYS) = Format(CInt(StrConv(P_SHKENTO_REC.LT_DAYS, vbUnicode)), "#0")
    Else
        SEISAN(Row, colLT_DAYS) = 0
    End If
    '収支
    SEISAN(Row, colSYUSHI_CODE) = StrConv(P_SHKENTO_REC.SYUSHI_CODE, vbUnicode)
    '基準在庫
    If IsNumeric(StrConv(P_SHKENTO_REC.ZAIKO_STANDARD, vbUnicode)) Then
        SEISAN(Row, colZAIKO_STANDARD) = Format(CLng(StrConv(P_SHKENTO_REC.ZAIKO_STANDARD, vbUnicode)), "#,##0")
    Else
        SEISAN(Row, colZAIKO_STANDARD) = 0
    End If
    '理論在庫
    SEISAN(Row, colZAIKO_QTY) = Format(CLng(StrConv(P_SHKENTO_REC.ZAIKO_QTY, vbUnicode)), "#,##0")
    'ﾛｯﾄ
    If IsNumeric(StrConv(P_SHKENTO_REC.LOT, vbUnicode)) Then
        SEISAN(Row, colLOT) = Format(CLng(StrConv(P_SHKENTO_REC.LOT, vbUnicode)), "#,##0")
    Else
        SEISAN(Row, colLOT) = 0
    End If
    '発注先
    SEISAN(Row, colORDER_CODE) = StrConv(P_SHKENTO_REC.ORDER_CODE, vbUnicode)
    '発注残数量
    SEISAN(Row, colSHIJI_Z_QTY) = Format(CLng(StrConv(P_SHKENTO_REC.SHIJI_Z_QTY, vbUnicode)), "#,##0")
    '論理　発注数
    SEISAN(Row, colSHIJI_QTY_R) = Format(CLng(StrConv(P_SHKENTO_REC.SHIJI_QTY_R, vbUnicode)), "#,##0")
    '確定　発注数
    SEISAN(Row, colSHIJI_QTY_K) = Format(CLng(StrConv(P_SHKENTO_REC.SHIJI_QTY_K, vbUnicode)), "#,##0")
    '単価
    SEISAN(Row, colTANKA) = Format(CDbl(StrConv(P_SHKENTO_REC.TANKA, vbUnicode)), "#,##0.00")
    '金額
    SEISAN(Row, colKINGAKU) = Format(CLng(StrConv(P_SHKENTO_REC.KINGAKU, vbUnicode)), "#,##0")
    
    Grid_Set_Proc = False

End Function






Private Function SUM_Make_Proc(Data_Flg As Boolean) As Integer
'----------------------------------------------------------------------------
'                   資材発注検討ﾃﾞｰﾀ作成
'----------------------------------------------------------------------------
Dim sts                     As Integer
Dim Main_com                As Integer

Dim com                     As Integer

Dim SKIP_Flg                As Boolean
    
    
Dim i                       As Integer
    


Dim S_YMD(0 To 2)           As String
Dim E_YMD(0 To 2)           As String

Dim JITU_QTY(0 To 2)        As Long

Dim Sumi_Zaiko_Qty          As Long
Dim Mi_Zaiko_Qty            As Long

Dim SHIJI_QTY               As Long

Dim JITU_AVE                As Double


Dim ZAIKO_STANDARD          As Long
    
    
    SUM_Make_Proc = True
    PR000701.MousePointer = vbHourglass


    
    
    For i = 0 To 2
        
        S_YMD(i) = Format(DateAdd("m", (i + 1) * -1, Text1(ptxS_YMD).Text), "YYYY/MM/DD")
        E_YMD(i) = Format(DateAdd("m", (i + 1) * -1, Text1(ptxE_YMD).Text), "YYYY/MM/DD")
    
    
    Next i


    



    '-----------------------------------------  集計ﾃﾞｰﾀ全件削除


    Main_com = BtOpGetFirst



    Do
    
        DoEvents
        
        sts = BTRV(Main_com, P_SHKENTO_POS, P_SHKENTO_REC, Len(P_SHKENTO_REC), K0_P_SHKENTO, Len(K0_P_SHKENTO), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
            
                Exit Do
            
            
            Case Else
                Call File_Error(sts, Main_com, "資材発注検討ﾃﾞｰﾀ")
                Exit Function
        End Select

        sts = BTRV(BtOpDelete, P_SHKENTO_POS, P_SHKENTO_REC, Len(P_SHKENTO_REC), K0_P_SHKENTO, Len(K0_P_SHKENTO), 0)
            
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "生産実績明細集計ﾃﾞｰﾀ")
        End Select

    
        Main_com = BtOpGetNext
    
    Loop


    '-----------------------------------------  集計開始

    Main_com = BtOpGetGreaterEqual

    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, "")

    Do

        DoEvents
        
        sts = BTRV(Main_com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            
        Select Case sts
            Case BtNoErr
            
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> SHIZAI Then
                    Exit Do
                End If
            
                If StrConv(ITEMREC.NAIGAI, vbUnicode) <> NAIGAI_NAI Then
                    Exit Do
                End If
            
            Case BtErrEOF
            
                Exit Do
            
            
            Case Else
                Call File_Error(sts, Main_com, "品目ﾏｽﾀ")
                Exit Function
        End Select

        Data_Flg = True


        '移動歴 実績集計

        com = BtOpGetGreaterEqual

        Call UniCode_Conv(K1_IDO.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K1_IDO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K1_IDO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
        Call UniCode_Conv(K1_IDO.JITU_DT, Format(S_YMD(2), "YYYYMMDD"))
        Call UniCode_Conv(K1_IDO.JITU_TM, "")

        For i = 0 To 2
            JITU_QTY(i) = 0
        Next i

        Do
            
            DoEvents
            
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
                
                
            Select Case sts
                Case BtNoErr
                
                    If StrConv(IDOREC.JGYOBU, vbUnicode) <> StrConv(ITEMREC.JGYOBU, vbUnicode) Then
                        Exit Do
                    End If
                
                    If StrConv(IDOREC.NAIGAI, vbUnicode) <> StrConv(ITEMREC.NAIGAI, vbUnicode) Then
                        Exit Do
                    End If
                
                    If StrConv(IDOREC.HIN_GAI, vbUnicode) <> StrConv(ITEMREC.HIN_GAI, vbUnicode) Then
                        Exit Do
                    End If
                
                
                    If StrConv(IDOREC.JITU_DT, vbUnicode) > Format(E_YMD(0), "YYYYMMDD") Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                
                    Exit Do
                
                
                Case Else
                    Call File_Error(sts, com, "在庫移動歴")
                    Exit Function
            
            End Select



            For i = 0 To UBound(YOIN_TBL)
            
                If StrConv(IDOREC.RIRK_ID, vbUnicode) = YOIN_TBL(i) Then
                
                    Select Case StrConv(IDOREC.JITU_DT, vbUnicode)
                    
                        Case Format(S_YMD(0), "YYYYMMDD") To Format(E_YMD(0), "YYYYMMDD")
                            JITU_QTY(0) = JITU_QTY(0) + (CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)))
                        Case Format(S_YMD(1), "YYYYMMDD") To Format(E_YMD(1), "YYYYMMDD")
                            JITU_QTY(1) = JITU_QTY(1) + (CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)))
                        Case Format(S_YMD(2), "YYYYMMDD") To Format(E_YMD(2), "YYYYMMDD")
                            JITU_QTY(2) = JITU_QTY(2) + (CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)))
                    End Select
                
                    Exit For
                
                End If
            
            Next i

            com = BtOpGetNext
        
        Loop


        '現在庫集計
        If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, _
                                Mi_Zaiko_Qty, _
                                StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
            Exit Function
        
        End If
        
        '注文残集計
        
        com = BtOpGetGreaterEqual
        
        Call UniCode_Conv(K1_P_SHORDER.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K1_P_SHORDER.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K1_P_SHORDER.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
        Call UniCode_Conv(K1_P_SHORDER.ORDER_DT, "")
        Call UniCode_Conv(K1_P_SHORDER.ORDER_NO, "")


        SHIJI_QTY = 0

        Do

            DoEvents
        
            sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K1_P_SHORDER, Len(K1_P_SHORDER), 1)
            
            Select Case sts
                Case BtNoErr
            
                    If StrConv(P_SHORDER_REC.JGYOBU, vbUnicode) <> StrConv(ITEMREC.JGYOBU, vbUnicode) Then
                        Exit Do
                    End If
        
                    If StrConv(P_SHORDER_REC.NAIGAI, vbUnicode) <> StrConv(ITEMREC.NAIGAI, vbUnicode) Then
                        Exit Do
                    End If
        
        
                    If StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode) <> StrConv(ITEMREC.HIN_GAI, vbUnicode) Then
                        Exit Do
                    End If
        
                Case BtErrEOF
                
                    Exit Do
                
                
                Case Else
                    Call File_Error(sts, com, "資材注文ﾃﾞｰﾀ")
                    Exit Function
            End Select

            If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = P_KAN_OFF Then
                SHIJI_QTY = SHIJI_QTY + (CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) - CLng(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)))
            End If
            
            com = BtOpGetNext

        Loop

        'ﾚｺｰﾄﾞｾｯﾄ
        
        If JITU_QTY(0) = 0 And _
            JITU_QTY(1) = 0 And _
            JITU_QTY(2) = 0 And _
            Sumi_Zaiko_Qty = 0 And _
            Mi_Zaiko_Qty = 0 And _
            SHIJI_QTY = 0 Then
        Else
        
                                                                        '事業部
            Call UniCode_Conv(P_SHKENTO_REC.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                                                                        '国内外
            Call UniCode_Conv(P_SHKENTO_REC.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                                                                        '品番(外部)
            Call UniCode_Conv(P_SHKENTO_REC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                                                        '実績
            For i = 0 To 2
                
                Call UniCode_Conv(P_SHKENTO_REC.JITU_TBL(i).JITU_YM, Left(E_YMD(i), 7))
                Call UniCode_Conv(P_SHKENTO_REC.JITU_TBL(i).JITU_QTY, Format(JITU_QTY(i), "00000000"))
            Next i
                                                                        'LT CODE(未使用)
            Call UniCode_Conv(P_SHKENTO_REC.LT_CODE, "")
                                                                        'LT
            Call UniCode_Conv(P_SHKENTO_REC.LT_DAYS, StrConv(ITEMREC.G_SHIIRE_TBL(0).LEAD_TIME, vbUnicode))
                                                                        '収支単位
            Call UniCode_Conv(P_SHKENTO_REC.SYUSHI_CODE, StrConv(ITEMREC.G_SYUSHI, vbUnicode))
                                                                        
                                                                        
                                                                        
                                                                        
                                                                        '理論在庫
            Call UniCode_Conv(P_SHKENTO_REC.ZAIKO_QTY, Format(Sumi_Zaiko_Qty + Mi_Zaiko_Qty, "00000000"))
                                                                        'LOT
            Call UniCode_Conv(P_SHKENTO_REC.LOT, StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode))
                                                                        '発注先
            Call UniCode_Conv(P_SHKENTO_REC.ORDER_CODE, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
                                                                        '発注残数量
            Call UniCode_Conv(P_SHKENTO_REC.SHIJI_Z_QTY, Format(SHIJI_QTY, "00000000"))
                                                                        '発注残ｺｰﾄﾞ(未使用)
            Call UniCode_Conv(P_SHKENTO_REC.SHIJI_Z_CODE, "")
                                                                        
                                                                        
                                                                        
                                                                        '発注数　確定
            Call UniCode_Conv(P_SHKENTO_REC.SHIJI_QTY_K, "00000000")
                                                                        '発注数　ｺｰﾄﾞ(未使用)
            Call UniCode_Conv(P_SHKENTO_REC.SHIJI_CODE, "")
                                                                        
                                                                        '単価
            If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)) Then
                Call UniCode_Conv(P_SHKENTO_REC.TANKA, StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))
            Else
                Call UniCode_Conv(P_SHKENTO_REC.TANKA, "00000000.00")
            End If
                                                                        
                                                                        '金額
            If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)) Then
                Call UniCode_Conv(P_SHKENTO_REC.KINGAKU, Format(Round(SHIJI_QTY * CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)), 0), "0000000000"))
            Else
                Call UniCode_Conv(P_SHKENTO_REC.KINGAKU, "0000000000")
            End If
                                                                        
                                                                                
                                                                        
                                                                        
                                                                        'SORT KEY
            JITU_AVE = Round((JITU_QTY(0) + JITU_QTY(1) + JITU_QTY(2)) / 3, 1)
            
            Call UniCode_Conv(P_SHKENTO_REC.SORT_KEY, Format(JITU_AVE, "00000000.0"))
    
    
                                                                        '基準在庫
            If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).LEAD_TIME, vbUnicode)) Then
                ZAIKO_STANDARD = Round((JITU_AVE / REC_DAYS) * CInt(StrConv(ITEMREC.G_SHIIRE_TBL(0).LEAD_TIME, vbUnicode)), 0)
                Call UniCode_Conv(P_SHKENTO_REC.ZAIKO_STANDARD, Format(ZAIKO_STANDARD, "00000000"))
            Else
                Call UniCode_Conv(P_SHKENTO_REC.ZAIKO_STANDARD, "00000000")
            End If
    
            If (Sumi_Zaiko_Qty + Mi_Zaiko_Qty) + SHIJI_QTY >= ZAIKO_STANDARD Then
                                                                        '発注数　理論
                Call UniCode_Conv(P_SHKENTO_REC.SHIJI_QTY_R, "00000000")
            Else
                If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode)) Then
                    Call UniCode_Conv(P_SHKENTO_REC.SHIJI_QTY_R, StrConv(ITEMREC.G_SHIIRE_TBL(0).LOT, vbUnicode))
                Else
                    Call UniCode_Conv(P_SHKENTO_REC.SHIJI_QTY_R, "00000000")
                End If
            End If
    
    
    
    
            Call UniCode_Conv(P_SHKENTO_REC.S_YMD, Format(Text1(ptxS_YMD).Text, "YYYYMMDD"))
            Call UniCode_Conv(P_SHKENTO_REC.E_YMD, Format(Text1(ptxE_YMD).Text, "YYYYMMDD"))
    
    
            Call UniCode_Conv(P_SHKENTO_REC.FILLER, "")
    
    
            sts = BTRV(BtOpInsert, P_SHKENTO_POS, P_SHKENTO_REC, Len(P_SHKENTO_REC), K0_P_SHKENTO, Len(K0_P_SHKENTO), 0)
            Select Case sts
                Case BtNoErr
                Case Else
                    Call File_Error(sts, BtOpInsert, "資材発注検討ﾌｧｲﾙ")
                    Exit Function
            End Select

        End If
        
        Main_com = BtOpGetNext

    Loop

    SUM_Make_Proc = False
    PR000701.MousePointer = vbDefault


End Function



Private Sub RE_SUM_PROC()
'----------------------------------------------------------------------------
'                   表示内容の再計算
'----------------------------------------------------------------------------
Dim i       As Integer
Dim TOTAL   As Long

    PR000701.MousePointer = vbHourglass

    TOTAL = 0

    For i = 1 To SEISAN.UpperBound(1)
    
        SEISAN(i, colKINGAKU) = Format(Round((CLng(SEISAN(i, colSHIJI_Z_QTY)) + CLng(SEISAN(i, colSHIJI_QTY_K))) * CDbl(SEISAN(i, colTANKA)), 0), "#,##0")

        TOTAL = TOTAL + CLng(SEISAN(i, colKINGAKU))
    
    Next i

    Text1(ptxTOTAL).Text = Format(TOTAL, "#,##0")


    Set TDBGrid1(pGridDETAIL).Array = SEISAN
    
    TDBGrid1(pGridDETAIL).ReBind
    TDBGrid1(pGridDETAIL).Update
    TDBGrid1(pGridDETAIL).MoveFirst


    PR000701.MousePointer = vbDefault

End Sub






Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   ﾃﾞｰﾀ出力処理
'----------------------------------------------------------------------------
Dim i           As Integer
Dim sts         As Integer
Dim SHIJI_QTY   As Long


    Update_Proc = True

    PR000701.MousePointer = vbHourglass

    For i = 1 To SEISAN.UpperBound(1)
    
        Call UniCode_Conv(K0_P_SHKENTO.JGYOBU, SHIZAI)
        Call UniCode_Conv(K0_P_SHKENTO.NAIGAI, NAIGAI_NAI)
        Call UniCode_Conv(K0_P_SHKENTO.HIN_GAI, SEISAN(i, colHIN_GAI))
    
    
        sts = BTRV(BtOpGetEqual, P_SHKENTO_POS, P_SHKENTO_REC, Len(P_SHKENTO_REC), K0_P_SHKENTO, Len(K0_P_SHKENTO), 0)
            
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "資材発注検討ﾌｧｲﾙ")
                Exit Function
        End Select
    
    
        If sts = BtNoErr Then
        
            If IsNumeric(SEISAN(i, colSHIJI_QTY_R)) Then
                Call UniCode_Conv(P_SHKENTO_REC.SHIJI_QTY_R, Format(CLng(SEISAN(i, colSHIJI_QTY_R)), "0000000"))
            Else
                Call UniCode_Conv(P_SHKENTO_REC.SHIJI_QTY_R, "00000000")
            End If
        
            If IsNumeric(SEISAN(i, colSHIJI_QTY_K)) Then
                Call UniCode_Conv(P_SHKENTO_REC.SHIJI_QTY_K, Format(CLng(SEISAN(i, colSHIJI_QTY_K)), "0000000"))
            Else
                Call UniCode_Conv(P_SHKENTO_REC.SHIJI_QTY_K, "00000000")
            End If
        
            If IsNumeric(SEISAN(i, colTANKA)) Then
                Call UniCode_Conv(P_SHKENTO_REC.TANKA, Format(CDbl(SEISAN(i, colTANKA)), "00000000.00"))
            Else
                Call UniCode_Conv(P_SHKENTO_REC.TANKA, "000000000.00")
            End If
        
        
        
                                                                        '金額
            
            SHIJI_QTY = CLng(StrConv(P_SHKENTO_REC.SHIJI_Z_QTY, vbUnicode)) + CLng(StrConv(P_SHKENTO_REC.SHIJI_QTY_K, vbUnicode))
            
            If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)) Then
                Call UniCode_Conv(P_SHKENTO_REC.KINGAKU, Format(Round(SHIJI_QTY * _
                                                                CDbl(StrConv(P_SHKENTO_REC.TANKA, vbUnicode)), 0), "0000000000"))
            Else
                Call UniCode_Conv(P_SHKENTO_REC.KINGAKU, "0000000000")
            End If
        
            sts = BTRV(BtOpUpdate, P_SHKENTO_POS, P_SHKENTO_REC, Len(P_SHKENTO_REC), K0_P_SHKENTO, Len(K0_P_SHKENTO), 0)
                
            Select Case sts
                Case BtNoErr
                Case Else
                    Call File_Error(sts, BtOpUpdate, "資材発注検討ﾌｧｲﾙ")
                    Exit Function
            End Select
        
        
        End If
        
        
        
        
        
        
        
        
    
    
    
    Next i



    PR000701.MousePointer = vbDefault
    Update_Proc = False
End Function

