VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PCB00101 
   Caption         =   "PCB.U　設変管理台帳　変換"
   ClientHeight    =   10875
   ClientLeft      =   2025
   ClientTop       =   -3210
   ClientWidth     =   17640
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
   OLEDropMode     =   1  '手動
   ScaleHeight     =   10875
   ScaleWidth      =   17640
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   372
      Left            =   2040
      OLEDragMode     =   1  '自動
      OLEDropMode     =   1  '手動
      TabIndex        =   5
      Top             =   600
      Width           =   8772
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   8895
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   15690
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "管理№"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "日　付"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "設計変更№"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "ｻｰﾋﾞｽ品番"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "工場品番"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "⇒"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "ｻｰﾋﾞｽ品番"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "工場品番"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "設変実施"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "変更部品"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "変更内容(変更/追加)"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "交換場所"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "設変原紙保管"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "備       考１"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "備       考２"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "備       考３"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "備       考４"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   17
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   714
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=17"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=794"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=688"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2011"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1905"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1826"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1720"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2381"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2275"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=1588"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1482"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=609"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=503"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=2831"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2725"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=1958"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=1852"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=979"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=873"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=1958"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=1852"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(41)=   "Column(10).Width=3334"
      Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=3228"
      Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(45)=   "Column(11).Width=2593"
      Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=2487"
      Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(49)=   "Column(12).Width=3281"
      Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=3175"
      Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(53)=   "Column(13).Width=12250"
      Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=12144"
      Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(57)=   "Column(14).Width=3281"
      Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=3175"
      Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(61)=   "Column(15).Width=3281"
      Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=3175"
      Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(65)=   "Column(16).Width=3281"
      Splits(0)._ColumnProps(66)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(16)._WidthInPix=3175"
      Splits(0)._ColumnProps(68)=   "Column(16).Order=17"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   0
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      OLEDropMode     =   1
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&HFFFF00&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(27)  =   ":id=14,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(28)  =   ":id=14,.fontname=ＭＳ ゴシック"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFF00&"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=110,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=107,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=108,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=109,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=3"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=86,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=83,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=84,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=85,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=66,.parent=13,.alignment=3"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=70,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=67,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=68,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=69,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=74,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=71,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=72,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=73,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(12).Style:id=78,.parent=13"
      _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=75,.parent=14"
      _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=76,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=77,.parent=17"
      _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=82,.parent=13"
      _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=79,.parent=14"
      _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=80,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=81,.parent=17"
      _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=90,.parent=13"
      _StyleDefs(95)  =   "Splits(0).Columns(14).HeadingStyle:id=87,.parent=14"
      _StyleDefs(96)  =   "Splits(0).Columns(14).FooterStyle:id=88,.parent=15"
      _StyleDefs(97)  =   "Splits(0).Columns(14).EditorStyle:id=89,.parent=17"
      _StyleDefs(98)  =   "Splits(0).Columns(15).Style:id=94,.parent=13"
      _StyleDefs(99)  =   "Splits(0).Columns(15).HeadingStyle:id=91,.parent=14"
      _StyleDefs(100) =   "Splits(0).Columns(15).FooterStyle:id=92,.parent=15"
      _StyleDefs(101) =   "Splits(0).Columns(15).EditorStyle:id=93,.parent=17"
      _StyleDefs(102) =   "Splits(0).Columns(16).Style:id=98,.parent=13"
      _StyleDefs(103) =   "Splits(0).Columns(16).HeadingStyle:id=95,.parent=14"
      _StyleDefs(104) =   "Splits(0).Columns(16).FooterStyle:id=96,.parent=15"
      _StyleDefs(105) =   "Splits(0).Columns(16).EditorStyle:id=97,.parent=17"
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
      TabIndex        =   2
      ToolTipText     =   "商品化構成を保存します"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登 録"
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
      Index           =   1
      Left            =   1845
      TabIndex        =   1
      ToolTipText     =   "商品化構成を読み込みます（Ｆ5）"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "読 込"
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
      TabIndex        =   0
      Top             =   0
      Width           =   1170
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   360
      TabIndex        =   8
      Top             =   10560
      Width           =   180
   End
   Begin VB.Label lblDisp_Count 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   13440
      TabIndex        =   7
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "件数"
      Height          =   252
      Index           =   1
      Left            =   12840
      TabIndex        =   6
      Top             =   720
      Width           =   612
   End
   Begin VB.Label Label1 
      Caption         =   "ファイル名"
      Height          =   252
      Index           =   0
      Left            =   720
      TabIndex        =   4
      Top             =   720
      Width           =   1332
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "PCB00101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim DEF_NAIGAI          As String * 1



'--------------------------------------------<ｸﾞﾘｯﾄﾞ>
Dim PCB_U      As New XArrayDB

Private Const Min_Row% = 1              '最小行数

Dim Max_Row    As Integer               'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 16             '最大列数

Private Const colKANRI_NO% = 0          '管理№
Private Const colEX_DATE% = 1           '日付
Private Const colSETUHEN_NO% = 2        '設変管理№
Private Const colBEF_HIN_GAI% = 3       '変更前　ｻｰﾋﾞｽ品番
Private Const colBEF_HIN_NAI% = 4       '変更前　工場品番

Private Const colDummy% = 5             '⇒

Private Const colAFT_HIN_GAI% = 6       '変更後　ｻｰﾋﾞｽ品番
Private Const colAFT_HIN_NAI% = 7       '変更後　工場品番

Private Const colSETUHEN_JITSU% = 8     '設変実施


Private Const colHEN_BUHIN% = 9         '変更部品
Private Const colHEN_NAIYO% = 10        '変更内容
Private Const colHEN_BASHO% = 11        '交換場所
Private Const colSETUHEN_HOKAN% = 12    '設変原紙保管
Private Const colBIKOU1% = 13           '備考1
Private Const colBIKOU2% = 14           '備考2
Private Const colBIKOU3% = 15           '備考3
Private Const colBIKOU4% = 16           '備考4

'--------------------------------------------<EXCEL>
Private Const selKANRI_NO% = 2          '管理№
Private Const selEX_DATE% = 3           '日付
Private Const selSETUHEN_NO% = 4        '設変管理№
Private Const selBEF_HIN_GAI% = 5       '変更前　ｻｰﾋﾞｽ品番
Private Const selBEF_HIN_NAI% = 6       '変更前　工場品番

Private Const selDummy% = 7             '⇒

Private Const selAFT_HIN_GAI% = 8       '変更後　ｻｰﾋﾞｽ品番
Private Const selAFT_HIN_NAI% = 9       '変更後　工場品番

Private Const selSETUHEN_JITSU% = 10        '変更部品


Private Const selHEN_BUHIN% = 11        '変更部品
Private Const selHEN_NAIYO% = 12        '変更内容
Private Const selHEN_BASHO% = 15        '交換場所
Private Const selSETUHEN_HOKAN% = 16    '設変原紙保管
Private Const selBIKOU1% = 17            '備考1
Private Const selBIKOU2% = 18            '備考2
Private Const selBIKOU3% = 19            '備考3
Private Const selBIKOU4% = 20            '備考4

    
        





Private Const LAST_UPDATE_DAY$ = "[PCB0010] 2014.06.27 16:00"

Private Sub Command1_Click(Index As Integer)

Dim sWk         As String
Dim i           As Long
Dim j           As Long


    Select Case Index



        Case 0          '読込み

            'ﾌｧｲﾙ名ﾁｪｯｸ
''            sWk = Text1.Text
''            For i = Len(sWk) To 1 Step -1
''                If Mid(sWk, i, 1) = "\" Then
''                    Exit For
''                End If
''            Next i
''
''            i = i + 1
''
''            If Mid(sWk, i, 9) <> "C_22T826_" Then
''                MsgBox "エラー" & vbCrLf & vbCrLf & "ファイル名が違います。", vbExclamation
''                Exit Sub
''            End If
''
''            j = InStr(i, sWk, ".")
''            If StrConv(Right(sWk, Len(sWk) - j), vbLowerCase) <> "txt" Then
''                MsgBox "エラー" & vbCrLf & vbCrLf & "拡張子が「txt」以外です。", vbExclamation
''                Exit Sub
''            End If



            If Trim(Text1.Text) = "" Then
                Exit Sub
            End If


            '取込みﾃﾞｰﾀ表示
            If List_Disp_Proc() Then
                Unload Me
            End If


            If PCB_U.Count(1) > 0 Then
                Command1(1).Enabled = True
            End If


        Case 1          '登録





            If Update_Proc() Then
                Unload Me
            End If



        Case 2          '終了

            Unload Me
    End Select



    Command1(Index).SetFocus


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
    
    
'    If Shift = vbAltMask Then
'
'        If TDBGrid1.AllowUpdate Then
'
'            TDBGrid1.AllowUpdate = False
'            TDBGrid1.AllowAddNew = False
'            TDBGrid1.AllowDelete = False
'
'
'            TDBGrid1.Columns(colTEI_LABELID).Visible = False
'            TDBGrid1.Columns(colHAKO_NO).Visible = False
'
'            TDBGrid1.Columns(colTEI_LABELID).Locked = True
'            TDBGrid1.Columns(colHAKO_NO).Locked = True
'
'
'
'        Else
'
'
'            TDBGrid1.AllowUpdate = True
'            TDBGrid1.AllowAddNew = True
'            TDBGrid1.AllowDelete = True
'
'
'            TDBGrid1.Columns(colTEI_LABELID).Visible = True
'            TDBGrid1.Columns(colHAKO_NO).Visible = True
'
'
'            TDBGrid1.Columns(colTEI_LABELID).Locked = False
'            TDBGrid1.Columns(colHAKO_NO).Locked = False
'
'        End If
'
'    End If
    
    
    Select Case KeyCode
        Case vbKeyF12
            Unload Me
    End Select
    
    
    

End Sub

Private Sub Form_Load()


Dim c       As String * 128
Dim i       As Integer


    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "PCB.U設変管理台帳　変換", Me.hwnd, 0)
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


    If GetIni(App.EXEName, "NAIGAI", App.EXEName, c) Then
        Beep
        MsgBox "国内外[NAIGAI]の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    DEF_NAIGAI = Trim(c)
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            PCB00101.Caption = "PCB.U　設変管理台帳　変換(" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)

'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i

    Unload SubMenu(i)



    PCB00101.Caption = PCB00101.Caption & " " & LAST_UPDATE_DAY

                                'モジュール品目マスタ ＯＰＥＮ
    If M_ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If


                                'PCB.U設変 ＯＰＥＮ
    If PCB_U_Open(BtOpenNomal) Then
        Unload Me
    End If



End Sub


Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.Text = Trim(Data.Files(1))

'    Text1.Text = Data.GetData(vbCFText)

    Command1(0).Value = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
Dim sts     As Integer
    
    
    sts = BTRV(BtOpClose, M_ITEM_POS, M_ITEM_REC, Len(M_ITEM_REC), K0_M_ITEM, Len(K0_M_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "モジュール品目マスタ")
        End If
    End If
    
    
    sts = BTRV(BtOpClose, PCB_U_POS, PCB_U_REC, Len(PCB_U_REC), K0_PCB_U, Len(K0_PCB_U), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "PCB.U設変管理台帳")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, PCB_U_POS, PCB_U_REC, Len(PCB_U_REC), K0_PCB_U, Len(K0_PCB_U), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If


    Set PCB00101 = Nothing



    End

End Sub

Private Sub SHORI_Click(Index As Integer)

    Select Case Index
    
        Case 0
            Command1(0).Value = True
        Case 1
            Command1(1).Value = True
        Case 2
            Command1(2).Value = True
    End Select



End Sub

Private Sub SubMenu_Click(Index As Integer)
    
Dim i   As Integer
    
    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    PCB00101.Caption = "PCB.U　設変管理台帳　変換（" + RTrim(JGYOBU_T(Index).NAME) + "）"
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub

Private Sub TDBGrid1_OLEDragDrop(ByVal Data As TrueDBGrid80.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Text1.Text = Trim(Data.Files(0))
'    Text1.Text = Data.GetData(0)


    Command1(0).Value = True


End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Text1.Text = Trim(Data.Files(1))
    
    Command1(0).Value = True


'    If Data.GetFormat(vbCFText) Then
'        Text1.Text = Data.GetData(vbCFText)
'        Command1(0).Value = True
'    End If

End Sub

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   「PCB.U設変」登録処理
'----------------------------------------------------------------------------

Dim sts             As Integer
Dim ans             As Integer
    
Dim Upd_Com         As Integer
    
Dim INS_NOW         As String * 14
    
Dim c               As String * 128

Dim Row             As Long

Dim com             As Integer


    If PCB_U.Count(1) = 0 Then
        Exit Function
    End If



    Update_Proc = True
    
    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "PCB.U設変管理台帳変換　[登録]処理開始！！", Me.hwnd, 0)


'    On Error Resume Next
'
'    sts = GetIni("FILE", PCB_U_ID, "SYS", c)
'    If sts <> False Then
'        Call LOG_OUT(LOG_F, "SYS.INI [PCB_U]読み込みエラー ")
'        Exit Function
'    End If
'
'    Kill RTrim(c)
'
'
'    On Error GoTo 0

                                    

'                                'PCB.U設変 ＯＰＥＮ
'    If PCB_U_Open(BtOpenNomal) Then
'        Unload Me
'    End If
                                    
                                    
    com = BtOpGetFirst
    Do
    
        DoEvents
        
        sts = BTRV(com, PCB_U_POS, PCB_U_REC, Len(PCB_U_REC), K0_PCB_U, Len(K0_PCB_U), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
            
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "PCB.U設変")
                Call Input_UnLock
                Exit Function
        End Select
    
    
        sts = BTRV(BtOpDelete, PCB_U_POS, PCB_U_REC, Len(PCB_U_REC), K0_PCB_U, Len(K0_PCB_U), 0)
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "PCB.U設変")
                Call Input_UnLock
                Exit Function
        End Select
            
        com = BtOpGetNext
    
    Loop
                                    
                                    
                                    
    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")
                                    
                                    'テーブルリセット
    
    For Row = 1 To PCB_U.UpperBound(1)
        
        
        DoEvents
        
'-------------------------------------  <モジュール品目マスタ処理>
        Call UniCode_Conv(K0_M_ITEM.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_M_ITEM.NAIGAI, DEF_NAIGAI)
        Call UniCode_Conv(K0_M_ITEM.HIN_GAI, PCB_U(Row, colAFT_HIN_GAI))
        
        
        sts = BTRV(BtOpGetEqual, M_ITEM_POS, M_ITEM_REC, Len(M_ITEM_REC), K0_M_ITEM, Len(K0_M_ITEM), 0)
        Select Case sts
            Case BtNoErr
            
                Upd_Com = BtOpUpdate
            Case BtErrKeyNotFound
                Upd_Com = BtOpInsert
            
            Case Else
                Call File_Error(sts, BtOpInsert, "モジュール品目マスタ")
                Call Input_UnLock
                Exit Function
        End Select
        
        
        
        
        
            
        If Upd_Com = BtOpInsert Then
            Call UniCode_Conv(M_ITEM_REC.JGYOBU, Last_JGYOBU)                        '事業部
            Call UniCode_Conv(M_ITEM_REC.NAIGAI, DEF_NAIGAI)                        '国内外
            Call UniCode_Conv(M_ITEM_REC.HIN_GAI, PCB_U(Row, colAFT_HIN_GAI))       '品番（外部）

            Call UniCode_Conv(M_ITEM_REC.MODULE_KBN, "")                            'モジュール対象区分
            Call UniCode_Conv(M_ITEM_REC.MODULE_UNIT_KBN, "")                       'モジュールユニット区分
            
            Call UniCode_Conv(M_ITEM_REC.KENSA_JIGU, "")                            '検査治具
            
            Call UniCode_Conv(M_ITEM_REC.SETUHEN_KBN, "")                           '設計変更対象区分
            
            Call UniCode_Conv(M_ITEM_REC.SENDO_LAST_DATE, "")                       '鮮度管理最終日
            
            Call UniCode_Conv(M_ITEM_REC.HITUYO_SU, "")                             '必要数　個
            Call UniCode_Conv(M_ITEM_REC.HITUYO_TUKI, "")                           '必要数　月

        
        
            Call UniCode_Conv(M_ITEM_REC.FILLER, "")
        
            Call UniCode_Conv(M_ITEM_REC.INS_TANTO, App.EXEName)                    '追加担当者
            Call UniCode_Conv(M_ITEM_REC.Ins_DateTime, INS_NOW)                     '追加日時
            Call UniCode_Conv(M_ITEM_REC.INS_PROG_ID, App.EXEName)                  '追加プログラムID
        
        
        Else
        
            Call UniCode_Conv(M_ITEM_REC.UPD_TANTO, App.EXEName)                    '更新担当者
            Call UniCode_Conv(M_ITEM_REC.UPD_DATETIME, INS_NOW)                     '更新日時
            Call UniCode_Conv(M_ITEM_REC.UPD_PROG_ID, App.EXEName)                  '更新プログラムID
        
        End If
            
            
        Call UniCode_Conv(M_ITEM_REC.SETUHEN_LAST_DATE, Format(PCB_U(Row, colEX_DATE), "YYYYMMDD"))    '設計変更最終日
                    
        Do
            sts = BTRV(Upd_Com, M_ITEM_POS, M_ITEM_REC, Len(M_ITEM_REC), K0_M_ITEM, Len(K0_M_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    
                    Beep
                    ans = MsgBox("「モジュール品目マスタ」他端末でデータ使用中です。<M_ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Call Input_UnLock
                        Exit Function
                    End If
                
                Case Else
                    Call Input_UnLock
                    Call File_Error(sts, Upd_Com, "モジュール品目マスタ")
                    Exit Function
            End Select
            
        Loop
            
        
'-------------------------------------  <PCB.U設変処理>
        Call UniCode_Conv(PCB_U_REC.JGYOBU, Last_JGYOBU)                         '事業部
        Call UniCode_Conv(PCB_U_REC.NAIGAI, DEF_NAIGAI)                         '国内外
        Call UniCode_Conv(PCB_U_REC.HIN_GAI, PCB_U(Row, colAFT_HIN_GAI))        '品番（外部）
        
        Call UniCode_Conv(PCB_U_REC.KANRI_NO, PCB_U(Row, colKANRI_NO))          '管理№
        Call UniCode_Conv(PCB_U_REC.EX_DATE, Format(PCB_U(Row, colEX_DATE), "YYYYMMDD"))            '日付
        Call UniCode_Conv(PCB_U_REC.SETUHEN_NO, PCB_U(Row, colSETUHEN_NO))      '設計変更№
        
        Call UniCode_Conv(PCB_U_REC.BEF_HIN_GAI, PCB_U(Row, colBEF_HIN_GAI))    '変更前　ｻｰﾋﾞｽ品番
        Call UniCode_Conv(PCB_U_REC.BEF_HIN_NAI, PCB_U(Row, colBEF_HIN_NAI))    '変更前　工場品番

        Call UniCode_Conv(PCB_U_REC.AFT_HIN_GAI, PCB_U(Row, colAFT_HIN_GAI))    '変更後　ｻｰﾋﾞｽ品番
        Call UniCode_Conv(PCB_U_REC.AFT_HIN_NAI, PCB_U(Row, colAFT_HIN_NAI))    '変更後  工場品番

        Call UniCode_Conv(PCB_U_REC.SETUHEN_JITSU, PCB_U(Row, colSETUHEN_JITSU))    '設変実施


        Call UniCode_Conv(PCB_U_REC.HEN_BUHIN, PCB_U(Row, colHEN_BUHIN))        '変更部品
        Call UniCode_Conv(PCB_U_REC.HEN_NAIYO, PCB_U(Row, colHEN_NAIYO))        '変更内容

        Call UniCode_Conv(PCB_U_REC.HEN_BASHO, PCB_U(Row, colHEN_BASHO))        '交換場所

        Call UniCode_Conv(PCB_U_REC.SETUHEN_HOKAN, PCB_U(Row, colSETUHEN_HOKAN))    '設変原紙保管
        
        Call UniCode_Conv(PCB_U_REC.BIKOU1, PCB_U(Row, colBIKOU1))              '備考１
        Call UniCode_Conv(PCB_U_REC.BIKOU2, PCB_U(Row, colBIKOU2))              '備考２
        Call UniCode_Conv(PCB_U_REC.BIKOU3, PCB_U(Row, colBIKOU3))              '備考３
        Call UniCode_Conv(PCB_U_REC.BIKOU4, PCB_U(Row, colBIKOU4))              '備考４
        
        
        Call UniCode_Conv(PCB_U_REC.FILLER, "")                                 '設変原紙保管

        Call UniCode_Conv(PCB_U_REC.INS_TANTO, App.EXEName)                     '追加担当者
        Call UniCode_Conv(PCB_U_REC.Ins_DateTime, INS_NOW)                      '追加日時

        Call UniCode_Conv(PCB_U_REC.UPD_TANTO, "")
        Call UniCode_Conv(PCB_U_REC.UPD_DATETIME, "")


        Do
            sts = BTRV(BtOpInsert, PCB_U_POS, PCB_U_REC, Len(PCB_U_REC), K0_PCB_U, Len(K0_PCB_U), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    
                    Beep
                    ans = MsgBox("「PCB.U設変」他端末でデータ使用中です。<PCB_U.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Call Input_UnLock
                        Exit Function
                    End If
                
                Case Else
                    Call Input_UnLock
                    Call File_Error(sts, BtOpInsert, "PCB.U設変")
                    Exit Function
            End Select
            
        Loop



    Next Row


'    sts = BTRV(BtOpClose, PCB_U_POS, PCB_U_REC, Len(PCB_U_REC), K0_PCB_U, Len(K0_PCB_U), 0)
'    If sts Then
'        If sts <> BtErrNoOpen Then
'            Call File_Error(sts, BtOpClose, "PCB.U設変管理台帳")
'        End If
'    End If



hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "PCB.U設変管理台帳変換　[登録]処理終了！！", Me.hwnd, 0)



    Call Input_UnLock


    Update_Proc = False

End Function

Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   「OCB.U 設変管理台帳」読込み処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
    
Dim INS_NOW         As String * 14
    
    
    

Dim xlApp           As Object
Dim xlBook          As Object
Dim xlSheet         As Object

Dim Row             As Long
Dim i               As Long


Dim END_GYO         As Integer
Dim SKIP_F          As Boolean








    List_Disp_Proc = True

    Call Input_Lock

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "PCB.U設変管理台帳変換　[読込]処理開始！！", Me.hwnd, 0)


    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    
    

    xlApp.Workbooks.Open (Text1.Text), ReadOnly:=True, UpdateLinks:=0



    Set PCB_U = Nothing
    
    Row = Min_Row - 1
    lblDisp_Count.Caption = ""


    Set xlSheet = xlApp.Worksheets(1)
    xlSheet.Activate
    
    END_GYO = 0
    For i = 1 To 1048576
            
            SKIP_F = False
            
            If Trim(xlSheet.Application.Cells(i, selKANRI_NO)) = "" And _
                Trim(xlSheet.Application.Cells(i, selEX_DATE)) = "" And _
                Trim(xlSheet.Application.Cells(i, selSETUHEN_NO)) = "" Then
            
                SKIP_F = True
                END_GYO = END_GYO + 1
                
                If END_GYO > 5 Then
                    Exit For
                End If
            Else
                
                
                
            END_GYO = 0
        
            If Trim(xlSheet.Application.Cells(i, selAFT_HIN_GAI)) = "" Then
                SKIP_F = True
            Else
                    
                If Not IsDate(Trim(xlSheet.Application.Cells(i, selEX_DATE))) Then
                    SKIP_F = True
                End If
                
                If Not SKIP_F Then
                
                    Row = Row + 1
                    PCB_U.ReDim Min_Row, Row, Min_Col, Max_Col
                        
                        
                    PCB_U(Row, colKANRI_NO) = Trim(xlSheet.Application.Cells(i, selKANRI_NO))
                    PCB_U(Row, colEX_DATE) = Trim(xlSheet.Application.Cells(i, selEX_DATE))
                    PCB_U(Row, colSETUHEN_NO) = Trim(xlSheet.Application.Cells(i, selSETUHEN_NO))
                    PCB_U(Row, colBEF_HIN_GAI) = Trim(xlSheet.Application.Cells(i, selBEF_HIN_GAI))
                    PCB_U(Row, colBEF_HIN_NAI) = Trim(xlSheet.Application.Cells(i, selBEF_HIN_NAI))
                    PCB_U(Row, colAFT_HIN_GAI) = Trim(xlSheet.Application.Cells(i, selAFT_HIN_GAI))
                    PCB_U(Row, colAFT_HIN_NAI) = Trim(xlSheet.Application.Cells(i, selAFT_HIN_NAI))
                    
                    PCB_U(Row, colSETUHEN_JITSU) = Trim(xlSheet.Application.Cells(i, selSETUHEN_JITSU))
                    
                    
                    PCB_U(Row, colHEN_BUHIN) = Trim(xlSheet.Application.Cells(i, selHEN_BUHIN))
                    PCB_U(Row, colHEN_NAIYO) = Trim(xlSheet.Application.Cells(i, selHEN_NAIYO))
                    PCB_U(Row, colHEN_BASHO) = Trim(xlSheet.Application.Cells(i, selHEN_BASHO))
                    PCB_U(Row, colSETUHEN_HOKAN) = Trim(xlSheet.Application.Cells(i, selSETUHEN_HOKAN))
                    PCB_U(Row, colBIKOU1) = Trim(xlSheet.Application.Cells(i, selBIKOU1))
                    PCB_U(Row, colBIKOU2) = Trim(xlSheet.Application.Cells(i, selBIKOU2))
                    PCB_U(Row, colBIKOU3) = Trim(xlSheet.Application.Cells(i, selBIKOU3))
                    PCB_U(Row, colBIKOU4) = Trim(xlSheet.Application.Cells(i, selBIKOU4))
    
    
                    End If
                End If
            
            End If
        
        
        
        Next i













    Set TDBGrid1.Array = PCB_U
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst

    lblDisp_Count.Caption = Row

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "PCB.U設変管理台帳変換　[読込]処理終了！！", Me.hwnd, 0)



    Call Input_UnLock


    List_Disp_Proc = False
    Exit Function


End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    PCB00101.MousePointer = vbHourglass

    Call Ctrl_Lock(PCB00101)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(PCB00101)


    PCB00101.MousePointer = vbDefault

End Sub

Private Sub Text1_OLESetData(Data As DataObject, DataFormat As Integer)
'    If DataFormat = vbCFText Then
'        Data.SetData Text1.SelText, vbCFText
'    End If
End Sub
