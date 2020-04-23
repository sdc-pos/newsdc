VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form SEM00401 
   Caption         =   "[請求システム]品名カテゴリーマスタメンテナンス([SEM0040] 2015.04.27　12：00)"
   ClientHeight    =   10035
   ClientLeft      =   2025
   ClientTop       =   2325
   ClientWidth     =   14850
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
   ScaleHeight     =   10035
   ScaleWidth      =   14850
   StartUpPosition =   2  '画面の中央
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   7695
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   13573
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   4
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "削"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "ｺｰﾄﾞ"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "名　　称"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "見積　　　　　　ﾛｯﾄ数"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "前後工数　　   　（秒/ﾛｯﾄ）"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "前後工数　    　　(秒/個)"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "特別単価         (作業工数 秒/個)"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "特別単価　   　　（工料＠）"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "特別単価　　   　（箱代＠）"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "メ　　　モ"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "追加日時"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "追加担当者"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "更新日時"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "更新担当者"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   14
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AllowColMove=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=14"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=476"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2249"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2117"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=4366"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=4233"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=1879"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=1746"
      Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=2619"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=2487"
      Splits(0)._ColumnProps(22)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(24)=   "Column(5).Width=2619"
      Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=2487"
      Splits(0)._ColumnProps(27)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(29)=   "Column(6).Width=2910"
      Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=2778"
      Splits(0)._ColumnProps(32)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(33)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(34)=   "Column(7).Width=2619"
      Splits(0)._ColumnProps(35)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(7)._WidthInPix=2487"
      Splits(0)._ColumnProps(37)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(38)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(39)=   "Column(8).Width=2619"
      Splits(0)._ColumnProps(40)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(8)._WidthInPix=2487"
      Splits(0)._ColumnProps(42)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(43)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(44)=   "Column(9).Width=7276"
      Splits(0)._ColumnProps(45)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(9)._WidthInPix=7144"
      Splits(0)._ColumnProps(47)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(48)=   "Column(10).Width=3149"
      Splits(0)._ColumnProps(49)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(10)._WidthInPix=3016"
      Splits(0)._ColumnProps(51)=   "Column(10)._ColStyle=8196"
      Splits(0)._ColumnProps(52)=   "Column(10).AllowFocus=0"
      Splits(0)._ColumnProps(53)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(54)=   "Column(11).Width=1720"
      Splits(0)._ColumnProps(55)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(11)._WidthInPix=1588"
      Splits(0)._ColumnProps(57)=   "Column(11)._ColStyle=8196"
      Splits(0)._ColumnProps(58)=   "Column(11).AllowFocus=0"
      Splits(0)._ColumnProps(59)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(60)=   "Column(12).Width=3149"
      Splits(0)._ColumnProps(61)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(12)._WidthInPix=3016"
      Splits(0)._ColumnProps(63)=   "Column(12)._ColStyle=8196"
      Splits(0)._ColumnProps(64)=   "Column(12).AllowFocus=0"
      Splits(0)._ColumnProps(65)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(66)=   "Column(13).Width=1720"
      Splits(0)._ColumnProps(67)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(13)._WidthInPix=1588"
      Splits(0)._ColumnProps(69)=   "Column(13)._ColStyle=8196"
      Splits(0)._ColumnProps(70)=   "Column(13).AllowFocus=0"
      Splits(0)._ColumnProps(71)=   "Column(13).Order=14"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFF&,.bold=0,.fontsize=1200"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=43,.parent=1,.bgcolor=&HFFFF00&,.bold=0,.fontsize=1200"
      _StyleDefs(25)  =   ":id=43,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(26)  =   ":id=43,.fontname=ＭＳ ゴシック"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=45,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=47,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=50,.parent=9,.bgcolor=&HFFFFFF&"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=51,.parent=10,.bgcolor=&HFFFF00&"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=82,.parent=43,.alignment=2"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=44"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=45"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=47"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=16,.parent=43"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=13,.parent=44"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=14,.parent=45"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=15,.parent=47"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=58,.parent=43"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=44"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=45"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=47"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=70,.parent=43,.alignment=1"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=44"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=45"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=47"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=28,.parent=43,.alignment=1,.bold=0,.fontsize=1200"
      _StyleDefs(55)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(56)  =   ":id=28,.fontname=ＭＳ ゴシック"
      _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=44"
      _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=45"
      _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=47"
      _StyleDefs(60)  =   "Splits(0).Columns(5).Style:id=66,.parent=43,.alignment=1,.bold=0,.fontsize=1200"
      _StyleDefs(61)  =   ":id=66,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(62)  =   ":id=66,.fontname=ＭＳ ゴシック"
      _StyleDefs(63)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=44"
      _StyleDefs(64)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=45"
      _StyleDefs(65)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=47"
      _StyleDefs(66)  =   "Splits(0).Columns(6).Style:id=32,.parent=43,.alignment=1"
      _StyleDefs(67)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=44"
      _StyleDefs(68)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=45"
      _StyleDefs(69)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=47"
      _StyleDefs(70)  =   "Splits(0).Columns(7).Style:id=62,.parent=43,.alignment=1"
      _StyleDefs(71)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=44"
      _StyleDefs(72)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=45"
      _StyleDefs(73)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=47"
      _StyleDefs(74)  =   "Splits(0).Columns(8).Style:id=74,.parent=43,.alignment=1"
      _StyleDefs(75)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=44"
      _StyleDefs(76)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=45"
      _StyleDefs(77)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=47"
      _StyleDefs(78)  =   "Splits(0).Columns(9).Style:id=90,.parent=43"
      _StyleDefs(79)  =   "Splits(0).Columns(9).HeadingStyle:id=87,.parent=44"
      _StyleDefs(80)  =   "Splits(0).Columns(9).FooterStyle:id=88,.parent=45"
      _StyleDefs(81)  =   "Splits(0).Columns(9).EditorStyle:id=89,.parent=47"
      _StyleDefs(82)  =   "Splits(0).Columns(10).Style:id=24,.parent=43,.bgcolor=&HC0C0C0&,.locked=-1"
      _StyleDefs(83)  =   "Splits(0).Columns(10).HeadingStyle:id=21,.parent=44"
      _StyleDefs(84)  =   "Splits(0).Columns(10).FooterStyle:id=22,.parent=45"
      _StyleDefs(85)  =   "Splits(0).Columns(10).EditorStyle:id=23,.parent=47"
      _StyleDefs(86)  =   "Splits(0).Columns(11).Style:id=20,.parent=43,.bgcolor=&HC0C0C0&,.locked=-1"
      _StyleDefs(87)  =   "Splits(0).Columns(11).HeadingStyle:id=17,.parent=44"
      _StyleDefs(88)  =   "Splits(0).Columns(11).FooterStyle:id=18,.parent=45"
      _StyleDefs(89)  =   "Splits(0).Columns(11).EditorStyle:id=19,.parent=47"
      _StyleDefs(90)  =   "Splits(0).Columns(12).Style:id=78,.parent=43,.bgcolor=&HC0C0C0&,.locked=-1"
      _StyleDefs(91)  =   "Splits(0).Columns(12).HeadingStyle:id=75,.parent=44"
      _StyleDefs(92)  =   "Splits(0).Columns(12).FooterStyle:id=76,.parent=45"
      _StyleDefs(93)  =   "Splits(0).Columns(12).EditorStyle:id=77,.parent=47"
      _StyleDefs(94)  =   "Splits(0).Columns(13).Style:id=86,.parent=43,.bgcolor=&HC0C0C0&,.locked=-1"
      _StyleDefs(95)  =   "Splits(0).Columns(13).HeadingStyle:id=83,.parent=44"
      _StyleDefs(96)  =   "Splits(0).Columns(13).FooterStyle:id=84,.parent=45"
      _StyleDefs(97)  =   "Splits(0).Columns(13).EditorStyle:id=85,.parent=47"
      _StyleDefs(98)  =   "Named:id=33:Normal"
      _StyleDefs(99)  =   ":id=33,.parent=0"
      _StyleDefs(100) =   "Named:id=34:Heading"
      _StyleDefs(101) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(102) =   ":id=34,.wraptext=-1"
      _StyleDefs(103) =   "Named:id=35:Footing"
      _StyleDefs(104) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(105) =   "Named:id=36:Selected"
      _StyleDefs(106) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(107) =   "Named:id=37:Caption"
      _StyleDefs(108) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(109) =   "Named:id=38:HighlightRow"
      _StyleDefs(110) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(111) =   "Named:id=39:EvenRow"
      _StyleDefs(112) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(113) =   "Named:id=40:OddRow"
      _StyleDefs(114) =   ":id=40,.parent=33"
      _StyleDefs(115) =   "Named:id=41:RecordSelector"
      _StyleDefs(116) =   ":id=41,.parent=34"
      _StyleDefs(117) =   "Named:id=42:FilterBar"
      _StyleDefs(118) =   ":id=42,.parent=33"
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   5565
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   1200
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      Caption         =   "検  索"
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
      Left            =   3780
      TabIndex        =   9
      Top             =   360
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
      Height          =   375
      Index           =   1
      Left            =   1785
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2430
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   1050
      MaxLength       =   5
      TabIndex        =   0
      Top             =   1200
      Width           =   750
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終　了"
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
      Left            =   2100
      TabIndex        =   5
      Top             =   360
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "編集確定"
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
      Left            =   420
      TabIndex        =   4
      Top             =   360
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "事業部"
      Height          =   255
      Index           =   1
      Left            =   4725
      TabIndex        =   10
      Top             =   1320
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "担当者"
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   8
      Top             =   1320
      Width           =   750
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
      Left            =   210
      TabIndex        =   6
      Top             =   9360
      Width           =   180
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "編集確定"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "検索"
         Index           =   2
      End
      Begin VB.Menu SHORI 
         Caption         =   "画面印刷"
         Index           =   3
      End
   End
End
Attribute VB_Name = "SEM00401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxTanto_Code% = 0        '担当者ｺｰﾄﾞ
Private Const ptxTanto_Name% = 1        '担当者名称


Private Const pcmbJGYOBU% = 0           '事業部


Dim ITEM_CATEGORY As New XArrayDB

Private Const Min_Row% = 1              '最小行数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 13             '最大列数

Private Const ColSHORI% = 0             '削除
Private Const ColCATEGORY_CODE% = 1     '品名ｶﾃｺﾞﾘｺｰﾄﾞ
Private Const ColCATEGORY_NAME% = 2     '品名ｶﾃｺﾞﾘ名称
Private Const ColSEI_LOT% = 3           '生産ロット
Private Const ColKOUSU_LOT% = 4         '前後工数(秒/ﾛｯﾄ)
Private Const ColKOUSU_QTY% = 5         '前後工数(秒/個)
Private Const ColTOKU_TANKA_QTY% = 6    '特別単価(作業工数　秒/個)
Private Const ColTOKU_TANKA_KOURYO% = 7 '特別単価(工料＠)
Private Const ColTOKU_TANKA_HAKO% = 8   '特別単価(箱代＠)
Private Const ColMEMO% = 9              '備考/メモ
Private Const ColINS_DATETIME% = 10     '追加　日時
Private Const ColINS_TANTO% = 11        '追加　担当者
Private Const ColUPD_DATETIME% = 12     '更新　日時
Private Const ColUPD_TANTO% = 13        '更新　担当者





Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    If Error_Check_Proc(Index) Then     'エラーチェック
        Exit Sub
    End If
        
    Command1(2).Value = True
    

End Sub

Private Sub Command1_Click(Index As Integer)

Dim yn  As Integer
Dim i   As Integer


    Select Case Index
    
        Case 0
    
    
            
            
            For i = ptxTanto_Code To ptxTanto_Name
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
    
            If Grid_Error_Check_Proc() Then
                Exit Sub
            End If
    
    
            yn = MsgBox("編集内容を確定しますか？", vbYesNo, "確認入力")
    
            If yn = vbYes Then
        
                If Update_Proc() Then
                    Unload Me
                End If
            
                If List_Disp_Proc() Then
                    Unload Me
                End If
            
                DoEvents
                
                MsgBox "編集内容の書き込み処理が終了しました。"
            
            
            End If
        
        
        
        
        Case 1
    
            Unload Me
    
    
        Case 2
            
            If List_Disp_Proc() Then
                Unload Me
            End If
    
    End Select





End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
Dim sts As Integer




    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If


    'コモンコントロールを初期化する
'    cc.dwSize = Len(cc)
'    cc.dwICC = ICC_BAR_CLASSES
    
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]品名カテゴリーマスタメンテナンス", Me.hwnd, 0)
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
                                '事業部取り込み
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
        
                                
    Combo1(pcmbJGYOBU).Clear
    For i = 0 To UBound(JGYOBU_T)
        Combo1(pcmbJGYOBU).AddItem JGYOBU_T(i).NAME & "                 " & JGYOBU_T(i).CODE
    Next i
                                
                                
                                    
                                'デフォルト事業部取り込み
    If GetIni("JIGYOBU", "DEF_NO", "SYS", c) Then
    Else
    
    
        For i = 0 To Combo1(pcmbJGYOBU).ListCount - 1
            If Trim(c) = Right(Combo1(pcmbJGYOBU).List(i), 1) Then
                Combo1(pcmbJGYOBU).ListIndex = i
                Exit For
            End If
        Next i
    
    
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品名カテゴリマスタＯＰＥＮ
    If ITEM_CATEGORY_Open(BtOpenNomal) Then
        Unload Me
    End If
                                



    Text1(ptxTanto_Code).SetFocus


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
Dim yn  As Integer
                                            
                                            
    yn = MsgBox("終了しますか？", vbYesNo, "確認入力")
    If yn = vbNo Then
        Cancel = True
        Exit Sub
    End If
                                            
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    End
End Sub
Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   データ内容の表示
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim Row         As Long
    
    
    List_Disp_Proc = True
                                    
    Call Input_Lock
                                    
Call SendMessageStr(hStatusWnd, SB_SETTEXT, 0, "検索処理　開始")
                        
                        'テーブルリセット
    Set ITEM_CATEGORY = Nothing
    Row = Min_Row - 1
        
    Last_JGYOBU = Right(Combo1(pcmbJGYOBU).Text, 1)
                        '品名ｶﾃｺﾞﾘﾏｽﾀ読み込み開始
    Call UniCode_Conv(K0_ITEM_CATEGORY.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM_CATEGORY.CATEGORY_CODE, "")
    
    com = BtOpGetGreaterEqual
    
    Do
        DoEvents
        sts = BTRV(com, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
        Select Case sts
            Case BtNoErr
                If Last_JGYOBU <> StrConv(ITEM_CATEGORYREC.JGYOBU, vbUnicode) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品名ｶﾃｺﾞﾘﾏｽﾀ")
                List_Disp_Proc = SYS_ERR
                Exit Function
        End Select
        
        Row = Row + 1
                        
                        
        If Grid_Set_Proc(Row) Then
            Exit Function
        End If
        
        
        com = BtOpGetNext
        
    Loop
                                'DBテーブルリンク
    Set TDBGrid1.Array = ITEM_CATEGORY
    
    
    TDBGrid1.Bookmark = Null
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.ScrollBars = dbgAutomatic
    
    If ITEM_CATEGORY.Count(1) > 0 Then
        TDBGrid1.MoveFirst
    End If
    
Call SendMessageStr(hStatusWnd, SB_SETTEXT, 0, "検索処理　終了")
    
    
    Call Input_UnLock
    
    
    List_Disp_Proc = False

    
End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    SEM00401.MousePointer = vbHourglass

    Call Ctrl_Lock(SEM00401)

    TDBGrid1.Enabled = False

End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(SEM00401)

    TDBGrid1.Enabled = True

    SEM00401.MousePointer = vbDefault

End Sub

Private Function Grid_Set_Proc(Row As Long) As Integer

Dim sts As Integer

    
    Grid_Set_Proc = True

    

    ITEM_CATEGORY.ReDim Min_Row, Row, Min_Col, Max_Col
    
    
    '削除
    ITEM_CATEGORY(Row, ColSHORI) = False
    '品名ｶﾃｺﾞﾘｺｰﾄﾞ
    ITEM_CATEGORY(Row, ColCATEGORY_CODE) = Trim(StrConv(ITEM_CATEGORYREC.CATEGORY_CODE, vbUnicode))
    '品名ｶﾃｺﾞﾘ名称
    ITEM_CATEGORY(Row, ColCATEGORY_NAME) = Trim(StrConv(ITEM_CATEGORYREC.CATEGORY_NAME, vbUnicode))
    '見積ﾛｯﾄ数
    If IsNumeric(StrConv(ITEM_CATEGORYREC.SEI_LOT, vbUnicode)) Then
        If Val(StrConv(ITEM_CATEGORYREC.SEI_LOT, vbUnicode)) = 0 Then
            ITEM_CATEGORY(Row, ColSEI_LOT) = ""
        Else
            ITEM_CATEGORY(Row, ColSEI_LOT) = Val(StrConv(ITEM_CATEGORYREC.SEI_LOT, vbUnicode))
        End If
    Else
        ITEM_CATEGORY(Row, ColSEI_LOT) = ""
    End If
    '前後工数(秒/ﾛｯﾄ)
    If IsNumeric(StrConv(ITEM_CATEGORYREC.KOUSU_LOT, vbUnicode)) Then
        If Val(StrConv(ITEM_CATEGORYREC.KOUSU_LOT, vbUnicode)) = 0 Then
            ITEM_CATEGORY(Row, ColKOUSU_LOT) = ""
        Else
            ITEM_CATEGORY(Row, ColKOUSU_LOT) = Val(StrConv(ITEM_CATEGORYREC.KOUSU_LOT, vbUnicode))
        End If
    Else
        ITEM_CATEGORY(Row, ColKOUSU_LOT) = ""
    End If
    '前後工数(秒/個)
    If IsNumeric(StrConv(ITEM_CATEGORYREC.KOUSU_QTY, vbUnicode)) Then
        If Val(StrConv(ITEM_CATEGORYREC.KOUSU_QTY, vbUnicode)) = 0 Then
            ITEM_CATEGORY(Row, ColKOUSU_QTY) = ""
        Else
            ITEM_CATEGORY(Row, ColKOUSU_QTY) = Val(StrConv(ITEM_CATEGORYREC.KOUSU_QTY, vbUnicode))
        End If
    Else
        ITEM_CATEGORY(Row, ColKOUSU_QTY) = ""
    End If
    
    '特別単価(作業工数　秒/個)
    If IsNumeric(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_QTY, vbUnicode)) Then
'        If Val(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_QTY, vbUnicode)) = 0 Then
'            ITEM_CATEGORY(Row, ColTOKU_TANKA_QTY) = ""
'        Else
            ITEM_CATEGORY(Row, ColTOKU_TANKA_QTY) = Val(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_QTY, vbUnicode))
'        End If
    Else
        ITEM_CATEGORY(Row, ColTOKU_TANKA_QTY) = ""
    End If
    '特別単価(工料＠)
    If IsNumeric(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_KOURYO, vbUnicode)) Then
'        If Val(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_KOURYO, vbUnicode)) = 0 Then
'            ITEM_CATEGORY(Row, ColTOKU_TANKA_QTY) = ""
'        Else
            ITEM_CATEGORY(Row, ColTOKU_TANKA_KOURYO) = Format(Val(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_KOURYO, vbUnicode)), "0.00")
'        End If
    Else
        ITEM_CATEGORY(Row, ColTOKU_TANKA_QTY) = ""
    End If
    '特別単価(箱代＠)
    If IsNumeric(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, vbUnicode)) Then
'        If Val(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, vbUnicode)) = 0 Then
'            ITEM_CATEGORY(Row, ColTOKU_TANKA_HAKO) = ""
'        Else
            ITEM_CATEGORY(Row, ColTOKU_TANKA_HAKO) = Format(Val(StrConv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, vbUnicode)), "0.00")
'        End If
    Else
        ITEM_CATEGORY(Row, ColTOKU_TANKA_HAKO) = ""
    End If
    '備考/メモ
    ITEM_CATEGORY(Row, ColMEMO) = Trim(StrConv(ITEM_CATEGORYREC.MEMO, vbUnicode))
    '追加日時
    ITEM_CATEGORY(Row, ColINS_DATETIME) = Trim(StrConv(ITEM_CATEGORYREC.Ins_DateTime, vbUnicode))
    '追加担当者
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(ITEM_CATEGORYREC.INS_TANTO, vbUnicode))
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            ITEM_CATEGORY(Row, ColINS_TANTO) = StrConv(ITEM_CATEGORYREC.INS_TANTO, vbUnicode) & " " & Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
        Case BtErrKeyNotFound
            ITEM_CATEGORY(Row, ColINS_TANTO) = StrConv(ITEM_CATEGORYREC.INS_TANTO, vbUnicode)
        Case Else
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Exit Function
    End Select
    
    
    '更新日時
    ITEM_CATEGORY(Row, ColUPD_DATETIME) = Trim(StrConv(ITEM_CATEGORYREC.UPD_DATETIME, vbUnicode))
    '更新担当者
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(ITEM_CATEGORYREC.UPD_TANTO, vbUnicode))
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            ITEM_CATEGORY(Row, ColUPD_TANTO) = StrConv(ITEM_CATEGORYREC.UPD_TANTO, vbUnicode) & " " & Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
        Case BtErrKeyNotFound
            ITEM_CATEGORY(Row, ColUPD_TANTO) = StrConv(ITEM_CATEGORYREC.UPD_TANTO, vbUnicode)
        Case Else
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Exit Function
    End Select
    
    
    
    Grid_Set_Proc = False
End Function



Private Sub SHORI_Click(Index As Integer)
    Select Case Index
    
        
        Case 0      '更新
        
        
            Command1(Index).Value = True
        
        
        Case 1      '終了
        
        
            Command1(Index).Value = True
        
        Case 2      '検索
        
        
            Command1(Index).Value = True
        
        
                    
    
    End Select

End Sub




Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   データ更新
'----------------------------------------------------------------------------
Dim sts         As Integer
    
Dim i           As Integer
    
Dim com         As Integer
    
Dim Upd_Now     As String
    
    
    Update_Proc = True
                                     
Call SendMessageStr(hStatusWnd, SB_SETTEXT, 0, "編集書き込み処理　開始")
                                     
                                     
    Set TDBGrid1.Array = ITEM_CATEGORY
    TDBGrid1.Refresh
    
    TDBGrid1.Update
                                     
    If ITEM_CATEGORY.Count(1) < 1 Then
        Update_Proc = False
        Exit Function
    End If
                                     
    Call Input_Lock
                                    
                                    
    Upd_Now = Format(Now, "YYYYMMDDHHMMSS")
                                    
    For i = 1 To ITEM_CATEGORY.Count(1)
                                    
        Call UniCode_Conv(K0_ITEM_CATEGORY.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_ITEM_CATEGORY.CATEGORY_CODE, ITEM_CATEGORY(i, ColCATEGORY_CODE))
                                
        sts = BTRV(BtOpGetEqual, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
        Select Case sts
            Case BtNoErr
                
                com = BtOpUpdate
            
            Case BtErrKeyNotFound
                
                com = BtOpInsert
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品名カテゴリーマスタ")
                Exit Function
        End Select
                                        
                                
        If ITEM_CATEGORY(i, ColSHORI) Then
                                
            sts = BTRV(BtOpDelete, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
            Select Case sts
                Case BtNoErr
                    
                
                Case BtErrKeyNotFound
                    
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品名カテゴリーマスタ")
                    Exit Function
            End Select
                                
                                
        Else
            
            If com = BtOpInsert Then
                                                                
                Call UniCode_Conv(ITEM_CATEGORYREC.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(ITEM_CATEGORYREC.CATEGORY_CODE, ITEM_CATEGORY(i, ColCATEGORY_CODE))
                Call UniCode_Conv(ITEM_CATEGORYREC.FILLER, "")
                Call UniCode_Conv(ITEM_CATEGORYREC.INS_TANTO, Text1(ptxTanto_Code).Text)
                Call UniCode_Conv(ITEM_CATEGORYREC.Ins_DateTime, Upd_Now)
            
                Call UniCode_Conv(ITEM_CATEGORYREC.UPD_TANTO, "")
                Call UniCode_Conv(ITEM_CATEGORYREC.UPD_DATETIME, "")
            
            
            End If



            Call UniCode_Conv(ITEM_CATEGORYREC.CATEGORY_NAME, Trim(ITEM_CATEGORY(i, ColCATEGORY_NAME)))

            If Len(ITEM_CATEGORY(i, ColSEI_LOT)) < 10 Then
                Call UniCode_Conv(ITEM_CATEGORYREC.SEI_LOT, String(10 - Len(ITEM_CATEGORY(i, ColSEI_LOT)), "0") & ITEM_CATEGORY(i, ColSEI_LOT))
            Else
                Call UniCode_Conv(ITEM_CATEGORYREC.SEI_LOT, ITEM_CATEGORY(i, ColSEI_LOT))
            End If


            If Len(ITEM_CATEGORY(i, ColKOUSU_LOT)) < 10 Then
                Call UniCode_Conv(ITEM_CATEGORYREC.KOUSU_LOT, String(10 - Len(ITEM_CATEGORY(i, ColKOUSU_LOT)), "0") & ITEM_CATEGORY(i, ColKOUSU_LOT))
            Else
                Call UniCode_Conv(ITEM_CATEGORYREC.KOUSU_LOT, ITEM_CATEGORY(i, ColKOUSU_LOT))
            End If

            If Len(ITEM_CATEGORY(i, ColKOUSU_QTY)) < 10 Then
                Call UniCode_Conv(ITEM_CATEGORYREC.KOUSU_QTY, String(10 - Len(ITEM_CATEGORY(i, ColKOUSU_QTY)), "0") & ITEM_CATEGORY(i, ColKOUSU_QTY))
            Else
                Call UniCode_Conv(ITEM_CATEGORYREC.KOUSU_QTY, ITEM_CATEGORY(i, ColKOUSU_QTY))
            End If

            If Trim(ITEM_CATEGORY(i, ColTOKU_TANKA_QTY)) = "" Then
                
                Call UniCode_Conv(ITEM_CATEGORYREC.TOKU_TANKA_QTY, "")
            Else
                If Len(ITEM_CATEGORY(i, ColTOKU_TANKA_QTY)) < 10 Then
                    Call UniCode_Conv(ITEM_CATEGORYREC.TOKU_TANKA_QTY, String(10 - Len(ITEM_CATEGORY(i, ColTOKU_TANKA_QTY)), "0") & ITEM_CATEGORY(i, ColTOKU_TANKA_QTY))
                Else
                    Call UniCode_Conv(ITEM_CATEGORYREC.TOKU_TANKA_QTY, ITEM_CATEGORY(i, ColTOKU_TANKA_QTY))
                End If
            End If

            If Trim(ITEM_CATEGORY(i, ColTOKU_TANKA_KOURYO)) = "" Then
                Call UniCode_Conv(ITEM_CATEGORYREC.TOKU_TANKA_KOURYO, "")
            Else
                Call UniCode_Conv(ITEM_CATEGORYREC.TOKU_TANKA_KOURYO, Format(Val(ITEM_CATEGORY(i, ColTOKU_TANKA_KOURYO)), "0000000000.00"))
            End If
            
            If Trim(ITEM_CATEGORY(i, ColTOKU_TANKA_HAKO)) = "" Then
                Call UniCode_Conv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, "")
            Else
                Call UniCode_Conv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, Format(Val(ITEM_CATEGORY(i, ColTOKU_TANKA_HAKO)), "0000000000.00"))
            End If

            Call UniCode_Conv(ITEM_CATEGORYREC.MEMO, Trim(ITEM_CATEGORY(i, ColMEMO)))
    
    
            If com = BtOpUpdate Then
                                                                
                Call UniCode_Conv(ITEM_CATEGORYREC.UPD_TANTO, Text1(ptxTanto_Code).Text)
                Call UniCode_Conv(ITEM_CATEGORYREC.UPD_DATETIME, Upd_Now)
            
            End If
               
 

            sts = BTRV(com, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
            Select Case sts
                Case BtNoErr
                
                Case Else
                    Call File_Error(sts, com, "品名カテゴリーマスタ")
                    Exit Function
            End Select
    
    
        End If
    Next i
                                    
                                    
                                    
    Call Input_UnLock
                                        
Call SendMessageStr(hStatusWnd, SB_SETTEXT, 0, "編集書き込み処理　終了")
                                        
                                        
    
    
    Update_Proc = False
    


End Function



Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
    
    If TDBGrid1.Bookmark = Null Then
        Exit Sub
    End If
    
    If TDBGrid1.Bookmark < 1 Then
        Exit Sub
    End If
    
    
    
    
    
    Set TDBGrid1.Array = ITEM_CATEGORY

    TDBGrid1.Refresh
    
    TDBGrid1.Update

    Select Case ColIndex
    
        
        Case ColSEI_LOT                 '見積ﾛｯﾄ数
        
            If IsNumeric(ITEM_CATEGORY(TDBGrid1.Bookmark, ColSEI_LOT)) And _
                IsNumeric(ITEM_CATEGORY(TDBGrid1.Bookmark, ColKOUSU_LOT)) Then
            
                
                If CCur(ITEM_CATEGORY(TDBGrid1.Bookmark, ColSEI_LOT)) <> 0 Then         '2015.04.27
                    ITEM_CATEGORY(TDBGrid1.Bookmark, ColKOUSU_QTY) = Val(ToHalfAdjust(CCur(ITEM_CATEGORY(TDBGrid1.Bookmark, ColKOUSU_LOT)) / CCur(ITEM_CATEGORY(TDBGrid1.Bookmark, ColSEI_LOT)), 0))
                Else                                                                    '2015.04.27
                    ITEM_CATEGORY(TDBGrid1.Bookmark, ColKOUSU_QTY) = 0                  '2015.04.27
                End If                                                                  '2015.04.27
            
            End If
        
        Case ColKOUSU_LOT               '前後工数(秒/ﾛｯﾄ)
    
            If IsNumeric(ITEM_CATEGORY(TDBGrid1.Bookmark, ColSEI_LOT)) And _
                IsNumeric(ITEM_CATEGORY(TDBGrid1.Bookmark, ColKOUSU_LOT)) Then
            
                If CCur(ITEM_CATEGORY(TDBGrid1.Bookmark, ColSEI_LOT)) <> 0 Then         '2015.04.27
                    ITEM_CATEGORY(TDBGrid1.Bookmark, ColKOUSU_QTY) = Val(ToHalfAdjust(CCur(ITEM_CATEGORY(TDBGrid1.Bookmark, ColKOUSU_LOT)) / CCur(ITEM_CATEGORY(TDBGrid1.Bookmark, ColSEI_LOT)), 0))
                Else                                                                    '2015.04.27
                    ITEM_CATEGORY(TDBGrid1.Bookmark, ColKOUSU_QTY) = 0                  '2015.04.27
                End If                                                                  '2015.04.27
            
            End If
    
    
    
    End Select


    Set TDBGrid1.Array = ITEM_CATEGORY
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.ScrollBars = dbgAutomatic
    TDBGrid1.SetFocus




End Sub

Private Sub TDBGrid1_BeforeInsert(Cancel As Integer)
    
    ITEM_CATEGORY.ReDim Min_Row, ITEM_CATEGORY.Count(1), Min_Col, Max_Col

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
Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts As Integer
    
    
    Error_Check_Proc = True
    
    Select Case Mode
    
    
        Case ptxTanto_Code     '担当者ｺｰﾄﾞ
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTanto_Code).Text)
            
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    Text1(ptxTanto_Name).Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    Text1(ptxTanto_Name).Text = ""
                    MsgBox "入力した項目はエラーです。(担当者)"
                    Text1(ptxTanto_Code).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Exit Function
            End Select
            
            
    
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function

Private Function Grid_Error_Check_Proc() As Integer
'----------------------------------------------------------------------------
'                   グリッド入力項目のエラーチェック
'----------------------------------------------------------------------------
Dim i   As Integer
Dim sts As Integer
    
    Grid_Error_Check_Proc = True
    
    
    
    
    Set TDBGrid1.Array = ITEM_CATEGORY
    
    
    TDBGrid1.Update
    
    If ITEM_CATEGORY.Count(1) < 1 Then
        Grid_Error_Check_Proc = False
        Exit Function
    End If
    
    
    
    
    
    
    For i = 1 To ITEM_CATEGORY.Count(1)
        
        
        If ITEM_CATEGORY(i, ColSHORI) Then
        Else
        
            If Trim(ITEM_CATEGORY(i, ColSEI_LOT)) = "" Then
            Else
                If IsNumeric(ITEM_CATEGORY(i, ColSEI_LOT)) Then
                Else
                    MsgBox "入力した項目は、エラーです。（生産ﾛｯﾄ）" & i & "行目"
                    
                    TDBGrid1.Bookmark = i
                    TDBGrid1.Col = ColSEI_LOT
                    TDBGrid1.SetFocus
                    
                    
                    Exit Function
                End If
            
            End If
            
            
            If Trim(ITEM_CATEGORY(i, ColKOUSU_LOT)) = "" Then
            Else
                If IsNumeric(ITEM_CATEGORY(i, ColKOUSU_LOT)) Then
                Else
                    MsgBox "入力した項目は、エラーです。（前後工数(秒/ﾛｯﾄ)）" & i & "行目"
                    
                    TDBGrid1.Bookmark = i
                    TDBGrid1.Col = ColKOUSU_LOT
                    TDBGrid1.SetFocus
                    
                    
                    Exit Function
                End If
            
            End If
            
            If Trim(ITEM_CATEGORY(i, ColKOUSU_QTY)) = "" Then
            Else
                If IsNumeric(ITEM_CATEGORY(i, ColKOUSU_QTY)) Then
                Else
                    MsgBox "入力した項目は、エラーです。（前後工数(秒/個)）" & i & "行目"
                    
                    TDBGrid1.Bookmark = i
                    TDBGrid1.Col = ColKOUSU_QTY
                    TDBGrid1.SetFocus
                    
                    
                    Exit Function
                End If
            
            End If
            
            
            If Trim(ITEM_CATEGORY(i, ColTOKU_TANKA_QTY)) = "" Then
            Else
                If IsNumeric(ITEM_CATEGORY(i, ColTOKU_TANKA_QTY)) Then
                Else
                    MsgBox "入力した項目は、エラーです。（特別単価(作業工数　秒/個)）" & i & "行目"
                    
                    TDBGrid1.Bookmark = i
                    TDBGrid1.Col = ColTOKU_TANKA_QTY
                    TDBGrid1.SetFocus
                    
                    
                    Exit Function
                End If
            
            End If
            
            
            If Trim(ITEM_CATEGORY(i, ColTOKU_TANKA_KOURYO)) = "" Then
            Else
                If IsNumeric(ITEM_CATEGORY(i, ColTOKU_TANKA_KOURYO)) Then
                Else
                    MsgBox "入力した項目は、エラーです。（特別単価(工料＠)" & i & "行目"
                    
                    TDBGrid1.Bookmark = i
                    TDBGrid1.Col = ColTOKU_TANKA_KOURYO
                    TDBGrid1.SetFocus
                    
                    
                    Exit Function
                End If
            
            End If
            
            
            
            If Trim(ITEM_CATEGORY(i, ColTOKU_TANKA_HAKO)) = "" Then
            Else
                If IsNumeric(ITEM_CATEGORY(i, ColTOKU_TANKA_HAKO)) Then
                Else
                    MsgBox "入力した項目は、エラーです。（特別単価(箱代＠)" & i & "行目"
                    
                    TDBGrid1.Bookmark = i
                    TDBGrid1.Col = ColTOKU_TANKA_HAKO
                    TDBGrid1.SetFocus
                    
                    
                    Exit Function
                End If
            
            End If
        
        End If
    
        
        
        
        
    Next i


    Grid_Error_Check_Proc = False

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

