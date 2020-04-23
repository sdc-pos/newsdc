VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form F1030651 
   BackColor       =   &H00FFFFFF&
   Caption         =   "「過日分」出荷確認"
   ClientHeight    =   6840
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   13125
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   13125
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Height          =   360
      Index           =   0
      Left            =   3240
      MaxLength       =   8
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      Left            =   4440
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   360
      Index           =   2
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   360
      Index           =   1
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   1095
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   4815
      Left            =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   8493
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "注文区分"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "出荷先"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "ID№"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "伝票№"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "品番（外部）"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "品番（内部）"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "品　名"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "出荷数"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "出庫済数"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "検品"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "伝票日付"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "印"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "取り込み日時"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "検品日"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "検品担当者"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   16
      Splits(0)._UserFlags=   0
      Splits(0).Locked=   -1  'True
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=16"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1667"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=3810"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3704"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2408"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2302"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=1349"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1244"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=2646"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2540"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=2646"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2540"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=4921"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=4815"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=1879"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=1773"
      Splits(0)._ColumnProps(32)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(34)=   "Column(8).Width=1879"
      Splits(0)._ColumnProps(35)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(8)._WidthInPix=1773"
      Splits(0)._ColumnProps(37)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(38)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(39)=   "Column(9).Width=926"
      Splits(0)._ColumnProps(40)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(9)._WidthInPix=820"
      Splits(0)._ColumnProps(42)=   "Column(9)._ColStyle=1"
      Splits(0)._ColumnProps(43)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(44)=   "Column(10).Width=2037"
      Splits(0)._ColumnProps(45)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(10)._WidthInPix=1931"
      Splits(0)._ColumnProps(47)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(48)=   "Column(11).Width=476"
      Splits(0)._ColumnProps(49)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(11)._WidthInPix=370"
      Splits(0)._ColumnProps(51)=   "Column(11)._ColStyle=8196"
      Splits(0)._ColumnProps(52)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(53)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(54)=   "Column(12).Width=873"
      Splits(0)._ColumnProps(55)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(12)._WidthInPix=767"
      Splits(0)._ColumnProps(57)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(58)=   "Column(13).Width=3810"
      Splits(0)._ColumnProps(59)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(60)=   "Column(13)._WidthInPix=3704"
      Splits(0)._ColumnProps(61)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(62)=   "Column(14).Width=3810"
      Splits(0)._ColumnProps(63)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(14)._WidthInPix=3704"
      Splits(0)._ColumnProps(65)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(66)=   "Column(15).Width=3969"
      Splits(0)._ColumnProps(67)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(15)._WidthInPix=3863"
      Splits(0)._ColumnProps(69)=   "Column(15).Order=16"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=10.5,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=10.5,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
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
      DeadAreaBackColor=   12632256
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=ＭＳ ゴシック"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bold=0,.fontsize=1050,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.bold=0,.fontsize=1050,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=31,.bold=0,.fontsize=1050,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ＭＳ ゴシック"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=87,.parent=2,.namedParent=89"
      _StyleDefs(23)  =   "FilterBarStyle:id=90,.parent=1,.namedParent=92"
      _StyleDefs(24)  =   "Splits(0).Style:id=53,.parent=1,.bgcolor=&HFFFF&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=62,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=54,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=55,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=56,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=58,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=57,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=59,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=60,.parent=9,.bgcolor=&HFF80&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=61,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=88,.parent=87"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=91,.parent=90"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=14,.parent=53"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=11,.parent=54"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=12,.parent=55"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=13,.parent=57"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=18,.parent=53"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=15,.parent=54"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=16,.parent=55"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=17,.parent=57"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=48,.parent=53"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=45,.parent=54"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=46,.parent=55"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=47,.parent=57"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=66,.parent=53"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=54"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=55"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=57"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=70,.parent=53"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=54"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=55"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=57"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=74,.parent=53"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=71,.parent=54"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=72,.parent=55"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=73,.parent=57"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=78,.parent=53"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=75,.parent=54"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=76,.parent=55"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=77,.parent=57"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=82,.parent=53,.alignment=1,.locked=0"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=79,.parent=54,.alignment=3"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=80,.parent=55,.alignment=3"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=81,.parent=57"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=86,.parent=53,.alignment=1,.locked=0"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=83,.parent=54,.alignment=3"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=84,.parent=55,.alignment=3"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=85,.parent=57"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=24,.parent=53,.alignment=2,.locked=0"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=21,.parent=54,.alignment=3"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=22,.parent=55,.alignment=3"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=23,.parent=57"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=28,.parent=53"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=25,.parent=54"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=26,.parent=55"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=27,.parent=57"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=40,.parent=53,.alignment=3,.locked=-1"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=37,.parent=54,.alignment=3"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=38,.parent=55,.alignment=3"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=39,.parent=57"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=44,.parent=53"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=41,.parent=54"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=42,.parent=55"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=43,.parent=57"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=52,.parent=53"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=49,.parent=54"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=50,.parent=55"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=51,.parent=57"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=96,.parent=53"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=93,.parent=54"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=94,.parent=55"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=95,.parent=57"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=100,.parent=53"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=97,.parent=54"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=98,.parent=55"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=99,.parent=57"
      _StyleDefs(100) =   "Named:id=29:Normal"
      _StyleDefs(101) =   ":id=29,.parent=0"
      _StyleDefs(102) =   "Named:id=30:Heading"
      _StyleDefs(103) =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(104) =   ":id=30,.wraptext=-1"
      _StyleDefs(105) =   "Named:id=31:Footing"
      _StyleDefs(106) =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(107) =   "Named:id=32:Selected"
      _StyleDefs(108) =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(109) =   "Named:id=33:Caption"
      _StyleDefs(110) =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(111) =   "Named:id=34:HighlightRow"
      _StyleDefs(112) =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(113) =   "Named:id=35:EvenRow"
      _StyleDefs(114) =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(115) =   "Named:id=36:OddRow"
      _StyleDefs(116) =   ":id=36,.parent=29"
      _StyleDefs(117) =   "Named:id=89:RecordSelector"
      _StyleDefs(118) =   ":id=89,.parent=30"
      _StyleDefs(119) =   "Named:id=92:FilterBar"
      _StyleDefs(120) =   ":id=92,.parent=29"
   End
   Begin VB.ComboBox Combo 
      Height          =   336
      Index           =   0
      Left            =   1320
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   240
      Width           =   852
   End
   Begin VB.CommandButton Command 
      Caption         =   "終  了"
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "データ"
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "最　新"
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "／"
      Height          =   240
      Index           =   3
      Left            =   11040
      TabIndex        =   22
      Top             =   360
      Width           =   240
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "伝票枚数　実績／予定"
      Height          =   240
      Index           =   2
      Left            =   7320
      TabIndex        =   21
      Top             =   360
      Width           =   2400
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "注文区分"
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   19
      Top             =   360
      Width           =   960
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
      Left            =   120
      TabIndex        =   20
      Top             =   6360
      Width           =   180
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "出荷先"
      Height          =   240
      Index           =   0
      Left            =   2400
      TabIndex        =   18
      Top             =   360
      Width           =   720
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Checked         =   -1  'True
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1030651"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxMUKE_CODE% = 0         '向け先（コード入力用）
Private Const ptxDEN_MAISU_JI% = 1      '伝票枚数　実績
Private Const ptxDEN_MAISU_YO% = 2      '伝票枚数　予定

Private Const pcmbCYU_KBN% = 0          '注文区分
Private Const pcmbMUKE_CODE% = 1        '向け先

Private Const Text_Max% = 2             '画面項目別最大ｲﾝﾃﾞｯｸｽ

Dim SYUKA As New XArrayDB

Private Const Min_Row% = 1              '最小行数
'Private Const Max_Row& = 2000           '最大行数
Dim Max_Row    As Integer               'グリッド最大表示件数

Dim SYUKA_DATA  As String               '出荷データフルパス


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 15             '最大列数

Private Const ColCYU_KBN% = 0           '注文区分
Private Const ColMUKE_CODE% = 1         '出荷先

Private Const ColID_NO% = 2             'ID№
Private Const ColDEN_NO% = 3            '伝票№
Private Const ColHIN_GAI% = 4           '品番（外部）
Private Const ColHIN_NAI% = 5           '品番（内部）
Private Const ColHIN_NAME% = 6          '品名
Private Const ColYOTEI_QTY% = 7         '出荷予定数
Private Const ColFIX_QTY% = 8           '出荷実績
Private Const ColKENPIN_MARK% = 9       '検品
Private Const ColDEN_DT% = 10            '伝票日付
Private Const ColSort_Mark% = 11         'ＳＯＲＴマーク
Private Const ColPrint% = 12            '出庫表印刷マーク
Private Const ColIns_Date% = 13         '取込み日時

Private Const ColKENPIN_Date% = 14      '検品日
Private Const ColKENPIN_TANTO% = 15     '検品担当者

Private Const Sort_MISYUKO$ = "0"       '未出庫
Private Const Sort_SYUKOSUMI$ = "1"     '出庫済
Private Const Sort_KENPIN$ = "2"        '検品済

Private Const KENPIN_ON$ = "○"         '検品済
Private Const KENPIN_OFF$ = "×"        '未検品

Private Sub Combo_Click(Index As Integer)
    Select Case Index
        Case pcmbCYU_KBN
            
            
            Text(ptxMUKE_CODE).SetFocus
    End Select

End Sub

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Select Case Index
        Case pcmbCYU_KBN
            Text(ptxMUKE_CODE).SetFocus
        Case pcmbMUKE_CODE
            
            If List_Disp_Proc Then
                Unload Me
            End If
    End Select

End Sub


Private Sub Combo_LostFocus(Index As Integer)

    Select Case Index
    
        Case pcmbMUKE_CODE
    
            Text(ptxMUKE_CODE).Text = Left(Right(Combo(Index).Text, 16), 8)
                
    
    End Select


End Sub

Private Sub Command_Click(Index As Integer)

Dim ans As Integer

    Select Case Index
        Case 7                              '再表示
            If List_Disp_Proc Then
                Unload Me
            End If
        Case 8                              'データ出力
        
            Beep
            ans = MsgBox("「出荷予定」データ出力しますか？", vbYesNo + vbQuestion, "確認入力")
            
            If ans = vbYes Then
                If OUTPUT_Proc() Then
                    Unload Me
                End If
            End If
        Case 11                             '終了
            Unload Me
        Case Else
            Beep
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
            Command(KeyCode - vbKeyF1).Value = True
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


    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                '出荷データファイル名取り込み
    If GetIni("FILE", "SYUKA_DATA", "SYS", c) Then
        Beep
        MsgBox "出荷データファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    SYUKA_DATA = Trim(c)
                                


                    '最大表示件数の獲得
    If GetIni(App.EXEName, "LISTMAX", "SYS", c) Then
        Beep
        MsgBox "最大表示件数の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    Max_Row = CInt(RTrim(c))
                                
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1030651.Caption = "「過日分」出荷確認（" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)

                                '向け先管理マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定ＯＰＥＮ
    If DEL_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If

'向け先設定
    If MTS_Set_Proc() Then
        Unload Me
    End If


'ｺﾝﾎﾞ初期設定
    
    Combo(pcmbCYU_KBN).AddItem "全て" & "   " & " "
    
    Combo(pcmbCYU_KBN).AddItem CYU_KBN_1 & "   " & CYU_KBN_TUK
    Combo(pcmbCYU_KBN).AddItem CYU_KBN_2 & "   " & CYU_KBN_SPO
    Combo(pcmbCYU_KBN).AddItem CYU_KBN_3 & "   " & CYU_KBN_HJU
'    Combo(pcmbCYU_KBN).AddItem CYU_KBN_4
    Combo(pcmbCYU_KBN).AddItem CYU_KBN_E & "   " & CYU_KBN_BOU
    Combo(pcmbCYU_KBN).AddItem CYU_KBN_T & "   " & CYU_KBN_KIN
    Combo(pcmbCYU_KBN).ListIndex = 0

    Combo(pcmbCYU_KBN).SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '向け先管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先管理マスタ")
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
                                            '出荷予定ＣＬＯＳＥ
    sts = BTRV(BtOpClose, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定")
        End If
    End If
    
    sts = BTRV(BtOpReset, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    'メニューより終了要求
    If JGYOBU_T(Index).CODE = " " Then
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1030651.Caption = "「過日分」出荷確認（" + RTrim(JGYOBU_T(Index).NAME) + ")"
    SubMenu(Index).Checked = True
    If Last_JGYOBU <> JGYOBU_T(Index).CODE Then
        Last_JGYOBU = JGYOBU_T(Index).CODE
        LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
        LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

    End If

End Sub

Private Function MTS_Set_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer
Dim Edit        As String

    MTS_Set_Proc = True
    
    Call Input_Lock
    
    
    Combo(pcmbMUKE_CODE).Clear
    
    
    Combo(pcmbMUKE_CODE).AddItem "全て　　　" & "   " & Space(16)
        
    
    com = BtOpGetFirst
    Do
        sts = BTRV(com, MTS_POS, MTSREC, Len(MTSREC), K1_MTS, Len(K1_MTS), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "向け先マスタ")
                Exit Function
        End Select
        
        Edit = StrConv(MTSREC.MUKE_DNAME, vbUnicode) & "   "
        Edit = Edit & StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode)
        
        
        Combo(pcmbMUKE_CODE).AddItem Edit
    
        com = BtOpGetNext
    
    Loop

    If Combo(pcmbMUKE_CODE).ListCount <= 0 Then
    Else
        Combo(pcmbMUKE_CODE).ListIndex = 0
    End If

    Call Input_UnLock

    MTS_Set_Proc = False
End Function


Private Function List_Disp_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim Row         As Long
    
Dim DEN_MAISU   As Long
Dim KAN_MAISU   As Long
    
Dim Skip_Flg    As Boolean
    
    List_Disp_Proc = True
                                    
                                    
    Me.MousePointer = vbArrowHourglass
                                    
'    Call Input_Lock
                                    
                                    
    If Trim(Right(Combo(pcmbCYU_KBN).Text, 1)) = "" Then
        TDBGrid1.Columns(ColCYU_KBN).Visible = True
    Else
        TDBGrid1.Columns(ColCYU_KBN).Visible = False
    End If
                                    
                                    
    If Trim(Right(Combo(pcmbMUKE_CODE).Text, 1)) = "" Then
        TDBGrid1.Columns(ColMUKE_CODE).Visible = True
    Else
        TDBGrid1.Columns(ColMUKE_CODE).Visible = False
    End If
                                    
                                    
                                    
                                    'テーブルリセット
    Set SYUKA = Nothing
                                    '出荷予定読み込み開始
    Call UniCode_Conv(K2_DEL_SYU.JGYOBU, Last_JGYOBU) '事業部
    
                                                    
    Call UniCode_Conv(K2_DEL_SYU.KEY_CYU_KBN, "")
    Call UniCode_Conv(K2_DEL_SYU.KEY_MUKE_CODE, "")
    Call UniCode_Conv(K2_DEL_SYU.KEY_SS_CODE, "")
    
    
    Row = Min_Row - 1
        
    DEN_MAISU = 0
    KAN_MAISU = 0
    
    
    
    com = BtOpGetGreaterEqual
    
''com = BtOpGetFirst
    Do
        sts = BTRV(com, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K2_DEL_SYU, Len(K2_DEL_SYU), 2)
    
        Select Case sts
            Case BtNoErr
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "出荷予定")
                List_Disp_Proc = SYS_ERR
                Exit Function
        End Select
                                '事業部 KEYﾌﾞﾚｰｸ
        If StrConv(DEL_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
            Exit Do
        End If
                                
        Skip_Flg = False
                                '注文区分 KEYﾌﾞﾚｰｸ
        
        If Trim(Right(Combo(pcmbCYU_KBN).Text, 1)) = "" Then
        Else
            If StrConv(DEL_SYUREC.CYU_KBN, vbUnicode) <> Right(Combo(pcmbCYU_KBN).Text, 1) Then
                Skip_Flg = True
            End If
        End If
                                '向け先 KEYﾌﾞﾚｰｸ
    
        If Trim(Right(Combo(pcmbMUKE_CODE).Text, 16)) = "" Then
        Else
            If StrConv(DEL_SYUREC.MUKE_CODE, vbUnicode) <> Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8) Or _
                StrConv(DEL_SYUREC.SS_CODE, vbUnicode) <> Right(Combo(pcmbMUKE_CODE).Text, 8) Then
                Skip_Flg = True
            End If
        End If
        
        If Not Skip_Flg Then
        
        
            DEN_MAISU = DEN_MAISU + 1
            
                                        '出荷完了
            If StrConv(DEL_SYUREC.KAN_KBN, vbUnicode) = KAN_KBN_FIN Then
                KAN_MAISU = KAN_MAISU + 1
            End If
            
            Row = Row + 1
            If Row > Max_Row Then
                Beep
                MsgBox "最大表示行数を超えました。"
                Exit Do
            End If
                    
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
        
        End If
        
        com = BtOpGetNext
        
        DoEvents
    Loop
    
                                
                                'DBテーブルリンク
    If DEN_MAISU < 1 Then
    Else
        SYUKA.QuickSort Min_Row, (SYUKA.UpperBound(1)), ColSort_Mark, XORDER_ASCEND, XTYPE_STRING, _
                                                        ColDEN_NO, XORDER_ASCEND, XTYPE_STRING
    End If
    
    Set TDBGrid1.Array = SYUKA
    
    TDBGrid1.Style.Locked = True
    
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    
    Text(ptxDEN_MAISU_JI).Text = Format(KAN_MAISU, "#,##0")
                                
    Text(ptxDEN_MAISU_YO).Text = Format(DEN_MAISU, "#,##0")
    
'    Call Input_UnLock
    
    
    Combo(pcmbMUKE_CODE).SetFocus
    
    Me.MousePointer = vbDefault
    
    List_Disp_Proc = False

    
End Function

Private Function OUTPUT_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer

    
Dim Ret         As Integer
    

Dim FileNo      As Integer
Dim fileName    As String
    
    
    OUTPUT_Proc = True
                                    
'    Call Input_Lock

    FileNo = FreeFile
    
    fileName = SYUKA_DATA
    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
    
    On Error GoTo Error_Proc
    
    Open (SYUKA_DATA) For Output As FileNo

    Write #FileNo, "注文区分：" & Combo(pcmbCYU_KBN).Text & "  向け先：" & Combo(pcmbMUKE_CODE).Text

    Write #FileNo, "ＩＤ№", "伝票№", "品番（外部）", "品番（内部）", "品名", "出荷予定数", "済み数", "検品", "伝票日付", Format(Now, "yyyy/mm/dd HH:mm:ss") & " 現在"

                                    '出荷予定読み込み開始
    Call UniCode_Conv(K2_DEL_SYU.JGYOBU, Last_JGYOBU) '事業部
    
                                                    '注文区分
    Call UniCode_Conv(K2_DEL_SYU.KEY_CYU_KBN, Right(Combo(pcmbCYU_KBN).Text, 1))
                                                    '向け先
    Call UniCode_Conv(K2_DEL_SYU.KEY_MUKE_CODE, Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8))
    Call UniCode_Conv(K2_DEL_SYU.KEY_SS_CODE, Right(Combo(pcmbMUKE_CODE).Text, 8))
    
    com = BtOpGetGreaterEqual
    Do
        sts = BTRV(com, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K2_DEL_SYU, Len(K2_DEL_SYU), 2)
    
        Select Case sts
            Case BtNoErr
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "出荷予定")
                OUTPUT_Proc = SYS_ERR
                Exit Function
        End Select
                                '事業部 KEYﾌﾞﾚｰｸ
        If StrConv(DEL_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
            Exit Do
        End If
                                '注文区分 KEYﾌﾞﾚｰｸ
        If StrConv(DEL_SYUREC.CYU_KBN, vbUnicode) <> Right(Combo(pcmbCYU_KBN).Text, 1) Then
            Exit Do
        End If
                                '向け先 KEYﾌﾞﾚｰｸ
        If StrConv(DEL_SYUREC.MUKE_CODE, vbUnicode) <> Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8) Or _
            StrConv(DEL_SYUREC.SS_CODE, vbUnicode) <> Right(Combo(pcmbMUKE_CODE).Text, 8) Then
            Exit Do
        End If
            
        Write #FileNo, StrConv(DEL_SYUREC.ID_NO, vbUnicode),
        Write #FileNo, StrConv(DEL_SYUREC.DEN_NO, vbUnicode),
        Write #FileNo, StrConv(DEL_SYUREC.HIN_NO, vbUnicode),
'2004        Write #FileNo, StrConv(del_syuREC.HIN_NAI, vbUnicode),
                                '品目マスタ読込み
        Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(DEL_SYUREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(DEL_SYUREC.HIN_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Write #FileNo, StrConv(ITEMREC.HIN_NAI, vbUnicode),
                Write #FileNo, StrConv(ITEMREC.HIN_NAME, vbUnicode),
            Case BtErrKeyNotFound
                Write #FileNo,
                Write #FileNo,
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        End Select
                                                                    '出荷予定数
        Write #FileNo, Format(CLng(StrConv(DEL_SYUREC.SURYO, vbUnicode)), "#,##0"),
                                                                    '出荷実績数
        Write #FileNo, Format(CLng(StrConv(DEL_SYUREC.JITU_SURYO, vbUnicode)), "#,##0"),
                                                                    '検品マーク
        If Len(Trim(StrConv(DEL_SYUREC.KENPIN_YMD, vbUnicode))) = 0 Then
                                '未検品
            Write #FileNo, KENPIN_OFF,
        Else
                                '検品済
            Write #FileNo, KENPIN_ON,
        End If
            
        Write #FileNo, Mid(StrConv(DEL_SYUREC.SYUKA_YMD, vbUnicode), 1, 4) & "/" _
                        & Mid(StrConv(DEL_SYUREC.SYUKA_YMD, vbUnicode), 5, 2) & "/" _
                        & Mid(StrConv(DEL_SYUREC.SYUKA_YMD, vbUnicode), 7, 2)
        
        com = BtOpGetNext
        
        DoEvents
    Loop

    Close #FileNo
    
'    Call Input_UnLock         '画面項目ロック解除
    
    Beep
    MsgBox "「" & fileName & "」は正常に出力されました。"

    Combo(pcmbMUKE_CODE).SetFocus
    
    OUTPUT_Proc = False
    
    Exit Function

Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "が使用中です。"
        OUTPUT_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        OUTPUT_Proc = True
    End If


End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1030651.MousePointer = vbHourglass

    Call Ctrl_Lock(F1030651)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1030651)


    F1030651.MousePointer = vbDefault

End Sub

Private Function Grid_Set_Proc(Row As Long) As Integer

Dim sts As Integer

    
    Grid_Set_Proc = True

    

    SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
    
    
    Select Case StrConv(DEL_SYUREC.CYU_KBN, vbUnicode)                '注文区分
            Case CYU_KBN_TUK
                SYUKA(Row, ColCYU_KBN) = CYU_KBN_1
            
            Case CYU_KBN_SPO
                SYUKA(Row, ColCYU_KBN) = CYU_KBN_2
            Case CYU_KBN_HJU
                SYUKA(Row, ColCYU_KBN) = CYU_KBN_3
            Case CYU_KBN_TOK
                SYUKA(Row, ColCYU_KBN) = CYU_KBN_4
            Case CYU_KBN_BOU
                SYUKA(Row, ColCYU_KBN) = CYU_KBN_E
            Case Else
                SYUKA(Row, ColCYU_KBN) = ""
    End Select
    
    
    Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(DEL_SYUREC.MUKE_CODE, vbUnicode))
    Call UniCode_Conv(K0_MTS.SS_CODE, StrConv(DEL_SYUREC.SS_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            SYUKA(Row, ColMUKE_CODE) = StrConv(DEL_SYUREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.MUKE_DNAME, vbUnicode)
        Case BtErrKeyNotFound
            SYUKA(Row, ColMUKE_CODE) = StrConv(DEL_SYUREC.MUKE_CODE, vbUnicode)
        Case Else
            Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
            Exit Function
    End Select
    
    
    
    
    
    SYUKA(Row, ColID_NO) = StrConv(DEL_SYUREC.ID_NO, vbUnicode)       'ＩＤ№
    SYUKA(Row, ColDEN_NO) = StrConv(DEL_SYUREC.DEN_NO, vbUnicode)     '伝票№
    SYUKA(Row, ColHIN_GAI) = StrConv(DEL_SYUREC.HIN_NO, vbUnicode)    '品番（外部）
    SYUKA(Row, ColHIN_NAI) = StrConv(DEL_SYUREC.HIN_NAI, vbUnicode)   '品番（内部）
                                                                    '品目マスタ読込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(DEL_SYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(DEL_SYUREC.HIN_NO, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            SYUKA(Row, ColHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    End Select
                                                                    '出荷予定数
    SYUKA(Row, ColYOTEI_QTY) = Format(CLng(StrConv(DEL_SYUREC.SURYO, vbUnicode)), "#,##0")
                                                                    '出荷実績数
    SYUKA(Row, ColFIX_QTY) = Format(CLng(StrConv(DEL_SYUREC.JITU_SURYO, vbUnicode)), "#,##0")
                                                                    '検品マーク
    If Len(Trim(StrConv(DEL_SYUREC.KENPIN_YMD, vbUnicode))) = 0 Then
                                '未検品
        SYUKA(Row, ColKENPIN_MARK) = KENPIN_OFF
    Else
                                '検品済
        SYUKA(Row, ColKENPIN_MARK) = KENPIN_ON
    End If
            
    SYUKA(Row, ColDEN_DT) = Mid(StrConv(DEL_SYUREC.SYUKA_YMD, vbUnicode), 1, 4) & "/" _
                            & Mid(StrConv(DEL_SYUREC.SYUKA_YMD, vbUnicode), 5, 2) & "/" _
                            & Mid(StrConv(DEL_SYUREC.SYUKA_YMD, vbUnicode), 7, 2)
    
    If CLng(StrConv(DEL_SYUREC.SURYO, vbUnicode)) > CLng(StrConv(DEL_SYUREC.JITU_SURYO, vbUnicode)) Then
                                '未出庫　または　出庫中
        SYUKA(Row, ColSort_Mark) = Sort_MISYUKO
    Else
                                '出庫完了　で　未検品
        If Len(Trim(StrConv(DEL_SYUREC.KENPIN_YMD, vbUnicode))) = 0 Then
            SYUKA(Row, ColSort_Mark) = Sort_SYUKOSUMI
        Else
            SYUKA(Row, ColSort_Mark) = Sort_KENPIN
        End If
    End If
    
    If Len(Trim(StrConv(DEL_SYUREC.PRINT_YMD, vbUnicode))) = 0 Then
            SYUKA(Row, ColPrint) = ""
    Else
            SYUKA(Row, ColPrint) = "○"
    End If
    
    
    If Trim(StrConv(DEL_SYUREC.INS_NOW, vbUnicode)) <> "" Then
        SYUKA(Row, ColIns_Date) = Mid(StrConv(DEL_SYUREC.INS_NOW, vbUnicode), 1, 4) & "/" _
                                    & Mid(StrConv(DEL_SYUREC.INS_NOW, vbUnicode), 5, 2) & "/" _
                                    & Mid(StrConv(DEL_SYUREC.INS_NOW, vbUnicode), 7, 2) & " " _
                                    & Mid(StrConv(DEL_SYUREC.INS_NOW, vbUnicode), 9, 2) & ":" _
                                    & Mid(StrConv(DEL_SYUREC.INS_NOW, vbUnicode), 11, 2) & ":" _
                                    & Mid(StrConv(DEL_SYUREC.INS_NOW, vbUnicode), 13, 2)

    Else
        SYUKA(Row, ColIns_Date) = ""
    End If
    
    
    If Trim(StrConv(DEL_SYUREC.KENPIN_YMD, vbUnicode)) <> "" Then
        SYUKA(Row, ColKENPIN_Date) = Mid(StrConv(DEL_SYUREC.KENPIN_YMD, vbUnicode), 1, 4) & "/" _
                                    & Mid(StrConv(DEL_SYUREC.KENPIN_YMD, vbUnicode), 5, 2) & "/" _
                                    & Mid(StrConv(DEL_SYUREC.KENPIN_YMD, vbUnicode), 7, 2) & " " _
                                    & Mid(StrConv(DEL_SYUREC.KENPIN_HMS, vbUnicode), 1, 2) & ":" _
                                    & Mid(StrConv(DEL_SYUREC.KENPIN_HMS, vbUnicode), 3, 2) & ":" _
                                    & Mid(StrConv(DEL_SYUREC.KENPIN_HMS, vbUnicode), 5, 2)

    Else
        SYUKA(Row, ColKENPIN_Date) = ""
    End If
    
    
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(DEL_SYUREC.KENPIN_TANTO_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(TANTOREC.TANTO_NAME, "")
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Exit Function
    End Select
    
    
    SYUKA(Row, ColKENPIN_TANTO) = StrConv(DEL_SYUREC.KENPIN_TANTO_CODE, vbUnicode) & " " & StrConv(TANTOREC.TANTO_NAME, vbUnicode)
    
    
    Grid_Set_Proc = False
End Function

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If


End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim sts As Integer
Dim i   As Integer

    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    Select Case Index
        Case ptxMUKE_CODE
            Call UniCode_Conv(K2_MTS.MUKE_CODE, Text(Index).Text)
            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K2_MTS, Len(K2_MTS), 2)
            Select Case sts
                Case BtNoErr
                    If Len(Trim(StrConv(MTSREC.SS_CODE, vbUnicode))) <> 0 Then
                        Beep
                        MsgBox "入力した項目はエラーです。(向け先コード)"
                        Exit Sub
                    End If
                                
                Case BtErrKeyNotFound
                                
                    Call UniCode_Conv(K3_MTS.SS_CODE, Text(Index).Text)
                                                        
                    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K3_MTS, Len(K3_MTS), 3)
                    Select Case sts
                        Case BtNoErr
                                        
                        Case BtErrKeyNotFound
                            Beep
                            MsgBox "入力した項目はエラーです。(向け先コード)"
                            Exit Sub
                                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                            Unload Me
                    End Select

                Case Else
                    Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                    Unload Me
            End Select


            For i = 0 To Combo(pcmbMUKE_CODE).ListCount - 1 '向け先
    
                If Right(Combo(pcmbMUKE_CODE).List(i), 16) = StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode) Then
                    Combo(pcmbMUKE_CODE).ListIndex = i
                    Exit For
                End If
            
    
            Next

            Combo(pcmbMUKE_CODE).SetFocus
    End Select

End Sub
