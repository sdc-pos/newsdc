VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PLN00601 
   Caption         =   "[商品化計画システム]勤務予定入力"
   ClientHeight    =   10695
   ClientLeft      =   2025
   ClientTop       =   -4470
   ClientWidth     =   14775
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
   ScaleHeight     =   10695
   ScaleWidth      =   14775
   StartUpPosition =   2  '画面の中央
   Begin TrueDBGrid80.TDBGrid TDBGrid2 
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   8880
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   2566
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).Locked=   -1  'True
      Splits(0).RecordSelectorWidth=   714
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4366"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
      Splits(0)._ColumnProps(4)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=5212"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5080"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=1561"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1429"
      Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(1)._UserFlags=   0
      Splits(1).Locked=   -1  'True
      Splits(1).RecordSelectorWidth=   714
      Splits(1).DividerColor=   14215660
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=3"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=4366"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
      Splits(1)._ColumnProps(4)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(6)=   "Column(1).Width=5212"
      Splits(1)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(8)=   "Column(1)._WidthInPix=5080"
      Splits(1)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(10)=   "Column(2).Width=1561"
      Splits(1)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(12)=   "Column(2)._WidthInPix=1429"
      Splits(1)._ColumnProps(13)=   "Column(2)._ColStyle=2"
      Splits(1)._ColumnProps(14)=   "Column(2).Order=3"
      Splits.Count    =   2
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
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
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(41)  =   ":id=28,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(42)  =   ":id=28,.fontname=ＭＳ ゴシック"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=1,.bold=0,.fontsize=900"
      _StyleDefs(47)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(48)  =   ":id=32,.fontname=ＭＳ ゴシック"
      _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14,.bold=0,.fontsize=900"
      _StyleDefs(50)  =   ":id=29,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(51)  =   ":id=29,.fontname=ＭＳ ゴシック"
      _StyleDefs(52)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(54)  =   "Splits(1).Style:id=47,.parent=1"
      _StyleDefs(55)  =   "Splits(1).CaptionStyle:id=56,.parent=4"
      _StyleDefs(56)  =   "Splits(1).HeadingStyle:id=48,.parent=2"
      _StyleDefs(57)  =   "Splits(1).FooterStyle:id=49,.parent=3"
      _StyleDefs(58)  =   "Splits(1).InactiveStyle:id=50,.parent=5"
      _StyleDefs(59)  =   "Splits(1).SelectedStyle:id=52,.parent=6"
      _StyleDefs(60)  =   "Splits(1).EditorStyle:id=51,.parent=7"
      _StyleDefs(61)  =   "Splits(1).HighlightRowStyle:id=53,.parent=8"
      _StyleDefs(62)  =   "Splits(1).EvenRowStyle:id=54,.parent=9"
      _StyleDefs(63)  =   "Splits(1).OddRowStyle:id=55,.parent=10"
      _StyleDefs(64)  =   "Splits(1).RecordSelectorStyle:id=57,.parent=11"
      _StyleDefs(65)  =   "Splits(1).FilterBarStyle:id=58,.parent=12"
      _StyleDefs(66)  =   "Splits(1).Columns(0).Style:id=62,.parent=47"
      _StyleDefs(67)  =   "Splits(1).Columns(0).HeadingStyle:id=59,.parent=48"
      _StyleDefs(68)  =   "Splits(1).Columns(0).FooterStyle:id=60,.parent=49"
      _StyleDefs(69)  =   "Splits(1).Columns(0).EditorStyle:id=61,.parent=51"
      _StyleDefs(70)  =   "Splits(1).Columns(1).Style:id=66,.parent=47,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(71)  =   ":id=66,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(72)  =   ":id=66,.fontname=ＭＳ ゴシック"
      _StyleDefs(73)  =   "Splits(1).Columns(1).HeadingStyle:id=63,.parent=48"
      _StyleDefs(74)  =   "Splits(1).Columns(1).FooterStyle:id=64,.parent=49"
      _StyleDefs(75)  =   "Splits(1).Columns(1).EditorStyle:id=65,.parent=51"
      _StyleDefs(76)  =   "Splits(1).Columns(2).Style:id=70,.parent=47,.alignment=1,.bold=0,.fontsize=900"
      _StyleDefs(77)  =   ":id=70,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(78)  =   ":id=70,.fontname=ＭＳ ゴシック"
      _StyleDefs(79)  =   "Splits(1).Columns(2).HeadingStyle:id=67,.parent=48,.bold=0,.fontsize=900"
      _StyleDefs(80)  =   ":id=67,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(81)  =   ":id=67,.fontname=ＭＳ ゴシック"
      _StyleDefs(82)  =   "Splits(1).Columns(2).FooterStyle:id=68,.parent=49"
      _StyleDefs(83)  =   "Splits(1).Columns(2).EditorStyle:id=69,.parent=51"
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
   Begin VB.CommandButton Command1 
      Caption         =   "更  新"
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
      Left            =   1920
      TabIndex        =   6
      ToolTipText     =   "商品化構成を読み込みます（Ｆ5）"
      Top             =   0
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "表　示"
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
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "商品化構成を読み込みます（Ｆ5）"
      Top             =   0
      Width           =   1380
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   7695
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   13573
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "担当者　　ｺｰﾄﾞ"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "担当者名"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   714
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1720"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1614"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3493"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3387"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1561"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1455"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(1)._UserFlags=   0
      Splits(1).AllowSizing=   -1  'True
      Splits(1).RecordSelectorWidth=   714
      Splits(1).AlternatingRowStyle=   -1  'True
      Splits(1).DividerColor=   14215660
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=3"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=1720"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=1614"
      Splits(1)._ColumnProps(4)=   "Column(0)._ColStyle=1"
      Splits(1)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(6)=   "Column(1).Width=3493"
      Splits(1)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(8)=   "Column(1)._WidthInPix=3387"
      Splits(1)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
      Splits(1)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(11)=   "Column(2).Width=1561"
      Splits(1)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(13)=   "Column(2)._WidthInPix=1455"
      Splits(1)._ColumnProps(14)=   "Column(2)._ColStyle=2"
      Splits(1)._ColumnProps(15)=   "Column(2).Order=3"
      Splits.Count    =   2
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowAddNew     =   -1  'True
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
      _StyleDefs(24)  =   "Splits(0).Style:id=67,.parent=1,.bgcolor=&HFFFF00&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=68,.parent=2,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(27)  =   ":id=68,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(28)  =   ":id=68,.fontname=ＭＳ ゴシック"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=69,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=72,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=71,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=74,.parent=9,.bgcolor=&HFFFF00&"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=75,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=94,.parent=67,.alignment=2"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=91,.parent=68"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=92,.parent=69"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=93,.parent=71"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=102,.parent=67,.locked=-1"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=99,.parent=68"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=100,.parent=69"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=101,.parent=71"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=98,.parent=67,.alignment=1"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=95,.parent=68"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=96,.parent=69"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=97,.parent=71"
      _StyleDefs(50)  =   "Splits(1).Style:id=13,.parent=1,.bgcolor=&HFFFF00&"
      _StyleDefs(51)  =   "Splits(1).CaptionStyle:id=22,.parent=4"
      _StyleDefs(52)  =   "Splits(1).HeadingStyle:id=14,.parent=2,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(53)  =   ":id=14,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(54)  =   ":id=14,.fontname=ＭＳ ゴシック"
      _StyleDefs(55)  =   "Splits(1).FooterStyle:id=15,.parent=3"
      _StyleDefs(56)  =   "Splits(1).InactiveStyle:id=16,.parent=5"
      _StyleDefs(57)  =   "Splits(1).SelectedStyle:id=18,.parent=6"
      _StyleDefs(58)  =   "Splits(1).EditorStyle:id=17,.parent=7"
      _StyleDefs(59)  =   "Splits(1).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(60)  =   "Splits(1).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFF00&"
      _StyleDefs(61)  =   "Splits(1).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(62)  =   "Splits(1).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(63)  =   "Splits(1).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(64)  =   "Splits(1).Columns(0).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(65)  =   "Splits(1).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(66)  =   "Splits(1).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(67)  =   "Splits(1).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(68)  =   "Splits(1).Columns(1).Style:id=32,.parent=13,.locked=-1"
      _StyleDefs(69)  =   "Splits(1).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(70)  =   "Splits(1).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(71)  =   "Splits(1).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(72)  =   "Splits(1).Columns(2).Style:id=46,.parent=13,.alignment=1"
      _StyleDefs(73)  =   "Splits(1).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(74)  =   "Splits(1).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(75)  =   "Splits(1).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(76)  =   "Named:id=33:Normal"
      _StyleDefs(77)  =   ":id=33,.parent=0"
      _StyleDefs(78)  =   "Named:id=34:Heading"
      _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(80)  =   ":id=34,.wraptext=-1"
      _StyleDefs(81)  =   "Named:id=35:Footing"
      _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(83)  =   "Named:id=36:Selected"
      _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(85)  =   "Named:id=37:Caption"
      _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(87)  =   "Named:id=38:HighlightRow"
      _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(89)  =   "Named:id=39:EvenRow"
      _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(91)  =   "Named:id=40:OddRow"
      _StyleDefs(92)  =   ":id=40,.parent=33"
      _StyleDefs(93)  =   "Named:id=41:RecordSelector"
      _StyleDefs(94)  =   ":id=41,.parent=34"
      _StyleDefs(95)  =   "Named:id=42:FilterBar"
      _StyleDefs(96)  =   ":id=42,.parent=33"
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
      Left            =   3480
      TabIndex        =   0
      ToolTipText     =   "商品化構成を保存します"
      Top             =   0
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "から2週間"
      Height          =   252
      Index           =   2
      Left            =   3240
      TabIndex        =   5
      Top             =   600
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "勤務予定日"
      Height          =   252
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   1212
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "読込"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "登録"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   2
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "PLN00601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Const ptxStart_Day% = 0


Dim PLN_O_HOURS         As New XArrayDB
Dim PLN_O_HOURS_G       As New XArrayDB



Private Const Min_Row% = 1              '最小行数
Dim Max_Row    As Integer               'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Max_Col As Long                 '最大列数

Private Const colTANTO_CODE% = 0        '担当者ｺｰﾄﾞ
Private Const colTANTO_NAME% = 1        '担当者名
Private Const colO_TIME% = 2            '勤務時間

Private List_Week   As Long             '表示するn週間
Private KADOU_RITU  As Double           '稼働率



Private Type Tanto_Tbl_tag
    TANTO_CODE  As String * 5
    O_Time()    As Double
End Type





Private Const LAST_UPDATE_DAY$ = "[PLN0060] 2011.10.07 12:30"

Private Sub Command1_Click(Index As Integer)

Dim sWk         As String
Dim i           As Long
Dim j           As Long


    Select Case Index



        Case 0          '読込み



            '取込みﾃﾞｰﾀ表示
            If List_Disp_Proc() Then
                Unload Me
            End If


            If PLN_O_HOURS.Count(1) > 0 Then
                Command1(1).Enabled = True
            Else
                Command1(1).Enabled = False
                Command1(0).SetFocus
            End If





        Case 1          '登録


            If Mid(Text1(ptxStart_Day).Text, 6, 5) <> Mid(TDBGrid1.Columns(2).Caption, 1, 5) Then
                MsgBox "指定した日付内容を表示後、再度「更新」を押下してください。"
                Exit Sub
            End If

            If Update_Proc() Then
                Unload Me
            End If


            '取込みﾃﾞｰﾀ表示
            If List_Disp_Proc() Then
                Unload Me
            End If


            If PLN_O_HOURS.Count(1) < 1 Then
                Command1(1).Enabled = False
                Command1(0).SetFocus
            End If


        Case 2          '終了

            Unload Me
    End Select



    Command1(Index).SetFocus


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    Select Case KeyCode
        
        Case vbKeyF12
            Unload Me
    End Select
    
    
    

End Sub

Private Sub Form_Load()


Dim c       As String * 128



    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[商品化計画システム]勤務予定入力画面", Me.hwnd, 0)
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




                                '表示期間取り込み
    If GetIni(App.EXEName, "WEEK", App.EXEName, c) Then
        List_Week = 2
    Else
        If Not IsNumeric(Trim(c)) Then
            List_Week = 2
        Else
            If Val(Trim(c)) < 1 Then
                List_Week = 1
            Else
                If Val(Trim(c)) > 8 Then
                    List_Week = 8
                Else
                    List_Week = Val(Trim(c))
                End If
            End If
        End If
    End If
    
    
    
    
    Call List_Make_Proc


                                '稼働率取り込み
    If GetIni(App.EXEName, "KADOU_RITU", App.EXEName, c) Then
        KADOU_RITU = 100
    Else
        If Not IsNumeric(Trim(c)) Then
            KADOU_RITU = 100
        Else
            If Val(Trim(c)) < 0 Then
                KADOU_RITU = 100
            Else
                KADOU_RITU = Val(Trim(c))
            End If
        End If
    End If








    PLN00601.Caption = PLN00601.Caption & " " & LAST_UPDATE_DAY
    
    Label1(2).Caption = "から" & List_Week & "週間"
    Text1(ptxStart_Day).Text = Format(Now, "YYYY/MM/DD")


                                
                                '担当者マスタ
    If TANTO_Open(BtOpenRead) Then
        Unload Me
    End If
                                
                                
                                '担当者別勤務時間データＯＰＥＮ
    If PLN_O_HOURS_Open(BtOpenNomal) Then
        Unload Me
    End If





End Sub



Private Sub Form_Unload(Cancel As Integer)
    
    
Dim sts     As Integer
    
    sts = BTRV(BtOpClose, PLN_O_HOURS_POS, PLN_O_HOURS_REC, Len(PLN_O_HOURS_REC), K0_PLN_O_HOURS, Len(K0_PLN_O_HOURS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者別勤務時間データ")
        End If
    End If
    
    
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, PLN_O_HOURS_POS, PLN_O_HOURS_REC, Len(PLN_O_HOURS_REC), K0_PLN_O_HOURS, Len(K0_PLN_O_HOURS), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If


    
    Set PLN00601 = Nothing



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



Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   「注文データ」登録処理
'----------------------------------------------------------------------------

Dim sts             As Integer
Dim ans             As Integer
    
Dim Upd_Com         As Integer
Dim Skip_Flg        As Integer
    
Dim INS_NOW         As String * 14

Dim Row             As Long
Dim i               As Long
Dim j               As Long

Dim Ing_Date        As String
Dim Start_Date      As String


    If PLN_O_HOURS.Count(1) = 0 Then
        Exit Function
    End If



    Update_Proc = True
    
    Call Input_Lock


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "担当者別勤務時間データ登録処理　処理開始！！", Me.hwnd, 0)

                                    
                                    
    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")
                                    
                                    'テーブルリセット
    
    Skip_Flg = True
    For Row = 1 To PLN_O_HOURS.UpperBound(1)
        
        Start_Date = Text1(ptxStart_Day).Text
        
        j = -1
        For i = colO_TIME To Max_Col - 1
        
            DoEvents
            
            Call UniCode_Conv(K0_PLN_O_HOURS.TANTO_CODE, PLN_O_HOURS(Row, colTANTO_CODE))
            j = j + 1
            Ing_Date = Format(DateAdd("d", j, Start_Date), "YYYY/MM/DD")
            Call UniCode_Conv(K0_PLN_O_HOURS.O_DATE, Format(Ing_Date, "YYYYMMDD"))
                    
        
            sts = BTRV(BtOpGetEqual, PLN_O_HOURS_POS, PLN_O_HOURS_REC, Len(PLN_O_HOURS_REC), K0_PLN_O_HOURS, Len(K0_PLN_O_HOURS), 0)
            Select Case sts
                Case BtNoErr
                    Upd_Com = BtOpUpdate
                Case BtErrKeyNotFound
                    Upd_Com = BtOpInsert
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    
                    Beep
                    ans = MsgBox("「担当者別勤務時間データ」他端末でデータ使用中です。<Y_SYUKA_TEI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                            
                        hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                            "担当者別勤務時間データ登録処理　キャンセル！！", Me.hwnd, 0)
                            
                            
                        Call Input_UnLock
                        Exit Function
                    End If
                Case Else
                        
                    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                        "担当者別勤務時間データ登録処理　異常停止！！", Me.hwnd, 0)
                        
                    Call Input_UnLock
                    Call File_Error(sts, BtOpGetEqual, "担当者別勤務時間データ")
                    Exit Function
            
            End Select
        
        
            If Upd_Com = BtOpInsert Then
                Call UniCode_Conv(PLN_O_HOURS_REC.TANTO_CODE, PLN_O_HOURS(Row, colTANTO_CODE))
                Call UniCode_Conv(PLN_O_HOURS_REC.O_DATE, Format(Ing_Date, "YYYYMMDD"))
                Call UniCode_Conv(PLN_O_HOURS_REC.FILLER, "")
                Call UniCode_Conv(PLN_O_HOURS_REC.INS_TANTO, App.EXEName)
                Call UniCode_Conv(PLN_O_HOURS_REC.Ins_DateTime, INS_NOW)
            End If
            
            If Not IsNumeric(PLN_O_HOURS(Row, colO_TIME + j)) Then
                PLN_O_HOURS(Row, colO_TIME + j) = "0.0"
            End If
            
            Call UniCode_Conv(PLN_O_HOURS_REC.O_Time, Format(PLN_O_HOURS(Row, colO_TIME + j), "00.0"))
            
            If Upd_Com = BtOpUpdate Then
                Call UniCode_Conv(PLN_O_HOURS_REC.UPD_TANTO, App.EXEName)
                Call UniCode_Conv(PLN_O_HOURS_REC.UPD_DATETIME, INS_NOW)
            End If
        
                    
            Do
                sts = BTRV(Upd_Com, PLN_O_HOURS_POS, PLN_O_HOURS_REC, Len(PLN_O_HOURS_REC), K0_PLN_O_HOURS, Len(K0_PLN_O_HOURS), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        
                        Beep
                        ans = MsgBox("「担当者別勤務時間データ」他端末でデータ使用中です。<Y_SYUKA_TEI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            
                            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                                "担当者別勤務時間データ登録処理　キャンセル！！", Me.hwnd, 0)
                            
                            Call Input_UnLock
                            Exit Function
                        End If
                    
                    Case Else
                        
                        hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                            "担当者別勤務時間データ登録処理　異常停止！！", Me.hwnd, 0)
                        
                        Call Input_UnLock
                        Call File_Error(sts, Upd_Com, "担当者別勤務時間データ")
                        Exit Function
                End Select
            
            Loop
            
    
        
        Next i

    Next Row


    Set TDBGrid1.Array = PLN_O_HOURS
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst



hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "注文データ登録処理　処理終了！！", Me.hwnd, 0)





    Update_Proc = False
    Call Input_UnLock
    Exit Function

Error_Proc:
    
    MsgBox "Err.Number= " & Err.Number & " " & Err.Description
    Call Input_UnLock

End Function

Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   「勤務予定データ」読込み処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer


Dim Row             As Long
Dim List_Day        As Long
Dim i               As Long
Dim j               As Long
Dim wkday           As Long


Dim Ing_Date        As String
Dim Start_Date      As String
Dim End_Date        As String

Dim Yobi            As Integer
Dim Yobi_NAME       As String

Dim Tanto_Tbl()     As Tanto_Tbl_tag
Dim Tanto_Tbl_G()   As Double


    List_Disp_Proc = True


    If Not IsDate(Text1(ptxStart_Day).Text) Then
        MsgBox "入力した項目はエラーです。(勤務予定日)"
        Text1(ptxStart_Day).SetFocus
        List_Disp_Proc = False
        Exit Function
    End If


'    If Text1(ptxStart_Day).Text < Format(Now, "YYYY/MM/DD") Then
'        MsgBox "入力した項目はエラーです。(勤務予定日 ＜　当日)"
'        Text1(ptxStart_Day).SetFocus
'        List_Disp_Proc = False
'        Exit Function
'    End If





    Call Input_Lock



hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "勤務予定入力　[検索]処理開始！！", Me.hwnd, 0)


    List_Day = List_Week * 7
    Start_Date = Text1(ptxStart_Day).Text
    End_Date = DateAdd("d", List_Day - 1, Start_Date)


    For i = 1 To List_Day
        
        
        Ing_Date = DateAdd("d", i - 1, Start_Date)
        Yobi = Weekday(Ing_Date)
        
        Select Case Yobi
            Case 1
                Yobi_NAME = "(" & "日" & ")"
            Case 2
                Yobi_NAME = "(" & "月" & ")"
            Case 3
                Yobi_NAME = "(" & "火" & ")"
            Case 4
                Yobi_NAME = "(" & "水" & ")"
            Case 5
                Yobi_NAME = "(" & "木" & ")"
            Case 6
                Yobi_NAME = "(" & "金" & ")"
            Case 7
                Yobi_NAME = "(" & "土" & ")"
        End Select
                
                
        
        TDBGrid1.Columns(i + 1).Caption = Right(Format(Ing_Date, "YYYY/MM/DD"), 5) & "   " & Yobi_NAME
        TDBGrid2.Columns(i + 1).Caption = Right(Format(Ing_Date, "YYYY/MM/DD"), 5) & "   " & Yobi_NAME

    Next i

    
    Call UniCode_Conv(K1_PLN_O_HOURS.O_DATE, Format(Start_Date, "YYYYMMDD"))
    Call UniCode_Conv(K1_PLN_O_HOURS.TANTO_CODE, "")
    i = -1
    com = BtOpGetGreater
    
    
    Do
        
        DoEvents
        
        sts = BTRV(com, PLN_O_HOURS_POS, PLN_O_HOURS_REC, Len(PLN_O_HOURS_REC), K1_PLN_O_HOURS, Len(K1_PLN_O_HOURS), 1)
        Select Case sts
            Case BtNoErr
            
                If StrConv(PLN_O_HOURS_REC.O_DATE, vbUnicode) > Format(End_Date, "YYYYMMDD") Then
                    Exit Do
                End If
            
                wkday = DateDiff("d", Start_Date, Mid(StrConv(PLN_O_HOURS_REC.O_DATE, vbUnicode), 1, 4) & "/" & _
                                                Mid(StrConv(PLN_O_HOURS_REC.O_DATE, vbUnicode), 5, 2) & "/" & _
                                                Mid(StrConv(PLN_O_HOURS_REC.O_DATE, vbUnicode), 7, 2))
                            
                            
                            
                If i = -1 Then
                    i = i + 1
                    ReDim Tanto_Tbl(0 To 0)
                    ReDim Tanto_Tbl(i).O_Time(0 To List_Day - 1)
                    
                    For j = 0 To UBound(Tanto_Tbl(i).O_Time)
                        Tanto_Tbl(i).O_Time(j) = 0
                    Next j
                    
                    
                    Tanto_Tbl(i).TANTO_CODE = StrConv(PLN_O_HOURS_REC.TANTO_CODE, vbUnicode)
                    
                    
                    Tanto_Tbl(i).O_Time(wkday) = CDbl(StrConv(PLN_O_HOURS_REC.O_Time, vbUnicode))
                    
                    
                Else
                
                    For i = 0 To UBound(Tanto_Tbl)
                        If Tanto_Tbl(i).TANTO_CODE = StrConv(PLN_O_HOURS_REC.TANTO_CODE, vbUnicode) Then
                            Exit For
                        End If
                    Next i
                
                    If i > UBound(Tanto_Tbl) Then
                        
                        ReDim Preserve Tanto_Tbl(0 To i)
                        ReDim Tanto_Tbl(i).O_Time(0 To List_Day - 1)
                        
                        For j = 0 To UBound(Tanto_Tbl(i).O_Time)
                            Tanto_Tbl(i).O_Time(j) = 0
                        Next j
                        
                        Tanto_Tbl(i).TANTO_CODE = StrConv(PLN_O_HOURS_REC.TANTO_CODE, vbUnicode)
                
                
                
                    End If
                
                
                    If Not IsNumeric(StrConv(PLN_O_HOURS_REC.O_Time, vbUnicode)) Then
                        Tanto_Tbl(i).O_Time(wkday) = 0#
                    Else
                        Tanto_Tbl(i).O_Time(wkday) = CDbl(StrConv(PLN_O_HOURS_REC.O_Time, vbUnicode))
                    End If
                End If
            
            
                            
            
            
            
            
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "勤務予定データ")
                Call Input_UnLock
                Exit Function
        End Select
        
        com = BtOpGetNext
        
    Loop


    com = BtOpGetFirst
    Do
        DoEvents
    
    
        sts = BTRV(com, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
            
                            
                If Trim(StrConv(TANTOREC.KUBUN, vbUnicode)) = "" Then
                Else
                    If i = -1 Then
                        i = i + 1
                        ReDim Tanto_Tbl(0 To 0)
                        ReDim Tanto_Tbl(i).O_Time(0 To List_Day - 1)
                        
                        For j = 0 To UBound(Tanto_Tbl(i).O_Time)
                            Tanto_Tbl(i).O_Time(j) = 0
                        Next j
                        
                        
                        Tanto_Tbl(i).TANTO_CODE = StrConv(TANTOREC.TANTO_CODE, vbUnicode)
                        
                        
                    
                    
                    Else
                
                        For i = 0 To UBound(Tanto_Tbl)
                            If Tanto_Tbl(i).TANTO_CODE = StrConv(TANTOREC.TANTO_CODE, vbUnicode) Then
                                Exit For
                            End If
                        Next i
                    
                        If i > UBound(Tanto_Tbl) Then
                            
                            ReDim Preserve Tanto_Tbl(0 To i)
                            ReDim Tanto_Tbl(i).O_Time(0 To List_Day - 1)
                            
                            Tanto_Tbl(i).TANTO_CODE = StrConv(TANTOREC.TANTO_CODE, vbUnicode)
                    
                            For j = 0 To UBound(Tanto_Tbl(i).O_Time)
                                Tanto_Tbl(i).O_Time(j) = 0
                            Next j
                    
                    
                        End If
                    
                    End If
                
                End If
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "担当者マスタ")
                Call Input_UnLock
                Exit Function
        End Select
        
        com = BtOpGetNext
    
    
    Loop


    Set PLN_O_HOURS = Nothing
    Row = Min_Row - 1
    
    
    If i = -1 Then
    Else
        For i = 0 To UBound(Tanto_Tbl)
        
            Row = Row + 1
            PLN_O_HOURS.ReDim Min_Row, Row, Min_Col, Max_Col
        
                    
            PLN_O_HOURS(Row, colTANTO_CODE) = Tanto_Tbl(i).TANTO_CODE
            
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Tanto_Tbl(i).TANTO_CODE)
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    PLN_O_HOURS(Row, colTANTO_NAME) = Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
                Case BtErrKeyNotFound
                    PLN_O_HOURS(Row, colTANTO_NAME) = "担当者マスタ未登録"
                Case Else
                    Call File_Error(sts, BtOpInsert, "担当者マスタ")
                    Call Input_UnLock
                    Exit Function
            End Select
        
            For j = 0 To UBound(Tanto_Tbl(i).O_Time)
                
                If Tanto_Tbl(i).O_Time(j) <> 0 Then
                    PLN_O_HOURS(Row, colO_TIME + j) = Format(Tanto_Tbl(i).O_Time(j), "0.0")
                Else
                    PLN_O_HOURS(Row, colO_TIME + j) = ""
                End If
            
            Next j
        
        Next i
    End If

    Set TDBGrid1.Array = PLN_O_HOURS
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst


    If i = -1 Then
    Else
        ReDim Tanto_Tbl_G(0 To List_Day - 1)
        For i = 0 To UBound(Tanto_Tbl_G)
            Tanto_Tbl_G(i) = 0
        Next i
    
        For i = 0 To UBound(Tanto_Tbl)
            For j = 0 To UBound(Tanto_Tbl(i).O_Time)
            
                Tanto_Tbl_G(j) = Tanto_Tbl_G(j) + Tanto_Tbl(i).O_Time(j)
            
            Next j
        
        Next i
    End If

    Set PLN_O_HOURS_G = Nothing
    PLN_O_HOURS_G.ReDim Min_Row, 2, Min_Col, Max_Col

    PLN_O_HOURS_G(1, colTANTO_NAME) = "合計(H)"
    
    If i = -1 Then
    Else
        For i = 0 To UBound(Tanto_Tbl_G)
            
            If Tanto_Tbl_G(i) = 0 Then
                PLN_O_HOURS_G(1, colO_TIME + i) = ""
            Else
                PLN_O_HOURS_G(1, colO_TIME + i) = Format(Tanto_Tbl_G(i), "#,##0.0")
            End If
        Next i
    End If
    
    
    
    
    
    PLN_O_HOURS_G(2, colTANTO_NAME) = "合計(m) 稼働率=" & Format(KADOU_RITU, "0.0") & "%"


    If i = -1 Then
    Else
        For i = 0 To UBound(Tanto_Tbl_G)
            
            If Not IsNumeric(Tanto_Tbl_G(i)) Then
                PLN_O_HOURS_G(2, colO_TIME + i) = ""
            Else
                If Tanto_Tbl_G(i) = 0 Then
                    PLN_O_HOURS_G(2, colO_TIME + i) = ""
                Else
                    PLN_O_HOURS_G(2, colO_TIME + i) = Format(Round((Tanto_Tbl_G(i) * 60) / (KADOU_RITU / 100), 2), "#,##0.0")
                End If
            End If
        Next i
    End If



    Set TDBGrid2.Array = PLN_O_HOURS_G
    TDBGrid2.ReBind
    
    TDBGrid2.Update
    TDBGrid2.MoveFirst










hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "勤務予定入力　[検索]処理終了！！", Me.hwnd, 0)



    Call Input_UnLock


    List_Disp_Proc = False

End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    PLN00601.MousePointer = vbHourglass

    Call Ctrl_Lock(PLN00601)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(PLN00601)


    PLN00601.MousePointer = vbDefault

End Sub


Private Sub List_Make_Proc()
'----------------------------------------------------------------------------
'                   Grid行作成
'----------------------------------------------------------------------------

Dim i       As Long
Dim NISUU   As Long


Dim TColumn     As TrueDBGrid80.Column
 



    

    NISUU = List_Week * 7


    For i = 3 To NISUU + 1
        
        Set TColumn = TDBGrid1.Columns.Add(i)
        With TColumn
            .Visible = True
            .Caption = ""
            .Width = TDBGrid1.Columns(2).Width
            .Font.NAME = "ＭＳ ゴシック"
            .Font.Size = 9
            .HeadFont.NAME = "ＭＳ ゴシック"
            .HeadFont.Size = 9
            .Alignment = dbgRight
        End With
    
    
        Set TColumn = TDBGrid2.Columns.Add(i)
        With TColumn
            .Visible = True
            .Caption = ""
            .Width = TDBGrid2.Columns(2).Width
            .Font.NAME = "ＭＳ ゴシック"
            .Font.Size = 9
            .HeadFont.NAME = "ＭＳ ゴシック"
            .HeadFont.Size = 9
            .Alignment = dbgRight
        End With
    
    Next i

    Max_Col = NISUU + 2

End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    
Dim sts     As Integer
Dim i       As Long
Dim j       As Long
    
    Set TDBGrid1.Array = PLN_O_HOURS
    TDBGrid1.Update
    
    
    
    If TDBGrid1.Bookmark = Null Then
        Exit Sub
    End If
    
    If TDBGrid1.Bookmark <= 0 Then
        Exit Sub
    End If


    Select Case ColIndex
    
        Case colTANTO_CODE
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, PLN_O_HOURS(TDBGrid1.Bookmark, ColIndex))
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    PLN_O_HOURS(TDBGrid1.Bookmark, colTANTO_NAME) = "担当者マスタ未登録"
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Exit Sub
            End Select
                
        Case colTANTO_NAME
    
        Case Else
            If Trim(PLN_O_HOURS(TDBGrid1.Bookmark, ColIndex)) = "" Then
            Else
                If Not IsNumeric(PLN_O_HOURS(TDBGrid1.Bookmark, ColIndex)) Then
                    MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(勤務時間＝数値)"
                    Exit Sub
                Else
                    PLN_O_HOURS(TDBGrid1.Bookmark, ColIndex) = Format(Val(PLN_O_HOURS(TDBGrid1.Bookmark, ColIndex)), "#0.0")
                    If Right(PLN_O_HOURS(TDBGrid1.Bookmark, ColIndex), 1) = "0" Or Right(PLN_O_HOURS(TDBGrid1.Bookmark, ColIndex), 1) = "5" Then
                    Else
                        MsgBox "[" & Format(TDBGrid1.Bookmark, "0") & "]行目 入力した項目はエラーです。(勤務時間＝0.5単位)"
                    End If
                End If
            End If
    
                
    
    End Select
    
    
    Set TDBGrid1.Array = PLN_O_HOURS
        
    
    TDBGrid1.Refresh
    TDBGrid1.Update
    
    
    TDBGrid1.SetFocus
    
    
    '-----------------------------------    合計再集計
    Set PLN_O_HOURS_G = Nothing
    PLN_O_HOURS_G.ReDim Min_Row, 2, Min_Col, Max_Col

    PLN_O_HOURS_G(1, colTANTO_NAME) = "合計(H)"
    
    For i = 1 To PLN_O_HOURS.UpperBound(1)
        For j = colO_TIME To Max_Col
        
            If IsNumeric(PLN_O_HOURS(i, j)) Then
                PLN_O_HOURS_G(1, j) = Format(Val(PLN_O_HOURS_G(1, j)) + Val(PLN_O_HOURS(i, j)), "#,##0.0")
            End If
        Next j
    Next i
    
    
    
    
    
    
    PLN_O_HOURS_G(2, colTANTO_NAME) = "合計(m) 稼働率=" & Format(KADOU_RITU, "0.0") & "%"


    For j = 2 To Max_Col
        If Val(PLN_O_HOURS_G(1, j)) = 0 Then
        Else
            PLN_O_HOURS_G(2, j) = Format(Round((Val(PLN_O_HOURS_G(1, j)) * 60) / (KADOU_RITU / 100), 2), "#,##0.0")
        End If
    Next j




    Set TDBGrid2.Array = PLN_O_HOURS_G
    TDBGrid2.ReBind
    
    TDBGrid2.Update
    TDBGrid2.MoveFirst
    
    
End Sub

Private Sub TDBGrid1_BeforeInsert(Cancel As Integer)
    PLN_O_HOURS.ReDim Min_Row, PLN_O_HOURS.Count(1), Min_Col, Max_Col

    Command1(1).Enabled = True

End Sub

