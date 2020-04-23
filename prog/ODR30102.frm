VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form ODR30102 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "子部品注文情報"
   ClientHeight    =   6630
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   15270
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   6930
      ScaleHeight     =   195
      ScaleWidth      =   270
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   6840
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   780
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   780
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "キャンセル"
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
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   1800
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid2 
      Height          =   3435
      Left            =   780
      TabIndex        =   6
      Top             =   2520
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   6059
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "注文№"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "仕入先"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "仕入先名"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "注文数"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "仕入残"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "希望納期"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "回答納期"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "ＫＥＹ項目"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "受入日"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "使用月"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "完F"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2117"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1984"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1402"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1270"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(10)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2831"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2699"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=8196"
      Splits(0)._ColumnProps(16)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=2461"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2328"
      Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=8194"
      Splits(0)._ColumnProps(22)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(4).Width=2461"
      Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2328"
      Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(28)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(30)=   "Column(5).Width=2461"
      Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=2328"
      Splits(0)._ColumnProps(33)=   "Column(5)._ColStyle=1"
      Splits(0)._ColumnProps(34)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(36)=   "Column(6).Width=2461"
      Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=2328"
      Splits(0)._ColumnProps(39)=   "Column(6)._ColStyle=8193"
      Splits(0)._ColumnProps(40)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(42)=   "Column(7).Width=2778"
      Splits(0)._ColumnProps(43)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(45)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(47)=   "Column(8).Width=2461"
      Splits(0)._ColumnProps(48)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(8)._WidthInPix=2328"
      Splits(0)._ColumnProps(50)=   "Column(8)._ColStyle=1"
      Splits(0)._ColumnProps(51)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(52)=   "Column(9).Width=2117"
      Splits(0)._ColumnProps(53)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(9)._WidthInPix=1984"
      Splits(0)._ColumnProps(55)=   "Column(9)._ColStyle=1"
      Splits(0)._ColumnProps(56)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(57)=   "Column(10).Width=529"
      Splits(0)._ColumnProps(58)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(10)._WidthInPix=397"
      Splits(0)._ColumnProps(60)=   "Column(10).Order=11"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageFooterFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   1
      Caption         =   "子部品　注文データ"
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      CellTipsWidth   =   0
      InsertMode      =   0   'False
      DeadAreaBackColor=   -2147483643
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFF00&,.bold=0,.fontsize=1125"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ＭＳ ゴシック"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&H80FF00&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFF80&"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=87,.parent=1,.bgcolor=&H80FF80&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=96,.parent=4,.bgcolor=&H80FF00&"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=89,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=91,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=94,.parent=9,.bgcolor=&H80FF80&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=95,.parent=10,.bgcolor=&H80FF00&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=106,.parent=87,.alignment=3,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=103,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=104,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=105,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=16,.parent=87,.locked=-1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=13,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=14,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=15,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=118,.parent=87,.alignment=3,.locked=-1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=115,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=116,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=117,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=122,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=119,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=120,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=121,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=20,.parent=87,.alignment=1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=17,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=18,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=19,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=24,.parent=87,.alignment=2"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=21,.parent=88"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=22,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=23,.parent=91"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=130,.parent=87,.alignment=2,.locked=-1"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=127,.parent=88"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=128,.parent=89"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=129,.parent=91"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=138,.parent=87"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=135,.parent=88"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=136,.parent=89"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=137,.parent=91"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=28,.parent=87,.alignment=2"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=25,.parent=88"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=26,.parent=89"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=27,.parent=91"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=32,.parent=87,.alignment=2"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=29,.parent=88"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=30,.parent=89"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=31,.parent=91"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=46,.parent=87"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=43,.parent=88"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=44,.parent=89"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=45,.parent=91"
      _StyleDefs(80)  =   "Named:id=33:Normal"
      _StyleDefs(81)  =   ":id=33,.parent=0"
      _StyleDefs(82)  =   "Named:id=34:Heading"
      _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(84)  =   ":id=34,.wraptext=-1"
      _StyleDefs(85)  =   "Named:id=35:Footing"
      _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(87)  =   "Named:id=36:Selected"
      _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(89)  =   "Named:id=37:Caption"
      _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(91)  =   "Named:id=38:HighlightRow"
      _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(93)  =   "Named:id=39:EvenRow"
      _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&H80FF00&"
      _StyleDefs(95)  =   "Named:id=40:OddRow"
      _StyleDefs(96)  =   ":id=40,.parent=33"
      _StyleDefs(97)  =   "Named:id=41:RecordSelector"
      _StyleDefs(98)  =   ":id=41,.parent=34"
      _StyleDefs(99)  =   "Named:id=42:FilterBar"
      _StyleDefs(100) =   ":id=42,.parent=33"
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   855
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1320
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   1508
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "子部品"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "子部品名"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "使用数"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "必要数"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "月初在庫"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "不足数"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "注文数"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "仕入残"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "ロット"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "仕入先"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "仕入先名"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "仕入単価"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "希望納期"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "回答納期"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   14
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=14"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1640"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2566"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2434"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(11)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1402"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1270"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=8194"
      Splits(0)._ColumnProps(17)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=1773"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1640"
      Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=8194"
      Splits(0)._ColumnProps(23)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=1773"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1640"
      Splits(0)._ColumnProps(28)=   "Column(4)._ColStyle=8194"
      Splits(0)._ColumnProps(29)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=1588"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1455"
      Splits(0)._ColumnProps(34)=   "Column(5)._ColStyle=8194"
      Splits(0)._ColumnProps(35)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=1588"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1455"
      Splits(0)._ColumnProps(40)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(41)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(43)=   "Column(7).Width=1588"
      Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=1455"
      Splits(0)._ColumnProps(46)=   "Column(7)._ColStyle=8194"
      Splits(0)._ColumnProps(47)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(49)=   "Column(8).Width=1588"
      Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1455"
      Splits(0)._ColumnProps(52)=   "Column(8)._ColStyle=8194"
      Splits(0)._ColumnProps(53)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(55)=   "Column(9).Width=1402"
      Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=1270"
      Splits(0)._ColumnProps(58)=   "Column(9)._ColStyle=8196"
      Splits(0)._ColumnProps(59)=   "Column(9).AllowFocus=0"
      Splits(0)._ColumnProps(60)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(61)=   "Column(10).Width=2117"
      Splits(0)._ColumnProps(62)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(10)._WidthInPix=1984"
      Splits(0)._ColumnProps(64)=   "Column(10)._ColStyle=8196"
      Splits(0)._ColumnProps(65)=   "Column(10).AllowFocus=0"
      Splits(0)._ColumnProps(66)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(67)=   "Column(11).Width=1773"
      Splits(0)._ColumnProps(68)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(69)=   "Column(11)._WidthInPix=1640"
      Splits(0)._ColumnProps(70)=   "Column(11)._ColStyle=8194"
      Splits(0)._ColumnProps(71)=   "Column(11).AllowFocus=0"
      Splits(0)._ColumnProps(72)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(73)=   "Column(12).Width=2302"
      Splits(0)._ColumnProps(74)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(75)=   "Column(12)._WidthInPix=2170"
      Splits(0)._ColumnProps(76)=   "Column(12)._ColStyle=1"
      Splits(0)._ColumnProps(77)=   "Column(12).AllowFocus=0"
      Splits(0)._ColumnProps(78)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(79)=   "Column(13).Width=2302"
      Splits(0)._ColumnProps(80)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(81)=   "Column(13)._WidthInPix=2170"
      Splits(0)._ColumnProps(82)=   "Column(13)._ColStyle=8193"
      Splits(0)._ColumnProps(83)=   "Column(13).AllowFocus=0"
      Splits(0)._ColumnProps(84)=   "Column(13).Order=14"
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
      Caption         =   "子部品　所要・注文情報"
      AllowArrows     =   0   'False
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      CellTipsWidth   =   0
      InsertMode      =   0   'False
      DeadAreaBackColor=   -2147483643
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H40FF00&,.bold=0,.fontsize=1125"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ＭＳ ゴシック"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H80000005&"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&H80FF00&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFF80&"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=87,.parent=1,.bgcolor=&H80FF00&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=96,.parent=4,.namedParent=37,.bgcolor=&H80FF00&"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=89,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=91,.parent=7,.bgcolor=&H80000005&"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=94,.parent=9,.namedParent=39"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=95,.parent=10,.namedParent=40,.bgcolor=&H80FF00&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=102,.parent=87,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=99,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=100,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=101,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=110,.parent=87,.locked=-1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=107,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=108,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=109,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=114,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=111,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=112,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=113,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=118,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=115,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=116,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=117,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=55,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=52,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=53,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=54,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=126,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=123,.parent=88"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=124,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=125,.parent=91"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=21,.parent=87,.alignment=1"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=18,.parent=88"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=19,.parent=89"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=20,.parent=91"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=25,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=22,.parent=88"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=23,.parent=89"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=24,.parent=91"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=17,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=14,.parent=88"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=15,.parent=89"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=16,.parent=91"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=130,.parent=87,.alignment=3,.locked=-1"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=127,.parent=88"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=128,.parent=89"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=129,.parent=91"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=29,.parent=87,.alignment=3,.locked=-1"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=26,.parent=88"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=27,.parent=89"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=28,.parent=91"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=43,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=30,.parent=88"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=31,.parent=89"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=32,.parent=91"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=47,.parent=87,.alignment=2"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=44,.parent=88"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=45,.parent=89"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=46,.parent=91"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=51,.parent=87,.alignment=2,.locked=-1"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=48,.parent=88"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=49,.parent=89"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=50,.parent=91"
      _StyleDefs(92)  =   "Named:id=33:Normal"
      _StyleDefs(93)  =   ":id=33,.parent=0"
      _StyleDefs(94)  =   "Named:id=34:Heading"
      _StyleDefs(95)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(96)  =   ":id=34,.wraptext=-1"
      _StyleDefs(97)  =   "Named:id=35:Footing"
      _StyleDefs(98)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(99)  =   "Named:id=36:Selected"
      _StyleDefs(100) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(101) =   "Named:id=37:Caption"
      _StyleDefs(102) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(103) =   "Named:id=38:HighlightRow"
      _StyleDefs(104) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(105) =   "Named:id=39:EvenRow"
      _StyleDefs(106) =   ":id=39,.parent=33,.bgcolor=&H80FF80&"
      _StyleDefs(107) =   "Named:id=40:OddRow"
      _StyleDefs(108) =   ":id=40,.parent=33,.bgcolor=&H40FF00&"
      _StyleDefs(109) =   "Named:id=41:RecordSelector"
      _StyleDefs(110) =   ":id=41,.parent=34"
      _StyleDefs(111) =   "Named:id=42:FilterBar"
      _StyleDefs(112) =   ":id=42,.parent=33"
      _StyleDefs(113) =   "Named:id=13:LockItem"
      _StyleDefs(114) =   ":id=13,.parent=39"
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "使用月"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   6060
      TabIndex        =   5
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Lab_Dsp 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2160
      TabIndex        =   4
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "担当者"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   840
      Width           =   720
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "キャンセル"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "画面印刷"
         Index           =   1
      End
   End
End
Attribute VB_Name = "ODR30102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'コンボ用添字
'Private Const pcmbHBUN = 0

'テキスト用添字
Private Const ptxTOP% = 0
Private Const ptxLAST% = 1

Private Const ptxTANTO_CD% = 0
Private Const ptxUSE_YY% = 1

Private Const ptxSUM_QTY% = 2

'ラベル用添字
Private Const plabTANTO_NM% = 0

'コマンドボタン用添字
'Private Const FuncCOR% = 0       '更新
Private Const FuncEND% = 0       '終了

'ListBox添字
Private Const plst_DISP% = 0     '表示用データ　Sort順＆Key

'グリッド更新マーク
Dim Grid_Cor_M      As Integer


'グリッド用定義
Private ORDR_GRID   As New XArrayDB
Private BUNNO_GRID   As New XArrayDB


Private Const Min_Row% = 1              '最小行数
'Private Max_Row As Long                 '最大表示行数
Private Const Max_Row = 9999            '最大行数


Private Const Min_Col% = 0                  '最小列数
Private Const Max_Col% = 13                 '最大列数

Private Const Col_ITEM% = 0                 '子部品コード
Private Const Col_ITEM_NM% = 1              '子部品名
Private Const Col_USE_QTY% = 2              '使用数量
Private Const Col_MRP_QTY% = 3              '必要数
Private Const Col_ZAI_QTY% = 4              '月初在庫
Private Const Col_FUSOKU% = 5               '不足数
Private Const Col_ORDR_QTY% = 6             '注文数
Private Const Col_ZAN_QTY% = 7              '仕入残
Private Const Col_LOT_QTY% = 8              'ロット数
Private Const Col_SECT_CD% = 9              '仕入先
Private Const Col_SECT_NM% = 10             '仕入先名
Private Const Col_TANKA% = 11               '仕入単価
Private Const Col_KIBOU_DT% = 12            '希望納期
Private Const Col_KAITO_DT% = 13            '回答納期


'Private Const Col2_DEL% = 0                '削除指示
Private Const Col2_ORDR_NO% = 0             '注文№
Private Const Col2_SECT% = 1                '仕入先
Private Const Col2_SECT_NM% = 2             '仕入先名
Private Const Col2_ORDR_QTY% = 3            '発注数
Private Const Col2_ZAN_QTY% = 4             '仕入残
Private Const Col2_KIBOU_DT% = 5            '希望納期
Private Const Col2_KAITO_DT% = 6            '回答納期
Private Const Col2_KKEY% = 7                'Key
Private Const Col2_KAN_DT% = 8              '受入日
Private Const Col2_USE_DT% = 9              '使用月
Private Const Col2_KAN_F% = 10              '完了F



Dim Mode        As Boolean
Dim row         As Long                 '対象　行

Dim Cor_Row     As Long                 'カレント行

Dim Init_F      As Integer
Dim W_SUM       As Double


Private Function ERR_CHK(Index As Integer)
Dim sts         As Integer
Dim yn          As Integer

Dim W_STR       As String


    ERR_CHK = True
    
                        '入力文字数チェック
    'If LenB(StrConv(Text1(Index), vbFromUnicode)) > Text1(Index).MaxLength Then
    '    MsgBox "入力した項目は（桁あふれエラー）です。", vbExclamation
    '    Exit Function
    'End If
    
    Select Case Index
        Case ptxTANTO_CD%
            Lab_Dsp(plabTANTO_NM) = ""


        Case ptxUSE_YY%
            
            
    End Select
    
    
    ERR_CHK = False
End Function



Private Function Grid_Err_Chk(Index As Integer, W_Aft As String)
'       グリッド入力内容エラーチェック
'
Dim sts         As Integer
Dim yn          As Integer

Dim W_STR       As String
Dim W_QTY       As Double



    Grid_Err_Chk = True
    
    Select Case Index
        Case Col2_ORDR_NO%                  '注文№
        
        Case Col2_SECT%                     '仕入先
                       
        Case Col2_ORDR_QTY%                 '注文数量
        
        
        Case Col2_ZAN_QTY%                  '仕入残
        
        
        Case Col2_KIBOU_DT%                 '希望納期
        
        
        Case Col2_KAITO_DT%                 '回答納期
    
        
        Case Else
        
        
    End Select
    
    DoEvents
    
    If Trim(W_Aft) <> "" Then
        Select Case Index
                 '注文数量       '希望納期
            Case Col2_ORDR_QTY%, Col2_KIBOU_DT%
                W_STR = Trim(BUNNO_GRID(Cor_Row, Col2_ORDR_QTY%))
                'If W_Str = "" Then
                '    MsgBox "親部品　未指定エラー！", vbExclamation
                '    Exit Function
                'End If
                
            Case Else
    
        End Select
    End If
    

    Grid_Err_Chk = False

End Function

Private Function Data_Disp()
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

'Dim X_i         As Long

'Dim W_Key       As String

Dim W_STR       As String
Dim W_Date      As String
'Dim cnt         As Integer

    Data_Disp = True
    
    row = Min_Row - 1
    Call Input_Lock                             '画面項目ロック
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "親部品　分納情報　検索・表示中。　＜Data_Disp＞", Me.hwnd, 0)
    DoEvents
    
    Set ORDR_GRID = Nothing
    Set BUNNO_GRID = Nothing
    
    '基情報のグリッド表示
    row = 1
    ORDR_GRID.ReDim Min_Row, row, Min_Col, Max_Col
    
    ORDR_GRID(row, Col_ITEM%) = Trim(DIS_ITEM)            '子部品コード
    ORDR_GRID(row, Col_ITEM_NM%) = Trim(DIS_ITEM_NM)      '子部品名
    ORDR_GRID(row, Col_USE_QTY%) = Trim(DIS_USE_QTY)      '使用数量
    
    ORDR_GRID(row, Col_MRP_QTY%) = Trim(DIS_MRP_QTY)      '必要数
    ORDR_GRID(row, Col_ZAI_QTY%) = Trim(DIS_ZAI_QTY)      '月初在庫
    ORDR_GRID(row, Col_FUSOKU%) = Trim(DIS_FUSOKU)        '不足数
    ORDR_GRID(row, Col_ORDR_QTY%) = Trim(DIS_ORDR_QTY)    '注文数
    ORDR_GRID(row, Col_ZAN_QTY%) = Trim(DIS_ZAN_QTY)      '仕入残
    
    ORDR_GRID(row, Col_LOT_QTY%) = Trim(DIS_LOT_QTY)      'ロット数
    ORDR_GRID(row, Col_SECT_CD%) = Trim(DIS_SECT_CD)      '仕入先
    ORDR_GRID(row, Col_SECT_NM%) = Trim(DIS_SECT_NM)      '仕入先名
    ORDR_GRID(row, Col_TANKA%) = Trim(DIS_TANKA)          '仕入単価
    ORDR_GRID(row, Col_KIBOU_DT%) = Trim(DIS_KIBOU_DT)    '希望納期
    ORDR_GRID(row, Col_KAITO_DT%) = Trim(DIS_KAITO_DE)    '回答納期

    
    
    Set TDBGrid1.Array = ORDR_GRID
    
    TDBGrid1.style.Locked = True
    
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    'TDBGrid1.ScrollBars = dbgAutomatic
    TDBGrid1.ScrollBars = dbgNone
    DoEvents
    Sleep (500)
    
    Set BUNNO_GRID = Nothing
    
    row = 0
    
    
    Call UniCode_Conv(K1_P_SHORDER.JGYOBU, Key_JIGYOBU)
    Call UniCode_Conv(K1_P_SHORDER.NAIGAI, Key_NAIGAI)
    Call UniCode_Conv(K1_P_SHORDER.HIN_GAI, Key_HinGai)
    Call UniCode_Conv(K1_P_SHORDER.ORDER_DT, "")
    Call UniCode_Conv(K1_P_SHORDER.ORDER_NO, "")
    com = BtOpGetGreaterEqual
        
    Do
        yn = 0
        Do
            sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K1_P_SHORDER, Len(K1_P_SHORDER), 1)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    Sleep (500)
                    yn = yn + 1
                    If yn >= 500 Then
                        yn = MsgBox("他で使用中です！<資材発注Ｆ>" & Chr(13) & Chr(10) & _
                                    "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                            
                        If yn = vbNo Then Exit Function
                    End If
                            
                Case Else
                    Call File_Error(sts, com, "P_SHORDER")
                    Exit Function
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        If Trim(StrConv(P_SHORDER_REC.JGYOBU, vbUnicode)) <> Trim(Key_JIGYOBU) Then Exit Do
        If Trim(StrConv(P_SHORDER_REC.NAIGAI, vbUnicode)) <> Trim(Key_NAIGAI) Then Exit Do
        If Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)) <> Trim(Key_HinGai) Then Exit Do
               
        sts = True
               
        If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = "1" Then sts = False   'キャンセル？
            
        'If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = "9" Then sts = False      '完了？
            
        'If Trim(StrConv(P_SHORDER_REC.USE_YM, vbUnicode)) = "" Then
        '    Call UniCode_Conv(P_SHORDER_REC.USE_YM, StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode))
        'End If
        
        If StrConv(P_SHORDER_REC.USE_YM, vbUnicode) <> Key_USE_YM Then sts = False
        
        If sts = True Then
            
            row = row + 1
            
            '編集
            
            If Grid_Set_Proc() Then
                Exit Function
            End If
            
        End If
        
        com = BtOpGetNext
    Loop
    
    
    Set TDBGrid2.Array = BUNNO_GRID
    
    'TDBGrid2.style.Locked = True
    
    TDBGrid2.ReBind
    TDBGrid2.Update
    TDBGrid2.MoveFirst
    TDBGrid2.ScrollBars = dbgAutomatic
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "現在の発注情報　表示終了。", Me.hwnd, 0)
    DoEvents
    
    Call Input_UnLock                             '画面項目ロック
    
    Data_Disp = False
    
    
End Function

Private Function Grid_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   グリッド表示（移動歴データ内容）
'               Row   行数
'               mode　FALSE:ﾁｪｯｸOFF  TRUE:ﾁｪｯｸON
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim yn          As Integer
Dim com         As Integer

Dim W_QTY       As Double
Dim W_STR       As String


    Grid_Set_Proc = True

    BUNNO_GRID.ReDim Min_Row, row, Min_Col, Col2_KAN_F%

    '注文№
    BUNNO_GRID(row, Col2_ORDR_NO) = StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode)
    '仕入先
    BUNNO_GRID(row, Col2_SECT) = StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode)
    '仕入先名
    BUNNO_GRID(row, Col2_SECT_NM) = ""
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "P_UKEHARAI")
            Exit Function
    End Select
    If sts = BtNoErr Then
        BUNNO_GRID(row, Col2_SECT_NM) = Trim(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode))
    End If
    
    
    '発注数
    BUNNO_GRID(row, Col2_ORDR_QTY) = Format(CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)), "#,##0.00")
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    '仕入残
    W_QTY = CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))
    W_QTY = W_QTY - CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode))
    If W_QTY < 0 Then W_QTY = 0
    
    
    '2008.12.02 完了Ｆで判定に変更！
    If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) <> "1" Then
        If CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)) <> 0 Then
            W_QTY = 0
        End If
        
        '残の計算ではなく、発注数を残とみなす！
        W_QTY = CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))
        
    Else
        W_QTY = 0
    End If
    
    
    '   2008.12.04 仕入残:計算式！
    W_QTY = CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))
    W_QTY = W_QTY - CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode))
    If W_QTY < 0 Then W_QTY = 0
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    '       2008.12.06      上記ブロックを下記（受入日付設定ブロックで減算）に変更
        
    
    
    
    
    '受入日
    If Trim(StrConv(P_SHORDER_REC.KAN_DT, vbUnicode)) = "" Then
        W_STR = ""
    Else
        W_STR = Left(StrConv(P_SHORDER_REC.KAN_DT, vbUnicode), 4) & "/" _
                    & Mid(StrConv(P_SHORDER_REC.KAN_DT, vbUnicode), 5, 2) & "/" _
                      & Right(StrConv(P_SHORDER_REC.KAN_DT, vbUnicode), 2)
    End If
    
    '           2008.12.04  受入データ読み、最終受入日に！
    '                       ついでに、仕入残を計算！
    W_QTY = CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))
    
    '           2008.12.17　完了F判定　追加
    If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = "1" Then
        W_QTY = 0
    End If
    
    
    W_STR = ""
    com = BtOpGetGreaterEqual
    Call UniCode_Conv(K0_P_SHUKEIRE.ORDER_NO, StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode))
    Call UniCode_Conv(K0_P_SHUKEIRE.SEQNO, "")
    Do
        Do
            sts = BTRV(com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K0_P_SHUKEIRE, Len(K0_P_SHUKEIRE), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    Sleep (500)
                Case Else
                    Call File_Error(sts, com, "P_UKEHARAI")
                    Exit Function
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        If Trim(StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode)) <> Trim(StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode)) Then
            Exit Do
        End If
        If Trim(StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode)) <> "" Then
            If Trim(StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode)) > W_STR Then
                W_STR = Left(StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode), 4) & "/" _
                    & Mid(StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode), 5, 2) & "/" _
                      & Right(StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode), 2)
            End If
        End If
        If StrConv(P_SHUKEIRE_REC.KEIJYO_YM, vbUnicode) <= GW_TOUGETU Then
            W_QTY = W_QTY - CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode))
        End If
        
        If W_QTY < 0 Then
            W_QTY = 0
        End If
        
        com = BtOpGetNext
    Loop
    
    BUNNO_GRID(row, Col2_ZAN_QTY) = Format(W_QTY, "#,##0.00")

    If W_STR <> "" Then
        
    End If
    
    BUNNO_GRID(row, Col2_KAN_DT%) = Trim(W_STR)
    
    
    '希望納期
    If Trim(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode)) = "" Then
        W_STR = ""
    Else
        W_STR = Left(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 4) & "/" _
                    & Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 5, 2) & "/" _
                      & Right(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 2)
    End If
    BUNNO_GRID(row, Col2_KIBOU_DT%) = Trim(W_STR)
    
    '回答納期
    If Trim(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode)) = "" Then
        W_STR = ""
    Else
        W_STR = Left(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 4) & "/" _
                    & Mid(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 5, 2) & "/" _
                      & Right(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 2)
    End If
    BUNNO_GRID(row, Col2_KAITO_DT%) = Trim(W_STR)
    
    
    '2008/07/16追加
    '使用月
    If Trim(StrConv(P_SHORDER_REC.USE_YM, vbUnicode)) = "" Then
        W_STR = ""
    Else
        W_STR = Left(StrConv(P_SHORDER_REC.USE_YM, vbUnicode), 4) & "/" _
                    & Mid(StrConv(P_SHORDER_REC.USE_YM, vbUnicode), 5, 2)
    End If
    
    BUNNO_GRID(row, Col2_USE_DT%) = Trim(W_STR)

    '2008/12/17追加
    BUNNO_GRID(row, Col2_KAN_F%) = Trim(StrConv(P_SHORDER_REC.KAN_F, vbUnicode))


    Grid_Set_Proc = False

End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    ODR30102.MousePointer = vbHourglass

    Call Ctrl_Lock(ODR30102)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(ODR30102)


    ODR30102.MousePointer = vbDefault

End Sub

Private Sub Command1_Click(Index As Integer)
Dim yn      As Integer
Dim X_i     As Integer
Dim W_After     As String

    Select Case Index
    
            
        Case FuncEND%
            If Grid_Cor_M = True Then
                yn = MsgBox("更新されていません！！" & Chr(13) & Chr(10) & _
                            "　キャンセルしますか？", vbYesNo + vbDefaultButton2 + vbQuestion, "確認入力")
            Else
                'yn = MsgBox("キャンセルしますか？", vbYesNo + vbDefaultButton1 + vbQuestion, "確認入力")
                yn = vbYes
            End If
            
            If yn = vbNo Then
                
                Exit Sub
            End If

            Init_F = 0
    
            Set ORDR_GRID = Nothing
            Set BUNNO_GRID = Nothing
            
            Set TDBGrid1.Array = ORDR_GRID
            TDBGrid1.ReBind
            TDBGrid1.Update
            
            Set TDBGrid2.Array = BUNNO_GRID
            TDBGrid2.ReBind
            TDBGrid2.Update
            
            ODR30102_Return = True                '確認画面ｷｬﾝｾﾙ終了
            Me.Visible = False
    
    End Select

End Sub

Private Sub Form_Activate()
Dim X_i As Integer

    If Init_F <> 0 Then Exit Sub
    
    ODR30102.Top = ODR30101.Top + (ODR30101.Height - ODR30102.Height)
    ODR30102.Left = ODR30101.Left + (ODR30101.Width - ODR30102.Width) / 2
    
    Text1(ptxTANTO_CD) = ODR30101.Text1(ptxTANTO_CD)
    Lab_Dsp(plabTANTO_NM) = ODR30101.Lab_Dsp(plabTANTO_NM)
    Text1(ptxUSE_YY) = ODR30101.Text1(ptxUSE_YY)
    
    
    If Data_Disp Then
        Call Input_UnLock                             '画面項目ロック
    End If
    
    ODR30102_Return = True
    TDBGrid2.SetFocus
    
    Grid_Cor_M = False
    
    Init_F = 1
    
End Sub

Private Sub Form_Load()

Dim i       As Integer
Dim c       As String * 128
Dim sts     As Integer

Dim sBuffer As String * 255
Dim com     As String

Dim W_Date  As String





'ステータスウィンドウを作成する
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "子部品　発注情報表示", Me.hwnd, 0)
'最後の要素を-1にすると
'親ウィンドウの全体の幅の残りの幅を
'自動的に割り当てる
'PanePos(0) = 200
'PanePos(1) = 300
'PanePos(2) = -1
Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)


'画面初期処理
    'Show
    
    'Text1(ptxTANTO_CD).SetFocus
    'Max_Row = 25000
    
    Init_F = 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim yn      As Integer

    If UnloadMode <> 0 Then Exit Sub
    
    yn = MsgBox("キャンセルしますか？", vbYesNo + vbDefaultButton1 + vbQuestion, "確認入力")
    If yn = vbNo Then
        Cancel = 1
        Exit Sub
    End If
    
    Me.Visible = False
    
End Sub

Private Sub SHORI_Click(Index As Integer)
Dim yn      As Integer


    Select Case Index
            
            
        Case 1      '画面印刷
            yn = MsgBox("画面印刷しますか？", vbYesNo + vbDefaultButton2 + vbQuestion, "確認入力")
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)
        
        
        Case 0      '終了
            Call Command1_Click(FuncEND)
    
    End Select


End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
'    TDBGrid1.Bookmark = -1                 '2016.02.15
    
End Sub

Private Sub TDBGrid2_AfterColUpdate(ByVal ColIndex As Integer)
Dim W_STR       As String
    
Dim W_After     As String

    If TDBGrid2.Bookmark <= 0 Then Exit Sub
    
    Cor_Row = TDBGrid2.Bookmark
    
    W_After = Trim(TDBGrid2.Text)
    
    
    TDBGrid2.Update
    Set BUNNO_GRID = TDBGrid2.Array


End Sub

Private Sub TDBGrid2_Change()

    'Grid_Cor_M = True               '発注情報が変化した！

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index))
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index))
    End If
End Sub
Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim sts As Integer

    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Text1(Index).Locked = True Then      'ロック中項目なら処理しない
        Call Tab_Ctrl(Shift)    '移動
        Exit Sub
    End If
                        '入力文字数チェック
    If ERR_CHK(Index) Then
        Call Text1_GotFocus(Index)
        Text1(Index).SetFocus
        Exit Sub
    End If
    
    If Index = ptxUSE_YY Then
        If Data_Disp Then
            MsgBox "指定条件の注文情報で、表示失敗！", vbExclamation
            Call Input_UnLock                             '画面項目ロック
            Call Text1_GotFocus(ptxTOP)
            Text1(ptxTOP).SetFocus
            Exit Sub
        End If
        
        
        TDBGrid1.SetFocus
        
        Exit Sub
    End If
    
    Call Tab_Ctrl(Shift)    '移動
    
End Sub

