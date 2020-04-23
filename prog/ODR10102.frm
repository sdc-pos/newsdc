VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form ODR10102 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "親部品　分納情報登録"
   ClientHeight    =   9630
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   13350
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
   ScaleHeight     =   9630
   ScaleWidth      =   13350
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
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
      Index           =   2
      Left            =   4020
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2700
      Width           =   1095
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
      TabIndex        =   1
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   780
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   6930
      ScaleHeight     =   195
      ScaleWidth      =   270
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   330
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
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "更　新"
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
      Left            =   165
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   1800
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   1095
      Left            =   540
      TabIndex        =   8
      Top             =   1320
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   1931
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "親部品注文№"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "分納"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "親部品"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "数 量"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "部材センター　注文納期　"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "組立可能日"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   " 　親部品 　回答納期"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "使用月"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "完了日付"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "ＫＥＹ項目"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).AllowFocus=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   688
      Splits(0).ScrollBars=   0
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2990"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2858"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1058"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=926"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8193"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2646"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2514"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8196"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=1402"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1270"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8194"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2831"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2699"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8193"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2831"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2699"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=8193"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=2831"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2699"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=8193"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=2117"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=1984"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=8193"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=2831"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2699"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=8193"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(9).Width=2778"
      Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=8196"
      Splits(0)._ColumnProps(50)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(51)=   "Column(9).Order=10"
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
      Enabled         =   0   'False
      HeadLines       =   2
      FootLines       =   1
      Caption         =   "親部品　注文情報"
      AllowArrows     =   0   'False
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFF00&,.locked=-1,.bold=0"
      _StyleDefs(7)   =   ":id=1,.fontsize=1125,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H80FF00&,.bold=0"
      _StyleDefs(11)  =   ":id=2,.fontsize=1125,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ＭＳ ゴシック"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.valignment=2,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(19)  =   ":id=7,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(20)  =   ":id=7,.fontname=ＭＳ ゴシック"
      _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFF80&"
      _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFF80&"
      _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(26)  =   "Splits(0).Style:id=87,.parent=1,.bgcolor=&H80FF80&"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=96,.parent=4,.bgcolor=&H80FF00&"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=89,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=91,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=94,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=95,.parent=10,.bgcolor=&HFFFF00&"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=102,.parent=87,.locked=-1"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=99,.parent=88"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=100,.parent=89"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=101,.parent=91"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=106,.parent=87,.alignment=2,.locked=-1"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=103,.parent=88"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=104,.parent=89"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=105,.parent=91"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=110,.parent=87,.locked=-1"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=107,.parent=88"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=108,.parent=89"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=109,.parent=91"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=114,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=111,.parent=88"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=112,.parent=89"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=113,.parent=91"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=118,.parent=87,.alignment=2,.locked=-1"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=115,.parent=88"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=116,.parent=89"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=117,.parent=91"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=122,.parent=87,.alignment=2,.locked=-1"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=119,.parent=88"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=120,.parent=89"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=121,.parent=91"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=126,.parent=87,.alignment=2,.locked=-1"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=123,.parent=88"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=124,.parent=89"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=125,.parent=91"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=130,.parent=87,.alignment=2,.locked=-1"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=127,.parent=88"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=128,.parent=89"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=129,.parent=91"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=134,.parent=87,.alignment=2,.locked=-1"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=131,.parent=88"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=132,.parent=89"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=133,.parent=91"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=138,.parent=87"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=135,.parent=88"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=136,.parent=89"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=137,.parent=91"
      _StyleDefs(78)  =   "Named:id=33:Normal"
      _StyleDefs(79)  =   ":id=33,.parent=0"
      _StyleDefs(80)  =   "Named:id=34:Heading"
      _StyleDefs(81)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(82)  =   ":id=34,.wraptext=-1"
      _StyleDefs(83)  =   "Named:id=35:Footing"
      _StyleDefs(84)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(85)  =   "Named:id=36:Selected"
      _StyleDefs(86)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(87)  =   "Named:id=37:Caption"
      _StyleDefs(88)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(89)  =   "Named:id=38:HighlightRow"
      _StyleDefs(90)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(91)  =   "Named:id=39:EvenRow"
      _StyleDefs(92)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(93)  =   "Named:id=40:OddRow"
      _StyleDefs(94)  =   ":id=40,.parent=33"
      _StyleDefs(95)  =   "Named:id=41:RecordSelector"
      _StyleDefs(96)  =   ":id=41,.parent=34"
      _StyleDefs(97)  =   "Named:id=42:FilterBar"
      _StyleDefs(98)  =   ":id=42,.parent=33"
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid2 
      Height          =   5955
      Left            =   2760
      TabIndex        =   9
      Top             =   3120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   10504
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   4
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "削除"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "分納"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "数 量"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "部材センター　注文納期　"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "組立可能日"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   " 　親部品 　回答納期"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "使用月"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "ＫＥＹ項目"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=979"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=17"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1058"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=926"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8193"
      Splits(0)._ColumnProps(10)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=1402"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1270"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=2831"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2699"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=8193"
      Splits(0)._ColumnProps(21)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=2831"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2699"
      Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=8193"
      Splits(0)._ColumnProps(27)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(29)=   "Column(5).Width=2831"
      Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=2699"
      Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=1"
      Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(34)=   "Column(6).Width=2117"
      Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=1984"
      Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=8193"
      Splits(0)._ColumnProps(38)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(40)=   "Column(7).Width=2778"
      Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(43)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(44)=   "Column(7).Order=8"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageFooterFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   1
      Caption         =   "親部品　分納情報"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=16,.parent=87,.alignment=2,.valignment=2"
      _StyleDefs(37)  =   ":id=16,.bgcolor=&H80000005&"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=88"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=89"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=91"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=106,.parent=87,.alignment=2,.locked=-1"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=103,.parent=88"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=104,.parent=89"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=105,.parent=91"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=114,.parent=87,.alignment=1,.bgcolor=&H80000005&"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=111,.parent=88"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=112,.parent=89"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=113,.parent=91"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=118,.parent=87,.alignment=2,.locked=-1"
      _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=115,.parent=88"
      _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=116,.parent=89"
      _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=117,.parent=91"
      _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=122,.parent=87,.alignment=2,.locked=-1"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=119,.parent=88"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=120,.parent=89"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=121,.parent=91"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=126,.parent=87,.alignment=2,.bgcolor=&H80000005&"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=123,.parent=88"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=124,.parent=89"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=125,.parent=91"
      _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=130,.parent=87,.alignment=2,.locked=-1"
      _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=127,.parent=88"
      _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=128,.parent=89"
      _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=129,.parent=91"
      _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=138,.parent=87"
      _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=135,.parent=88"
      _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=136,.parent=89"
      _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=137,.parent=91"
      _StyleDefs(69)  =   "Named:id=33:Normal"
      _StyleDefs(70)  =   ":id=33,.parent=0"
      _StyleDefs(71)  =   "Named:id=34:Heading"
      _StyleDefs(72)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(73)  =   ":id=34,.wraptext=-1"
      _StyleDefs(74)  =   "Named:id=35:Footing"
      _StyleDefs(75)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(76)  =   "Named:id=36:Selected"
      _StyleDefs(77)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(78)  =   "Named:id=37:Caption"
      _StyleDefs(79)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(80)  =   "Named:id=38:HighlightRow"
      _StyleDefs(81)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(82)  =   "Named:id=39:EvenRow"
      _StyleDefs(83)  =   ":id=39,.parent=33,.bgcolor=&H80FF00&"
      _StyleDefs(84)  =   "Named:id=40:OddRow"
      _StyleDefs(85)  =   ":id=40,.parent=33"
      _StyleDefs(86)  =   "Named:id=41:RecordSelector"
      _StyleDefs(87)  =   ":id=41,.parent=34"
      _StyleDefs(88)  =   "Named:id=42:FilterBar"
      _StyleDefs(89)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "合計数"
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
      Index           =   2
      Left            =   3240
      TabIndex        =   11
      Top             =   2760
      Width           =   720
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   840
      Width           =   720
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "更新"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "画面印刷"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "ｷｬﾝｾﾙ"
         Index           =   2
      End
   End
End
Attribute VB_Name = "ODR10102"
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
Private Const FuncCOR% = 0       '更新
Private Const FuncEND% = 1       '終了

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


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 9              '最大列数

Private Const Col_ORDR_NO% = 0              '親部品　注文№
Private Const Col_BUNNO% = 1                '分納回数
Private Const Col_OYA_ITEM% = 2             '親部品コード
Private Const Col_ORDR_QTY% = 3             '注文数量
Private Const Col_NOUKI% = 4                '親部品　注文納期
Private Const Col_OK_DT% = 5              '組立可能日
Private Const Col_KAITO% = 6                '親部品　回答納期
Private Const Col_USE_YM% = 7               '使用月
Private Const Col_FIN_DT% = 8               '完了日付
Private Const Col_KEY% = 9                  'データＫｅｙ情報

Private Const Col2_DEL% = 0                 '削除指示
Private Const Col2_BUNNO% = 1               '分納回数
Private Const Col2_ORDR_QTY% = 2            '注文数量
Private Const Col2_NOUKI% = 3               '親部品　注文納期
Private Const Col2_OK_DT% = 4             '組立可能日
Private Const Col2_KAITO% = 5               '親部品　回答納期
Private Const Col2_USE_YM% = 6              '使用月




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

    Grid_Err_Chk = True
    
    Select Case Index
        Case Col2_DEL%                   '削除マーク
            
            If BUNNO_GRID(Cor_Row, Index) Then
                W_STR = Trim(BUNNO_GRID(Cor_Row, Col2_NOUKI%))
                If W_STR = "" Then
                    BUNNO_GRID(Cor_Row, Index) = False
                    MsgBox "未登録行　→　削除不可！", vbExclamation
                    
                    TDBGrid2.ReBind
                    TDBGrid2.Update
                    TDBGrid2.MoveFirst
                    TDBGrid2.ScrollBars = dbgAutomatic
                    Exit Function
                End If
            End If
                        
        Case Col2_ORDR_QTY%              '注文数量
            
            If Trim(W_Aft) = "" Then
                BUNNO_GRID(Cor_Row, Col2_DEL%) = False
                BUNNO_GRID(Cor_Row, Col2_NOUKI%) = ""
                BUNNO_GRID(Cor_Row, Col2_OK_DT%) = ""
                BUNNO_GRID(Cor_Row, Col2_KAITO%) = ""
                BUNNO_GRID(Cor_Row, Col2_USE_YM%) = ""
                TDBGrid2.ReBind
                TDBGrid2.Update
                TDBGrid2.MoveFirst
                TDBGrid2.ScrollBars = dbgAutomatic
                Grid_Err_Chk = False
                Exit Function
            End If
            
        
            If Not IsNumeric(W_Aft) Then
                MsgBox "注文数量　数値エラー！", vbExclamation
                Exit Function
            End If
            BUNNO_GRID(Cor_Row, Col2_NOUKI%) = DIS_NOUKI 'ORDR_GRID(1, Col_NOUKI%)
            BUNNO_GRID(Cor_Row, Col2_OK_DT%) = DIS_OK_DT 'ORDR_GRID(1, Col_OK_DT%)
            BUNNO_GRID(Cor_Row, Col2_USE_YM%) = DIS_USE_YM 'ORDR_GRID(1, Col_USE_YM%)
            
            Call SUM_QTY
            
            TDBGrid2.ReBind
            TDBGrid2.Update
            TDBGrid2.MoveFirst
            TDBGrid2.ScrollBars = dbgAutomatic
            
            
            
        Case Col2_KAITO%                 '親部品　注文納期
            If IsDate(W_Aft) Then
                W_STR = Format(W_Aft, "yyyy/mm/dd")
                BUNNO_GRID(Cor_Row, Index) = W_STR
                
                TDBGrid2.ReBind
                TDBGrid2.Update
                TDBGrid2.MoveFirst
                TDBGrid2.ScrollBars = dbgAutomatic
            Else
                If Trim(W_Aft) <> "" Or Trim(BUNNO_GRID(Cor_Row, Col2_ORDR_QTY)) <> "" Then
                    MsgBox "親部品　回答納期　日付エラー！", vbExclamation
                    Exit Function
                End If
            End If

        
    End Select
    
    DoEvents
    
    If Trim(W_Aft) <> "" Then
        Select Case Index
                 '注文数量       '回答納期
            Case Col2_ORDR_QTY%, Col2_KAITO%
                W_STR = Trim(BUNNO_GRID(Cor_Row, Col2_ORDR_QTY))
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

Dim X_i         As Long

Dim W_Key       As String

Dim W_STR       As String
Dim W_Date      As String
Dim cnt         As Integer

    Data_Disp = True
    
    row = Min_Row - 1
    Call Input_Lock                             '画面項目ロック
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "親部品　分納情報　検索・表示中。　＜Data_Disp＞", Me.hwnd, 0)
    DoEvents
    
    
    Text1(ptxSUM_QTY) = ""
    W_SUM = 0
    
    Set ORDR_GRID = Nothing
    Set BUNNO_GRID = Nothing
    
    Call UniCode_Conv(K0_ODR_ORDER.SHIMUKE, GW_SIMUKE)
    Call UniCode_Conv(K0_ODR_ORDER.JGYOBU, GW_JIGYOBU)
    Call UniCode_Conv(K0_ODR_ORDER.NAIGAI, GW_NAIGAI)
    Call UniCode_Conv(K0_ODR_ORDER.HIN_GAI, GW_HINGAI)
    Call UniCode_Conv(K0_ODR_ORDER.ORDER_NO, DIS_ORDR_NO)
    Call UniCode_Conv(K0_ODR_ORDER.INS_NO, DIS_KEY)
    Call UniCode_Conv(K0_ODR_ORDER.BUN_NO, "")
    
    com = BtOpGetEqual
        
        Do
            sts = BTRV(BtOpGetEqual, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_ORDER, Len(K0_ODR_ORDER), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    yn = MsgBox("他で使用中です！<注文Ｆ>" & Chr(13) & Chr(10) & _
                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If yn = vbNo Then Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "ODR_ORDER")
                    Exit Function
            End Select
        Loop
    
    
    If sts <> BtNoErr Then
        MsgBox "基の発注情報　読込エラー！", vbExclamation
        Call Input_UnLock                             '画面項目ロック
        Exit Function
    End If
    
    DIS_BUNNO = CStr(CInt(StrConv(ODR_ORDER_REC.BUN_KB, vbUnicode)))
'2008.03.21    DIS_ORDR_QTY = CStr(CDbl(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode)))
    W_STR = Trim(StrConv(ODR_ORDER_REC.CYUMON_DT, vbUnicode))
    DIS_NOUKI = Left(W_STR, 4) & "/" & Mid(W_STR, 5, 2) & "/" & Right(W_STR, 2)
    W_STR = Trim(StrConv(ODR_ORDER_REC.KAITO_DT, vbUnicode))
    If W_STR = "" Then
        DIS_KAITO = ""
    Else
        DIS_KAITO = Left(W_STR, 4) & "/" & Mid(W_STR, 5, 2) & "/" & Right(W_STR, 2)
    End If
    
    '基情報のグリッド表示
    row = 1
    ORDR_GRID.ReDim Min_Row, row, Min_Col, Max_Col
    
    ORDR_GRID(row, Col_ORDR_NO%) = DIS_ORDR_NO
    ORDR_GRID(row, Col_BUNNO%) = DIS_BUNNO          '分納回数
    ORDR_GRID(row, Col_OYA_ITEM%) = DIS_OYA_ITEM    '親部品コード
    ORDR_GRID(row, Col_ORDR_QTY%) = DIS_ORDR_QTY    '注文数量
    ORDR_GRID(row, Col_NOUKI%) = DIS_NOUKI          '親部品　注文納期
    ORDR_GRID(row, Col_OK_DT%) = DIS_OK_DT      '組立可能日
    ORDR_GRID(row, Col_KAITO%) = DIS_KAITO          '親部品　回答納期
    ORDR_GRID(row, Col_USE_YM%) = DIS_USE_YM        '使用月
    ORDR_GRID(row, Col_FIN_DT%) = DIS_FIN_DT        '完了日付
    ORDR_GRID(row, Col_KEY%) = DIS_KEY              'データＫｅｙ情報
    
    
    Set TDBGrid1.Array = ORDR_GRID
    
    TDBGrid1.style.Locked = True
    
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    'TDBGrid1.ScrollBars = dbgAutomatic
    TDBGrid1.ScrollBars = dbgNone
    
    
    Set BUNNO_GRID = Nothing
    
    row = 0
    cnt = 0
    
    Call UniCode_Conv(K0_ODR_ORDER.SHIMUKE, GW_SIMUKE)
    Call UniCode_Conv(K0_ODR_ORDER.JGYOBU, GW_JIGYOBU)
    Call UniCode_Conv(K0_ODR_ORDER.NAIGAI, GW_NAIGAI)
    Call UniCode_Conv(K0_ODR_ORDER.HIN_GAI, GW_HINGAI)
    Call UniCode_Conv(K0_ODR_ORDER.ORDER_NO, DIS_ORDR_NO)
    Call UniCode_Conv(K0_ODR_ORDER.INS_NO, DIS_KEY)
    W_STR = String(UBound(K0_ODR_ORDER.BUN_NO), "0") & "1"
    Call UniCode_Conv(K0_ODR_ORDER.BUN_NO, W_STR)
    
    com = BtOpGetGreaterEqual
    Do
        Do
            sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_ORDER, Len(K0_ODR_ORDER), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    yn = MsgBox("他で使用中です！<注文Ｆ>" & Chr(13) & Chr(10) & _
                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If yn = vbNo Then
                        Call Input_UnLock                             '画面項目ロック
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "ODR_ORDER")
                    Call Input_UnLock                             '画面項目ロック
                    Exit Function
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode)) <> Trim(GW_SIMUKE) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode)) <> Trim(GW_HINGAI) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.INS_NO, vbUnicode)) <> Trim(DIS_KEY) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.ORDER_NO, vbUnicode)) <> Trim(DIS_ORDR_NO) Then Exit Do
        
        row = row + 1
        cnt = cnt + 1
        DIS_BUNNO = Format(cnt, "00")
        DIS2_QTY = CStr(CDbl(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode)))
        W_STR = Trim(StrConv(ODR_ORDER_REC.KAITO_DT, vbUnicode))
        If W_STR = "" Then
            W_STR = Trim(StrConv(ODR_ORDER_REC.CYUMON_DT, vbUnicode))
        End If
        DIS2_KAITO = Left(W_STR, 4) & "/" & Mid(W_STR, 5, 2) & "/" & Right(W_STR, 2)
        
        If Grid_Set_Proc() Then
            Exit Function
        End If
            
        com = BtOpGetNext
    Loop
    
    DIS2_QTY = ""
    DIS2_KAITO = ""
    If cnt < 30 Then
        For cnt = row To 30
            W_Key = Format(cnt, "00")
                    
            DIS_BUNNO = W_Key
                    
            row = row + 1
            
            If Grid_Set_Proc() Then
                Exit Function
            End If
        Next cnt
    End If
    
    Set TDBGrid2.Array = BUNNO_GRID
    
    'TDBGrid2.style.Locked = True
    
    TDBGrid2.ReBind
    TDBGrid2.Update
    TDBGrid2.MoveFirst
    TDBGrid2.ScrollBars = dbgAutomatic
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "現在の分納情報　表示終了　→　登録・修正して下さい。", Me.hwnd, 0)
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
Dim W_Row       As Long


    Grid_Set_Proc = True

    BUNNO_GRID.ReDim Min_Row, row, Min_Col, Col2_USE_YM%

    BUNNO_GRID(row, Col2_BUNNO) = DIS_BUNNO
    BUNNO_GRID(row, Col2_ORDR_QTY) = DIS2_QTY
    BUNNO_GRID(row, Col2_NOUKI%) = DIS_NOUKI
    BUNNO_GRID(row, Col2_OK_DT%) = DIS_OK_DT
    BUNNO_GRID(row, Col2_KAITO%) = DIS2_KAITO
    BUNNO_GRID(row, Col2_USE_YM%) = DIS_USE_YM

    Grid_Set_Proc = False

End Function

Private Sub SUM_QTY()
Dim W_Row   As Integer
Dim W_STR   As String
    
    
    W_SUM = 0
    For W_Row = Min_Row To BUNNO_GRID.UpperBound(1)
        If Trim(BUNNO_GRID(W_Row, Col2_ORDR_QTY)) <> "" Then
            If IsNumeric(Trim(BUNNO_GRID(W_Row, Col2_ORDR_QTY))) Then
                W_SUM = W_SUM + CDbl(Trim(BUNNO_GRID(W_Row, Col2_ORDR_QTY)))
                
            End If
        End If
    Next W_Row
    
    
    W_STR = Format(W_SUM, "#,##0.00")
    
    If Right(W_STR, 1) = "0" Then
        W_STR = Trim(Left(W_STR, Len(W_STR) - 1))
    End If
    If Right(W_STR, 1) = "0" Then
        W_STR = Trim(Left(W_STR, Len(W_STR) - 2))
    End If
    
    Text1(ptxSUM_QTY) = W_STR
    
    
End Sub

Private Function Rec_UPDT(In_Lock As Integer)
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim X_i         As Integer

Dim W_Key       As String
Dim W_No        As String
Dim W_STR       As String
Dim W_Date      As String
    
    If In_Lock = True Then
        Rec_UPDT = True
    End If
    If In_Lock = True Then
        Call Input_Lock
    End If
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "親部品　分納情報　更新中！［旧情報削除］　＜Rec_UPDT＞", Me.hwnd, 0)
    DoEvents
    
    '最初に、現在の分納情報を削除する。
    Call UniCode_Conv(K0_ODR_ORDER.SHIMUKE, GW_SIMUKE)
    Call UniCode_Conv(K0_ODR_ORDER.JGYOBU, GW_JIGYOBU)
    Call UniCode_Conv(K0_ODR_ORDER.NAIGAI, GW_NAIGAI)
    Call UniCode_Conv(K0_ODR_ORDER.HIN_GAI, GW_HINGAI)
    Call UniCode_Conv(K0_ODR_ORDER.ORDER_NO, DIS_ORDR_NO)
    Call UniCode_Conv(K0_ODR_ORDER.INS_NO, DIS_KEY)
    W_STR = String(UBound(K0_ODR_ORDER.BUN_NO), "0") & "1"
    Call UniCode_Conv(K0_ODR_ORDER.BUN_NO, W_STR)
    
    com = BtOpGetGreaterEqual
    Do
        Do
            sts = BTRV(com + BtSNoWait, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_ORDER, Len(K0_ODR_ORDER), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    yn = MsgBox("他で使用中です！<注文Ｆ>" & Chr(13) & Chr(10) & _
                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If yn = vbNo Then GoTo Err_Exit
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "ODR_ORDER")
                    GoTo Err_Exit
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode)) <> Trim(GW_SIMUKE) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode)) <> Trim(GW_HINGAI) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.ORDER_NO, vbUnicode)) <> Trim(DIS_ORDR_NO) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.INS_NO, vbUnicode)) <> Trim(DIS_KEY) Then Exit Do
        
        If Trim(StrConv(ODR_ORDER_REC.BUN_NO, vbUnicode)) <> "" Then
            If CInt(Trim(StrConv(ODR_ORDER_REC.BUN_NO, vbUnicode))) <> 0 Then
                Do
                    sts = BTRV(BtOpDelete, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_ORDER, Len(K0_ODR_ORDER), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                            Sleep (500)
                        Case Else
                            Call File_Error(sts, BtOpDelete, "ODR_ORDER")
                            GoTo Err_Exit
                    End Select
                Loop
            End If
        End If
        
        com = BtOpGetNext
    Loop
    
    '基（親）情報の読込
    Call UniCode_Conv(K0_ODR_ORDER.SHIMUKE, GW_SIMUKE)
    Call UniCode_Conv(K0_ODR_ORDER.JGYOBU, GW_JIGYOBU)
    Call UniCode_Conv(K0_ODR_ORDER.NAIGAI, GW_NAIGAI)
    Call UniCode_Conv(K0_ODR_ORDER.HIN_GAI, GW_HINGAI)
    Call UniCode_Conv(K0_ODR_ORDER.ORDER_NO, DIS_ORDR_NO)
    Call UniCode_Conv(K0_ODR_ORDER.INS_NO, DIS_KEY)
    Call UniCode_Conv(K0_ODR_ORDER.BUN_NO, "")
    
    com = BtOpGetEqual
    Do
        Do
            sts = BTRV(BtOpGetEqual, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_ORDER, Len(K0_ODR_ORDER), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    yn = MsgBox("他で使用中です！<注文Ｆ>" & Chr(13) & Chr(10) & _
                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If yn = vbNo Then GoTo Err_Exit
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "ODR_ORDER")
                    GoTo Err_Exit
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        sts = BtErrEOF
        If Trim(StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode)) <> Trim(GW_SIMUKE) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode)) <> Trim(GW_HINGAI) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.ORDER_NO, vbUnicode)) <> Trim(DIS_ORDR_NO) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.INS_NO, vbUnicode)) <> Trim(DIS_KEY) Then Exit Do
        
        sts = BtNoErr
        Exit Do
    Loop
    If sts <> BtNoErr Then
        MsgBox "基の発注情報　読込エラー！", vbExclamation
        GoTo Err_Exit
    End If
    
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "親部品　分納情報　更新中！［新分納情報追加更新］　＜Rec_UPDT＞", Me.hwnd, 0)
    DoEvents
    
    '表の分納内容で登録
    
    X_i = 0
    
    For Cor_Row = Min_Row To BUNNO_GRID.UpperBound(1)
        If Trim(BUNNO_GRID(Cor_Row, Col2_ORDR_QTY)) <> "" Then
            X_i = X_i + 1
            Call UniCode_Conv(ODR_ORDER_REC.BUN_NO, Format(X_i, "000"))
            
            Call UniCode_Conv(ODR_ORDER_REC.BUN_KB, String(UBound(ODR_ORDER_REC.BUN_KB) + 1, "0"))
            
            Call UniCode_Conv(ODR_ORDER_REC.ODR_QTY, Format(CLng(BUNNO_GRID(Cor_Row, Col2_ORDR_QTY)), "00000"))
            
            W_STR = BUNNO_GRID(Cor_Row, Col2_KAITO)
            W_Date = Left(W_STR, 4) & Mid(W_STR, 6, 2) & Right(W_STR, 2)
            Call UniCode_Conv(ODR_ORDER_REC.KAITO_DT, W_Date)
            
            Call UniCode_Conv(ODR_ORDER_REC.UPD_TANTO, Text1(ptxTANTO_CD))
            W_STR = Format(Date, "yyyymmdd")
            Call UniCode_Conv(ODR_ORDER_REC.INS_DT, W_STR)
            'Call UniCode_Conv(ODR_ORDER_REC.UPD_DT, W_STR)
            W_STR = Format(Time, "hhmmss")
            Call UniCode_Conv(ODR_ORDER_REC.INS_TM, W_STR)
            'Call UniCode_Conv(ODR_ORDER_REC.UPD_TM, W_STR)
            Call UniCode_Conv(ODR_ORDER_REC.UPD_PG, Trim(App.EXEName))
            
            Do
                sts = BTRV(BtOpInsert, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_ORDER, Len(K0_ODR_ORDER), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                        Sleep (500)
                    Case Else
                        Call File_Error(sts, BtOpInsert, "ODR_ORDER")
                        GoTo Err_Exit
                End Select
            Loop
        End If
    
    Next Cor_Row
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "親部品　分納情報　更新中！［基（親）情報更新］　＜Rec_UPDT＞", Me.hwnd, 0)
    DoEvents
    
    '基（親）情報の更新
    '基（親）情報の読込
    Call UniCode_Conv(K0_ODR_ORDER.SHIMUKE, GW_SIMUKE)
    Call UniCode_Conv(K0_ODR_ORDER.JGYOBU, GW_JIGYOBU)
    Call UniCode_Conv(K0_ODR_ORDER.NAIGAI, GW_NAIGAI)
    Call UniCode_Conv(K0_ODR_ORDER.HIN_GAI, GW_HINGAI)
    Call UniCode_Conv(K0_ODR_ORDER.ORDER_NO, DIS_ORDR_NO)
    Call UniCode_Conv(K0_ODR_ORDER.INS_NO, DIS_KEY)
    Call UniCode_Conv(K0_ODR_ORDER.BUN_NO, "")
    
    Do
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_ORDER, Len(K0_ODR_ORDER), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    yn = MsgBox("他で使用中です！<注文Ｆ>" & Chr(13) & Chr(10) & _
                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If yn = vbNo Then GoTo Err_Exit
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "ODR_ORDER")
                    GoTo Err_Exit
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
                
        Call UniCode_Conv(ODR_ORDER_REC.BUN_KB, Format(X_i, "000"))
        
        
        Call UniCode_Conv(ODR_ORDER_REC.UPD_TANTO, Text1(ptxTANTO_CD))
        W_STR = Format(Date, "yyyymmdd")
        Call UniCode_Conv(ODR_ORDER_REC.UPD_DT, W_STR)
        W_STR = Format(Time, "hhmmss")
        Call UniCode_Conv(ODR_ORDER_REC.UPD_TM, W_STR)
        Call UniCode_Conv(ODR_ORDER_REC.UPD_PG, Trim(App.EXEName))
        
        Do
            sts = BTRV(BtOpUpdate, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_ORDER, Len(K0_ODR_ORDER), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    Sleep (500)
                Case Else
                    Call File_Error(sts, BtOpUpdate, "ODR_ORDER")
                    GoTo Err_Exit
            End Select
        Loop
        
        sts = BtNoErr
        Exit Do
    Loop
    If sts <> BtNoErr Then
        MsgBox "基の発注情報　読込エラー！", vbExclamation
        GoTo Err_Exit
    End If
        
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "分納情報の更新 正常終了｡ < Rec_UPDT > ", Me.hwnd, 0)
    DoEvents
    
    
    Rec_UPDT = False
Err_Exit:

    If In_Lock = True Then
        Call Input_UnLock
    End If

End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    ODR10102.MousePointer = vbHourglass

    Call Ctrl_Lock(ODR10102)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(ODR10102)


    ODR10102.MousePointer = vbDefault

End Sub

Private Sub Command1_Click(Index As Integer)
Dim yn      As Integer
Dim X_i     As Integer
Dim W_After     As String

    Select Case Index
    
        Case FuncCOR%
            If Grid_Cor_M <> True Then
                Exit Sub
            End If
            
            Set BUNNO_GRID = TDBGrid2.Array
            TDBGrid1.Update
            
    
            For Cor_Row = Min_Row To BUNNO_GRID.UpperBound(1)
            
                For X_i = Col2_DEL To Col2_KAITO%
                    
                    W_After = BUNNO_GRID(Cor_Row, X_i)
                    
                    If Grid_Err_Chk(X_i, W_After) Then
                        TDBGrid2.SetFocus
                        Exit Sub
                    End If
                
                Next X_i
                
            Next Cor_Row
            
            If CDbl(DIS_ORDR_QTY) <> CDbl(Text1(ptxSUM_QTY)) Then
                MsgBox "注文数の合計が不一致エラー！", vbExclamation
                TDBGrid2.SetFocus
                Exit Sub
            End If
            
            
            yn = MsgBox("更新しますか？", vbYesNo + vbDefaultButton1 + vbQuestion, "確認入力")
            'yn = vbYes
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            '更新処理
            If Rec_UPDT(True) Then
                MsgBox "更新失敗！", vbExclamation
                Call Input_UnLock
            End If
            
            Grid_Cor_M = False
            Init_F = 0
            ODR10102_Return = False                '確認画面 更新＆終了
            Me.Visible = False
            Exit Sub
            
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
            ODR10102_Return = True                '確認画面ｷｬﾝｾﾙ終了
            Me.Visible = False
    
    End Select

End Sub

Private Sub Form_Activate()
Dim X_i As Integer

    If Init_F <> 0 Then Exit Sub
    
    ODR10102.Top = ODR10101.Top + (ODR10101.Height - ODR10102.Height)
    ODR10102.Left = ODR10101.Left + (ODR10101.Width - ODR10102.Width) / 2
    
    Text1(ptxTANTO_CD) = ODR10101.Text1(ptxTANTO_CD)
    Lab_Dsp(plabTANTO_NM%) = ODR10101.Lab_Dsp(plabTANTO_NM%)
    Text1(ptxUSE_YY) = ODR10101.Text1(ptxUSE_YY)
    
    
    If Data_Disp Then
        Call Input_UnLock                             '画面項目ロック
    End If
    
    ODR10102_Return = True
    TDBGrid2.SetFocus
    
    Grid_Cor_M = False
    
    Init_F = 1
    
End Sub

Private Sub Form_Load()
Dim cc As tagINITCOMMONCONTROLSEX
'Dim PanePos(2) As Long

Dim i       As Integer
Dim c       As String * 128
Dim sts     As Integer

Dim sBuffer As String * 255
Dim com     As String

Dim W_Date  As String




'コモンコントロールを初期化する
cc.dwSize = Len(cc)
cc.dwICC = ICC_BAR_CLASSES

'ステータスウィンドウを作成する
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "親部品　分納情報登録", Me.hwnd, 0)
'ペイン複数作る
'最後の要素を-1にすると
'親ウィンドウの全体の幅の残りの幅を
'自動的に割り当てる
'PanePos(0) = 200
'PanePos(1) = 300
'PanePos(2) = -1
'Call SendMessageAny(hStatusWnd, SB_SETPARTS, 3, PanePos(0))
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
    
        Case 0      '更新
            Call Command1_Click(FuncCOR)
            
            
        Case 1      '画面印刷
            yn = MsgBox("画面印刷しますか？", vbYesNo + vbDefaultButton2 + vbQuestion, "確認入力")
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)
        
        
        Case 2      '終了
            Call Command1_Click(FuncEND)
    
    End Select


End Sub
Private Sub TDBGrid1_DblClick()

    If TDBGrid1.Bookmark = -1 Then
    Else
        
        'ODR10102.Show vbModal
        
        'If KENPIN_Update_Proc() Then
        '    Unload Me
        'End If
    End If
    
    '再表示
'    If List_Disp Then
'        Unload Me
'    End If


End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    TDBGrid1.Bookmark = -1
    
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

    Grid_Cor_M = True               '分納情報が変化した！

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
    
    If Index = ptxUSE_YY% Then
        If Data_Disp Then
            MsgBox "指定条件の注文情報で、表示失敗！", vbExclamation
            Call Input_UnLock                             '画面項目ロック
            Call Text1_GotFocus(ptxTOP%)
            Text1(ptxTOP%).SetFocus
            Exit Sub
        End If
        
        
        TDBGrid1.SetFocus
        
        Exit Sub
    End If
    
    Call Tab_Ctrl(Shift)    '移動
    
End Sub

