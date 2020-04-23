VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form ODR10104 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "親部品引当情報 (2010/05/08)"
   ClientHeight    =   8220
   ClientLeft      =   150
   ClientTop       =   240
   ClientWidth     =   11640
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
   ScaleHeight     =   8220
   ScaleWidth      =   11640
   StartUpPosition =   2  '画面の中央
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
      Height          =   390
      Index           =   2
      Left            =   8700
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "2010/05/08[10:10]"
      Top             =   675
      Width           =   2265
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
      Height          =   390
      Index           =   1
      Left            =   8700
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   300
      Width           =   2265
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
      Height          =   390
      Index           =   0
      Left            =   3300
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   600
      Width           =   2490
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   4020
      ScaleHeight     =   195
      ScaleWidth      =   270
      TabIndex        =   2
      Top             =   180
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
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   240
      Width           =   1800
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   4290
      Left            =   150
      TabIndex        =   1
      Top             =   3825
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   7567
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "回答納期"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "注文納期"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "使用月"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "親品番"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "注文№"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "必要総数"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "引当数"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "不足数"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1931"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1799"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1931"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1799"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1773"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1640"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=3519"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3387"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8708"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2302"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2170"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2117"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1984"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=8706"
      Splits(0)._ColumnProps(30)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(32)=   "Column(6).Width=2117"
      Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=1984"
      Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=514"
      Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(37)=   "Column(7).Width=2117"
      Splits(0)._ColumnProps(38)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(7)._WidthInPix=1984"
      Splits(0)._ColumnProps(40)=   "Column(7)._ColStyle=514"
      Splits(0)._ColumnProps(41)=   "Column(7).Order=8"
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
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "親部品　引当／不足情報"
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
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=88,.parent=2,.alignment=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=89,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=91,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=94,.parent=9,.bgcolor=&H80FF80&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=95,.parent=10,.bgcolor=&H80FF00&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=87,.alignment=2"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=87,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=87,.alignment=2"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=106,.parent=87,.alignment=3,.locked=-1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=103,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=104,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=105,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=87"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=16,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=13,.parent=88"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=14,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=15,.parent=91"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=20,.parent=87,.alignment=1"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=17,.parent=88"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=18,.parent=89"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=19,.parent=91"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=24,.parent=87,.alignment=1"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=21,.parent=88"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=22,.parent=89"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=23,.parent=91"
      _StyleDefs(68)  =   "Named:id=33:Normal"
      _StyleDefs(69)  =   ":id=33,.parent=0"
      _StyleDefs(70)  =   "Named:id=34:Heading"
      _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(72)  =   ":id=34,.wraptext=-1"
      _StyleDefs(73)  =   "Named:id=35:Footing"
      _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   "Named:id=36:Selected"
      _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=37:Caption"
      _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(79)  =   "Named:id=38:HighlightRow"
      _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(81)  =   "Named:id=39:EvenRow"
      _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&H80FF00&"
      _StyleDefs(83)  =   "Named:id=40:OddRow"
      _StyleDefs(84)  =   ":id=40,.parent=33"
      _StyleDefs(85)  =   "Named:id=41:RecordSelector"
      _StyleDefs(86)  =   ":id=41,.parent=34"
      _StyleDefs(87)  =   "Named:id=42:FilterBar"
      _StyleDefs(88)  =   ":id=42,.parent=33"
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid2 
      Height          =   2565
      Left            =   1125
      TabIndex        =   3
      Top             =   1125
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   4524
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "種別"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "使用月"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "回答納期"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "注文№"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "在庫数"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "引当数"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "残数"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3784"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3651"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=131588"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1773"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1640"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=131585"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1931"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1799"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=131585"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2302"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2170"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=139780"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2117"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1984"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=139778"
      Splits(0)._ColumnProps(25)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(5).Width=2117"
      Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=1984"
      Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=131586"
      Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(32)=   "Column(6).Width=2117"
      Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=1984"
      Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=131586"
      Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
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
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "子部品　在庫情報"
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
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=88,.parent=2,.alignment=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=89,.parent=3,.alignment=2"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=91,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=94,.parent=9,.bgcolor=&H80FF80&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=95,.parent=10,.bgcolor=&H80FF00&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=87,.alignment=3"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=87,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=87,.alignment=2"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=106,.parent=87,.alignment=3,.locked=-1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=103,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=104,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=105,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=16,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=13,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=14,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=15,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=20,.parent=87,.alignment=1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=17,.parent=88"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=18,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=19,.parent=91"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=24,.parent=87,.alignment=1"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=21,.parent=88"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=22,.parent=89"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=23,.parent=91"
      _StyleDefs(64)  =   "Named:id=33:Normal"
      _StyleDefs(65)  =   ":id=33,.parent=0"
      _StyleDefs(66)  =   "Named:id=34:Heading"
      _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(68)  =   ":id=34,.wraptext=-1"
      _StyleDefs(69)  =   "Named:id=35:Footing"
      _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(71)  =   "Named:id=36:Selected"
      _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=37:Caption"
      _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(75)  =   "Named:id=38:HighlightRow"
      _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=39:EvenRow"
      _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&H80FF00&"
      _StyleDefs(79)  =   "Named:id=40:OddRow"
      _StyleDefs(80)  =   ":id=40,.parent=33"
      _StyleDefs(81)  =   "Named:id=41:RecordSelector"
      _StyleDefs(82)  =   ":id=41,.parent=34"
      _StyleDefs(83)  =   "Named:id=42:FilterBar"
      _StyleDefs(84)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "データ日時"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   7575
      TabIndex        =   9
      Top             =   750
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "実行ﾊﾟｿｺﾝ"
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
      Left            =   7575
      TabIndex        =   8
      Top             =   375
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "子品番"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   675
      Width           =   720
   End
End
Attribute VB_Name = "ODR10104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'コンボ用添字
'Private Const pcmbHBUN = 0

'テキストボックス添字
Private Const ptxKO_HINBAN = 0
Private Const ptxPC_NM = 1
Private Const ptxTime_Stump = 2


'グリッド用定義
Private ORDR_GRID   As New XArrayDB


Private Const Min_Row% = 1              '最小行数
Private Const Max_Row = 9999            '最大行数

'親情報用
Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 7              '最大列数

Private Const Col_KAITO% = 0            '回答納期
Private Const Col_NOUKI% = 1            '注文納期
Private Const Col_USE_YM% = 2           '使用月
Private Const Col_OYA_HIN% = 3          '親品番
Private Const Col_ODR_NO% = 4           '注文№
Private Const Col_ALL_QTY% = 5          '展開数
Private Const Col_USE_QTY% = 6          '引当数
Private Const Col_FUSOKU_QTY% = 7       '不足数



'子品番情報用
Private Const Min_Col_Ko% = 0              '最小列数
Private Const Max_Col_Ko% = 7              '最大列数

Private Const Col_KUBUN% = 0            '在庫区分
Private Const Col_USE_YM_Ko% = 1        '使用月
Private Const Col_TAISYO% = 2           '対象日付
Private Const Col_ODR_NO_Ko% = 3        '注文№
Private Const Col_MOTO_QTY% = 4         '元数
Private Const Col_HIKI_QTY% = 5         '引当数
Private Const Col_ZAN_QTY% = 6          '残数


Dim row         As Long                 '対象　行

Dim Cor_Row     As Long                 'カレント行

Dim Init_F      As Integer
Private Function Data_Disp() As Integer
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim svJGYOBU    As String * 1
Dim svNAIGAI    As String * 1
Dim svHin_gai   As String * 20

Dim sumQty      As Double
Dim sumReq      As Double

Dim W_TimeStump As String

    Data_Disp = True
    
    row = Min_Row - 1
    Call Input_Lock                             '画面項目ロック
    
    DoEvents
    
    Set ORDR_GRID = Nothing
    
    If ODR_TEMP1_Open(BtOpenExec) Then
        MsgBox "処理を中断します。<TEMP1>", vbExclamation
        GoTo Err_Exit
    End If
    
    If ODR_TEMP2_Open(BtOpenExec) Then
        MsgBox "処理を中断します。<TEMP2>", vbExclamation
        GoTo Err_Exit
    End If
    
    row = 0
    
    Call UniCode_Conv(K0_ODR_TEMP2.KO_JGYOBU, Key_Ko_JIGYOBU)   '事業部
    Call UniCode_Conv(K0_ODR_TEMP2.KO_NAIGAI, Key_Ko_NAIGAI)    '国内外
    Call UniCode_Conv(K0_ODR_TEMP2.KO_HIN_GAI, Key_Ko_HinGai)   '子品番
    Call UniCode_Conv(K0_ODR_TEMP2.IO_KB, "")                   'io区分
    Call UniCode_Conv(K0_ODR_TEMP2.USE_YM, "")                  '使用月
    Call UniCode_Conv(K0_ODR_TEMP2.ANS_NOUKI_DT, "")            '対象日付   YYYYMMDD    (回答納期）
    Call UniCode_Conv(K0_ODR_TEMP2.ORDER_NO, "")                '注文№

    com = BtOpGetGreaterEqual
        
    Do
        sts = BTRV(com, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                Exit Do
                        
            Case Else
                Call File_Error(sts, com, "ODR_TEMP2")
                GoTo Err_Exit
        End Select
        
        If Trim(StrConv(ODR_TP2_R.KO_JGYOBU, vbUnicode)) <> Trim(Key_Ko_JIGYOBU) Then
            Exit Do
        End If
        
        If Trim(StrConv(ODR_TP2_R.KO_NAIGAI, vbUnicode)) <> Trim(Key_Ko_NAIGAI) Then
            Exit Do
        End If
        
        If Trim(StrConv(ODR_TP2_R.KO_HIN_GAI, vbUnicode)) <> Trim(Key_Ko_HinGai) Then
            Exit Do
        End If
    
                '編集
        row = row + 1
        If Grid_Set_Proc_Ko() Then
            GoTo Err_Exit
        End If
        
        com = BtOpGetNext
    Loop
    
    
    Set TDBGrid2.Array = ORDR_GRID
    
    
    TDBGrid2.ReBind
    TDBGrid2.Update
    TDBGrid2.MoveFirst
    TDBGrid2.ScrollBars = dbgAutomatic
    
    DoEvents
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>   親情報
    
    
    row = 0
    
    Call UniCode_Conv(K4_ODR_TEMP1.KO_JGYOBU, Key_Ko_JIGYOBU)   '事業部
    Call UniCode_Conv(K4_ODR_TEMP1.KO_NAIGAI, Key_Ko_NAIGAI)    '国内外
    Call UniCode_Conv(K4_ODR_TEMP1.KO_HIN_GAI, Key_Ko_HinGai)   '子品番
    Call UniCode_Conv(K4_ODR_TEMP1.KAN_KB, "1")                 '親品番　完了区分
    
    Call UniCode_Conv(K4_ODR_TEMP1.KAITO_DT, "")                '親注文の回答納期
    Call UniCode_Conv(K4_ODR_TEMP1.CYUMON_DT, "")               '部材センター注文納期（YYYYMMDD）
    Call UniCode_Conv(K4_ODR_TEMP1.USE_YM, "")                  '使用月
    Call UniCode_Conv(K4_ODR_TEMP1.SHIMUKE, "")                 '仕向け先
    Call UniCode_Conv(K4_ODR_TEMP1.JGYOBU, "")                  '事業部
    Call UniCode_Conv(K4_ODR_TEMP1.NAIGAI, "")                  '国内外
    Call UniCode_Conv(K4_ODR_TEMP1.HIN_GAI, "")                 '親品番
    Call UniCode_Conv(K4_ODR_TEMP1.ORDER_NO, "")                '親品番　注文№
    Call UniCode_Conv(K4_ODR_TEMP1.INS_NO, "")                  '登録順
    Call UniCode_Conv(K4_ODR_TEMP1.BUN_NO, "")                  '分納回数
    
    com = BtOpGetGreaterEqual
        
    Do
        sts = BTRV(com, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K4_ODR_TEMP1, Len(K4_ODR_TEMP1), 4)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                Exit Do
                        
            Case Else
                Call File_Error(sts, com, "ODR_TEMP1")
                GoTo Err_Exit
        End Select
        
        If Trim(StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode)) <> Trim(Key_Ko_JIGYOBU) Then
            Exit Do
        End If
        
        If Trim(StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode)) <> Trim(Key_Ko_NAIGAI) Then
            Exit Do
        End If
        
        If Trim(StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode)) <> Trim(Key_Ko_HinGai) Then
            Exit Do
        End If
        
        If Trim(StrConv(ODR_TP1_R.KAN_KB, vbUnicode)) <> "1" Then
            Exit Do
        End If
        
        W_TimeStump = "20"
        W_TimeStump = W_TimeStump & Left(StrConv(ODR_TP1_R.UPDT_DT, vbUnicode), 2) & "/"
        W_TimeStump = W_TimeStump & Mid(StrConv(ODR_TP1_R.UPDT_DT, vbUnicode), 3, 2) & "/"
        W_TimeStump = W_TimeStump & Right(StrConv(ODR_TP1_R.UPDT_DT, vbUnicode), 2) & "["
        W_TimeStump = W_TimeStump & Left(StrConv(ODR_TP1_R.UPDT_TM, vbUnicode), 2) & ":"
        W_TimeStump = W_TimeStump & Right(StrConv(ODR_TP1_R.UPDT_TM, vbUnicode), 2) & "]"
        Text1(ptxTime_Stump) = W_TimeStump
                '編集
        row = row + 1
        If Grid_Set_Proc() Then
            GoTo Err_Exit
        End If
        
        com = BtOpGetNext
    Loop
    
    
    Set TDBGrid1.Array = ORDR_GRID
    
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    
    DoEvents
    
    Call Input_UnLock                             '画面項目ロック
    
    Data_Disp = False
    
Err_Exit:
    
    sts = BTRV(BtOpClose, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K0_ODR_TEMP1, Len(K0_ODR_TEMP1), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_TEMP1")
        End If
    End If
    
    sts = BTRV(BtOpClose, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_TEMP2")
        End If
    End If
    
End Function
Private Function Grid_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   グリッド表示（移動歴データ内容）
'               Row   行数
'               mode　FALSE:ﾁｪｯｸOFF  TRUE:ﾁｪｯｸON
'----------------------------------------------------------------------------
Dim W_QTY       As Double
Dim W_STR       As String


    Grid_Set_Proc = True

    ORDR_GRID.ReDim Min_Row, row, Min_Col, Max_Col
    
    '回答納期
    If Trim(StrConv(ODR_TP1_R.KAITO_DT, vbUnicode)) = "" Then
        W_STR = ""
    Else
        If Trim(StrConv(ODR_TP1_R.KAITO_DT, vbUnicode)) = "99999999" Then
            W_STR = ""
        Else
            W_STR = Mid(StrConv(ODR_TP1_R.KAITO_DT, vbUnicode), 3, 2) & "/" & _
                    Mid(StrConv(ODR_TP1_R.KAITO_DT, vbUnicode), 5, 2) & "/" _
                        & Right(StrConv(ODR_TP1_R.KAITO_DT, vbUnicode), 2)
        End If
        
    End If
    ORDR_GRID(row, Col_KAITO) = W_STR
    
    '注文納期
    If Trim(StrConv(ODR_TP1_R.CYUMON_DT, vbUnicode)) = "" Then
        W_STR = ""
    Else
        W_STR = Mid(StrConv(ODR_TP1_R.CYUMON_DT, vbUnicode), 3, 2) & "/" & _
                Mid(StrConv(ODR_TP1_R.CYUMON_DT, vbUnicode), 5, 2) & "/" _
                    & Right(StrConv(ODR_TP1_R.CYUMON_DT, vbUnicode), 2)
    End If
    ORDR_GRID(row, Col_NOUKI) = W_STR
    
    '使用月
    If Trim(StrConv(ODR_TP1_R.USE_YM, vbUnicode)) = "" Then
        W_STR = ""
    Else
        W_STR = Left(StrConv(ODR_TP1_R.USE_YM, vbUnicode), 4) & "/" & _
                    Right(StrConv(ODR_TP1_R.USE_YM, vbUnicode), 2)
    End If
    ORDR_GRID(row, Col_USE_YM) = W_STR
    
    '親品番
    ORDR_GRID(row, Col_OYA_HIN) = Trim(StrConv(ODR_TP1_R.HIN_GAI, vbUnicode))
    
    '注文№
    ORDR_GRID(row, Col_ODR_NO) = Trim(StrConv(ODR_TP1_R.ORDER_NO, vbUnicode))
    
    '展開数
    W_QTY = CDbl(StrConv(ODR_TP1_R.ALL_QTY, vbUnicode))
    ORDR_GRID(row, Col_ALL_QTY) = Format(W_QTY, "###,##0.00")
    
    '引当数
    W_QTY = CDbl(StrConv(ODR_TP1_R.ALL_QTY, vbUnicode)) - CDbl(StrConv(ODR_TP1_R.FUSOKU_QTY, vbUnicode))
    ORDR_GRID(row, Col_USE_QTY) = Format(W_QTY, "###,##0.00")
    
    
    '不足数
                                                            '09/03/11 0の時、空白にしてみた。(^_^;)
    W_QTY = CDbl(StrConv(ODR_TP1_R.FUSOKU_QTY, vbUnicode))
    If W_QTY <> 0 Then
        ORDR_GRID(row, Col_FUSOKU_QTY) = Format(W_QTY, "###,##0.00")
    Else
        ORDR_GRID(row, Col_FUSOKU_QTY) = ""
    End If
    
    
    Grid_Set_Proc = False

End Function
Private Function Grid_Set_Proc_Ko() As Integer
'----------------------------------------------------------------------------
'                   グリッド表示（移動歴データ内容）
'               Row   行数
'               mode　FALSE:ﾁｪｯｸOFF  TRUE:ﾁｪｯｸON
'----------------------------------------------------------------------------
Dim W_QTY       As Double
Dim W_STR       As String


    Grid_Set_Proc_Ko = True

    ORDR_GRID.ReDim Min_Row, row, Min_Col, Max_Col
    
    'IO区分
    Select Case StrConv(ODR_TP2_R.IO_KB, vbUnicode)
        Case "a"
            W_STR = "在庫データ"
        Case "b"
            W_STR = "マイナス注文"
        Case "c"
            W_STR = "仕入済（準在庫）"
        Case "d"
            W_STR = "在訂±"
        Case "e"
            W_STR = "半製品"
        Case "f"
            W_STR = "仕入残"
        Case "g"
            W_STR = "仕入残(引当不可)"
            
        Case Else
            W_STR = ""
    End Select
    ORDR_GRID(row, Col_KUBUN) = W_STR
    
    '使用月
    If Trim(StrConv(ODR_TP2_R.USE_YM, vbUnicode)) = "" Then
        W_STR = ""
    Else
        W_STR = Left(StrConv(ODR_TP2_R.USE_YM, vbUnicode), 4) & "/" & _
                    Right(StrConv(ODR_TP2_R.USE_YM, vbUnicode), 2)
    End If
    ORDR_GRID(row, Col_USE_YM_Ko) = W_STR
    
'Private Const Col_ODR_NO_Ko% = 3        '注文№
'Private Const Col_MOTO_QTY% = 4         '元数
'Private Const Col_HIKI_QTY% = 5         '引当数
'Private Const Col_ZAN_QTY% = 6          '残数
    
    '対象日付（回答納期）
    If Trim(StrConv(ODR_TP2_R.ANS_NOUKI_DT, vbUnicode)) = "" Then
        W_STR = ""
        If StrConv(ODR_TP2_R.IO_KB, vbUnicode) = "g" Then
            W_STR = "未設定"
        End If
    Else
        If Trim(StrConv(ODR_TP2_R.ANS_NOUKI_DT, vbUnicode)) = "99999999" Then
            W_STR = ""
        Else
            W_STR = Mid(StrConv(ODR_TP2_R.ANS_NOUKI_DT, vbUnicode), 3, 2) & "/" & _
                    Mid(StrConv(ODR_TP2_R.ANS_NOUKI_DT, vbUnicode), 5, 2) & "/" _
                        & Right(StrConv(ODR_TP2_R.ANS_NOUKI_DT, vbUnicode), 2)
        End If
        
    End If
    ORDR_GRID(row, Col_TAISYO%) = W_STR
    
    
    '注文№
    ORDR_GRID(row, Col_ODR_NO_Ko) = Trim(StrConv(ODR_TP2_R.ORDER_NO, vbUnicode))
    '元数
    W_QTY = CDbl(StrConv(ODR_TP2_R.MOTO_QTY, vbUnicode))
    ORDR_GRID(row, Col_MOTO_QTY) = Format(W_QTY, "###,##0.00")
    
    '引当数
    W_QTY = CDbl(StrConv(ODR_TP2_R.MOTO_QTY, vbUnicode)) - CDbl(StrConv(ODR_TP2_R.ZAI_QTY, vbUnicode))
    ORDR_GRID(row, Col_HIKI_QTY) = Format(W_QTY, "###,##0.00")
    
    
    '残数
                                                            '0の時、空白にしてみた。(^_^;)
    W_QTY = CDbl(StrConv(ODR_TP2_R.ZAI_QTY, vbUnicode))
    If W_QTY <> 0 Then
        ORDR_GRID(row, Col_ZAN_QTY) = Format(W_QTY, "###,##0.00")
    Else
        ORDR_GRID(row, Col_ZAN_QTY) = ""
    End If
    
    
    Grid_Set_Proc_Ko = False

End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    ODR10104.MousePointer = vbHourglass

    Call Ctrl_Lock(ODR10104)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(ODR10104)


    ODR10104.MousePointer = vbDefault

End Sub

Private Sub Command1_Click(Index As Integer)

    Select Case Index
    
            
        Case 0
            Init_F = 0
            Set ORDR_GRID = Nothing
            Set TDBGrid1.Array = ORDR_GRID
            TDBGrid1.ReBind
            TDBGrid1.Update
            DoEvents
            
            'ODR10104_Return = True                '確認画面ｷｬﾝｾﾙ終了
            
            Me.Visible = False
    
    End Select

End Sub

Private Sub Form_Activate()
Dim X_i As Integer

    If Init_F <> 0 Then Exit Sub
    
    ODR10104.Top = ODR10101.Top + (ODR10101.Height - ODR10104.Height)
    ODR10104.Left = ODR10101.Left + (ODR10101.Width - ODR10104.Width) / 2
    
    Text1(ptxKO_HINBAN) = Trim(Key_Ko_HinGai)
    
    If Data_Disp Then
        Call Input_UnLock                             '画面項目ロック
    End If
    
    'ODR10104_Return = True
    TDBGrid1.SetFocus
    
    
    Init_F = 1
    
End Sub

Private Sub Form_Load()
Dim c           As String * 128
Dim W_PC        As String

    Init_F = 0
    c = Space(255)
    If GetComputerNameA(c, 255) <> 0 Then
        W_PC = Left(c, InStr(c, vbNullChar) - 1)
    Else
        W_PC = "000"
    End If
    Text1(ptxPC_NM) = Trim(W_PC)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Me.Visible = False
    
End Sub


Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    TDBGrid1.Bookmark = -1
    
End Sub
