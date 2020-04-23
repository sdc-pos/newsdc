VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form ODR10101 
   Caption         =   "親部品　注文情報登録"
   ClientHeight    =   11730
   ClientLeft      =   165
   ClientTop       =   390
   ClientWidth     =   15210
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11730
   ScaleWidth      =   15210
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H8000000F&
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
      Left            =   13680
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   16
      Text            =   "YY/MM/DD"
      Top             =   840
      Width           =   1155
   End
   Begin VB.OptionButton Option1 
      Caption         =   "子部品情報"
      Height          =   255
      Index           =   1
      Left            =   9000
      TabIndex        =   15
      Top             =   900
      Width           =   1995
   End
   Begin VB.OptionButton Option1 
      Caption         =   "分納情報"
      Height          =   255
      Index           =   0
      Left            =   7080
      TabIndex        =   13
      Top             =   900
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "データ出力"
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
      Index           =   3
      Left            =   4140
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "展　開"
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
      Left            =   2160
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   1800
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Index           =   0
      Left            =   9000
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   2055
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
      Left            =   4860
      MaxLength       =   7
      TabIndex        =   2
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
      Left            =   900
      MaxLength       =   5
      TabIndex        =   1
      Top             =   780
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   13260
      ScaleHeight     =   195
      ScaleWidth      =   270
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   330
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
      Left            =   6120
      TabIndex        =   6
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   1800
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   10035
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   17701
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "№"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   4
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "削除"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "親部品注文           №"
      Columns(2).DataField=   ""
      Columns(2).DataWidth=   10
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "分納"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "親部品"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "部品名"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "数 量"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "部材ｾﾝﾀｰ　注文納期"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "　　組立　　可能日付"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "  親部品 　回答納期"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "使用月"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "完了日付"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Key_登録№"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Key_注文№"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Key_親品番"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "完了M"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "変更前　注文納期"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "変更前　回答納期"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   18
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=18"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=714"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=582"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8721"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=714"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=582"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=529"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2646"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2514"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=532"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=1058"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=926"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=8721"
      Splits(0)._ColumnProps(21)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=3175"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=3043"
      Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=532"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(28)=   "Column(5).Width=3519"
      Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=3387"
      Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=8724"
      Splits(0)._ColumnProps(32)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(34)=   "Column(6).Width=1402"
      Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=1270"
      Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=530"
      Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(39)=   "Column(7).Width=2461"
      Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=2328"
      Splits(0)._ColumnProps(42)=   "Column(7)._ColStyle=529"
      Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(44)=   "Column(8).Width=2461"
      Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=2328"
      Splits(0)._ColumnProps(47)=   "Column(8)._ColStyle=529"
      Splits(0)._ColumnProps(48)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(49)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(50)=   "Column(9).Width=2461"
      Splits(0)._ColumnProps(51)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(9)._WidthInPix=2328"
      Splits(0)._ColumnProps(53)=   "Column(9)._ColStyle=529"
      Splits(0)._ColumnProps(54)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(55)=   "Column(10).Width=1931"
      Splits(0)._ColumnProps(56)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(10)._WidthInPix=1799"
      Splits(0)._ColumnProps(58)=   "Column(10)._ColStyle=529"
      Splits(0)._ColumnProps(59)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(60)=   "Column(11).Width=2461"
      Splits(0)._ColumnProps(61)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(11)._WidthInPix=2328"
      Splits(0)._ColumnProps(63)=   "Column(11)._ColStyle=8721"
      Splits(0)._ColumnProps(64)=   "Column(11).AllowFocus=0"
      Splits(0)._ColumnProps(65)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(66)=   "Column(12).Width=1773"
      Splits(0)._ColumnProps(67)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(12)._WidthInPix=1640"
      Splits(0)._ColumnProps(69)=   "Column(12)._ColStyle=8724"
      Splits(0)._ColumnProps(70)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(71)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(72)=   "Column(13).Width=1773"
      Splits(0)._ColumnProps(73)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(74)=   "Column(13)._WidthInPix=1640"
      Splits(0)._ColumnProps(75)=   "Column(13)._ColStyle=532"
      Splits(0)._ColumnProps(76)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(77)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(78)=   "Column(14).Width=1773"
      Splits(0)._ColumnProps(79)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(80)=   "Column(14)._WidthInPix=1640"
      Splits(0)._ColumnProps(81)=   "Column(14)._ColStyle=532"
      Splits(0)._ColumnProps(82)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(83)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(84)=   "Column(15).Width=873"
      Splits(0)._ColumnProps(85)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(86)=   "Column(15)._WidthInPix=741"
      Splits(0)._ColumnProps(87)=   "Column(15)._ColStyle=532"
      Splits(0)._ColumnProps(88)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(89)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(90)=   "Column(16).Width=2461"
      Splits(0)._ColumnProps(91)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(92)=   "Column(16)._WidthInPix=2328"
      Splits(0)._ColumnProps(93)=   "Column(16)._ColStyle=8724"
      Splits(0)._ColumnProps(94)=   "Column(16).Visible=0"
      Splits(0)._ColumnProps(95)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(96)=   "Column(17).Width=2461"
      Splits(0)._ColumnProps(97)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(98)=   "Column(17)._WidthInPix=2328"
      Splits(0)._ColumnProps(99)=   "Column(17)._ColStyle=8724"
      Splits(0)._ColumnProps(100)=   "Column(17).Visible=0"
      Splits(0)._ColumnProps(101)=   "Column(17).Order=18"
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
      Caption         =   "親部品　注文情報"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.valignment=2,.bgcolor=&HFF0000&,.bold=0"
      _StyleDefs(7)   =   ":id=1,.fontsize=1125,.italic=0,.underline=0,.strikethrough=0,.charset=128"
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
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFF0000&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFF00&"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=87,.parent=1,.bgcolor=&H80FF00&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=96,.parent=4,.namedParent=37,.bgcolor=&H80FF00&"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=88,.parent=2,.alignment=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=89,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=91,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=94,.parent=9,.namedParent=39,.bgcolor=&H80FF80&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=95,.parent=10,.namedParent=40,.bgcolor=&H80FF00&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=43,.parent=87,.alignment=2,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=30,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=31,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=32,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=25,.parent=87,.alignment=2,.bgcolor=&H80000005&"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=22,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=23,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=24,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=102,.parent=87,.bgcolor=&H80000005&"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=99,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=100,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=101,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=17,.parent=87,.alignment=2,.locked=-1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=14,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=15,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=16,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=110,.parent=87,.bgcolor=&H80000005&"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=107,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=108,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=109,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=29,.parent=87,.locked=-1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=26,.parent=88"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=27,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=28,.parent=91"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=114,.parent=87,.alignment=1,.bgcolor=&H80000005&"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=111,.parent=88"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=112,.parent=89"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=113,.parent=91"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=118,.parent=87,.alignment=2,.bgcolor=&H80000005&"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=115,.parent=88"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=116,.parent=89"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=117,.parent=91"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=21,.parent=87,.alignment=2"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=18,.parent=88"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=19,.parent=89"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=20,.parent=91"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=126,.parent=87,.alignment=2,.bgcolor=&H80000005&"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=123,.parent=88"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=124,.parent=89"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=125,.parent=91"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=130,.parent=87,.alignment=2,.bgcolor=&H80000005&"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=127,.parent=88"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=128,.parent=89"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=129,.parent=91"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=134,.parent=87,.alignment=2,.bgcolor=&H80FF00&"
      _StyleDefs(81)  =   ":id=134,.locked=-1"
      _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=131,.parent=88"
      _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=132,.parent=89"
      _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=133,.parent=91"
      _StyleDefs(85)  =   "Splits(0).Columns(12).Style:id=138,.parent=87,.locked=-1"
      _StyleDefs(86)  =   "Splits(0).Columns(12).HeadingStyle:id=135,.parent=88"
      _StyleDefs(87)  =   "Splits(0).Columns(12).FooterStyle:id=136,.parent=89"
      _StyleDefs(88)  =   "Splits(0).Columns(12).EditorStyle:id=137,.parent=91"
      _StyleDefs(89)  =   "Splits(0).Columns(13).Style:id=47,.parent=87"
      _StyleDefs(90)  =   "Splits(0).Columns(13).HeadingStyle:id=44,.parent=88"
      _StyleDefs(91)  =   "Splits(0).Columns(13).FooterStyle:id=45,.parent=89"
      _StyleDefs(92)  =   "Splits(0).Columns(13).EditorStyle:id=46,.parent=91"
      _StyleDefs(93)  =   "Splits(0).Columns(14).Style:id=51,.parent=87"
      _StyleDefs(94)  =   "Splits(0).Columns(14).HeadingStyle:id=48,.parent=88"
      _StyleDefs(95)  =   "Splits(0).Columns(14).FooterStyle:id=49,.parent=89"
      _StyleDefs(96)  =   "Splits(0).Columns(14).EditorStyle:id=50,.parent=91"
      _StyleDefs(97)  =   "Splits(0).Columns(15).Style:id=55,.parent=87"
      _StyleDefs(98)  =   "Splits(0).Columns(15).HeadingStyle:id=52,.parent=88"
      _StyleDefs(99)  =   "Splits(0).Columns(15).FooterStyle:id=53,.parent=89"
      _StyleDefs(100) =   "Splits(0).Columns(15).EditorStyle:id=54,.parent=91"
      _StyleDefs(101) =   "Splits(0).Columns(16).Style:id=59,.parent=87,.locked=-1"
      _StyleDefs(102) =   "Splits(0).Columns(16).HeadingStyle:id=56,.parent=88"
      _StyleDefs(103) =   "Splits(0).Columns(16).FooterStyle:id=57,.parent=89"
      _StyleDefs(104) =   "Splits(0).Columns(16).EditorStyle:id=58,.parent=91"
      _StyleDefs(105) =   "Splits(0).Columns(17).Style:id=63,.parent=87,.locked=-1"
      _StyleDefs(106) =   "Splits(0).Columns(17).HeadingStyle:id=60,.parent=88"
      _StyleDefs(107) =   "Splits(0).Columns(17).FooterStyle:id=61,.parent=89"
      _StyleDefs(108) =   "Splits(0).Columns(17).EditorStyle:id=62,.parent=91"
      _StyleDefs(109) =   "Named:id=33:Normal"
      _StyleDefs(110) =   ":id=33,.parent=0"
      _StyleDefs(111) =   "Named:id=34:Heading"
      _StyleDefs(112) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(113) =   ":id=34,.wraptext=-1"
      _StyleDefs(114) =   "Named:id=35:Footing"
      _StyleDefs(115) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(116) =   "Named:id=36:Selected"
      _StyleDefs(117) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(118) =   "Named:id=37:Caption"
      _StyleDefs(119) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(120) =   "Named:id=38:HighlightRow"
      _StyleDefs(121) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(122) =   "Named:id=39:EvenRow"
      _StyleDefs(123) =   ":id=39,.parent=33,.bgcolor=&H80FF00&"
      _StyleDefs(124) =   "Named:id=40:OddRow"
      _StyleDefs(125) =   ":id=40,.parent=33,.bgcolor=&HFF0000&"
      _StyleDefs(126) =   "Named:id=41:RecordSelector"
      _StyleDefs(127) =   ":id=41,.parent=34"
      _StyleDefs(128) =   "Named:id=42:FilterBar"
      _StyleDefs(129) =   ":id=42,.parent=33"
      _StyleDefs(130) =   "Named:id=13:IO_OK"
      _StyleDefs(131) =   ":id=13,.parent=42,.bgcolor=&H80000005&"
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   6750
      TabIndex        =   14
      Top             =   675
      Width           =   4455
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "繰越日"
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
      Left            =   12900
      TabIndex        =   17
      Top             =   900
      Width           =   720
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "仕向先"
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
      Index           =   3
      Left            =   8220
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
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
      Left            =   4080
      TabIndex        =   10
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
      Left            =   1680
      TabIndex        =   9
      Top             =   840
      Width           =   2235
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
      Left            =   120
      TabIndex        =   8
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
         Caption         =   "展開"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "データ出力"
         Index           =   2
      End
      Begin VB.Menu SHORI 
         Caption         =   "画面印刷"
         Index           =   3
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   4
      End
   End
End
Attribute VB_Name = "ODR10101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'コンボ用添字
'Private Const pcmbJI = 0            '事業部
Private Const pcmbSM = 0            '仕向け先


'テキスト用添字
Private Const ptxTOP% = 0
Private Const ptxLAST% = 1

Private Const ptxTANTO_CD% = 0
Private Const ptxUSE_YY% = 1
Private Const ptxSHIME_DT% = 2

'ラベル用添字
Private Const plabTANTO_NM% = 0

'コマンドボタン用添字
Private Const FuncCOR% = 0       '更新
Private Const FuncEND% = 1       '終了
Private Const FuncREQ% = 2       '展開
Private Const FuncOUT% = 3       'データ出力

'ListBox添字
'Private Const plstSRCH% = 0         '


'グリッド更新マーク
Dim Grid_Cor_M      As Integer
Dim Grid_Req_M      As Integer
Dim Data_Out_Need   As Integer

'グリッド用定義
Private ORDR_GRID   As New XArrayDB

Private Const Min_Row% = 1                  '最小行数
'Private Max_Row As Long                    '最大表示行数
Private Const Max_Row = 9999                '最大行数

Private Const Min_Col% = 0                  '最小列数
Private Const Max_Col% = 17                 '最大列数 15-->17 2016.11.25
    
Private Const Col_No% = 0                   '行№
Private Const Col_DEL% = 1                  '削除マーク
Private Const Col_ORDR_NO% = 2              '親部品　注文№
Private Const Col_BUNNO% = 3                '分納回数
Private Const Col_OYA_ITEM% = 4             '親部品コード
Private Const Col_OYA_NM% = 5               '親部品コード
Private Const Col_ORDR_QTY% = 6             '注文数量
Private Const Col_NOUKI% = 7                '親部品　注文納期
Private Const Col_OK_DT% = 8                '組立可能日
Private Const Col_KAITO% = 9                '親部品　回答納期
Private Const Col_USE_YM% = 10              '使用月
Private Const Col_FIN_DT% = 11              '完了日付
Private Const Col_KEY% = 12                 'データＫｅｙ情報   登録順
Private Const Col_KEY_ORDR% = 13            'データＫｅｙ情報   注文№
Private Const Col_KEY_OYA% = 14             'データＫｅｙ情報   親品番

Private Const Col_KEY_FIN% = 15             '並替えＫｅｙ情報   完了Ｍ      '2012/03/15 追加

Private Const Col_SV_NOUKI% = 16            '親部品　注文納期 (保存)    '2016.11.25
Private Const Col_SV_KAITO% = 17            '親部品　回答納期 (保存)    '2016.11.25



Dim row         As Long                     '対象　行

Dim Cor_Row     As Long                     'カレント行

Dim Init_F_10101      As Integer

Dim GW_MOTO_ORDR As String   '2009.01.15
Dim GW_MOTO_OYA  As String

'Private Const Last_Update$ = "[ODR1010] 2017.01.16 13:15"
Private Const Last_Update$ = "[ODR1010] 2019.01.08 14:15"



Private Function USE_YM_SAVE()
Dim sts         As Integer
Dim com         As Integer
Dim yn          As Integer
Dim W_YYMM      As String

    USE_YM_SAVE = True
    
    W_YYMM = Left(Text1(ptxUSE_YY), 4) & Right(Text1(ptxUSE_YY), 2)
    
    
    Call UniCode_Conv(K1_ODR_ORDER.SHIMUKE, GW_SIMUKE)
    Call UniCode_Conv(K1_ODR_ORDER.JGYOBU, GW_JIGYOBU)
    Call UniCode_Conv(K1_ODR_ORDER.NAIGAI, GW_NAIGAI)
    Call UniCode_Conv(K1_ODR_ORDER.USE_YM, W_YYMM)
    
    Call UniCode_Conv(K1_ODR_ORDER.HIN_GAI, "")
    Call UniCode_Conv(K1_ODR_ORDER.ORDER_NO, "")
    Call UniCode_Conv(K1_ODR_ORDER.INS_NO, "")
    Call UniCode_Conv(K1_ODR_ORDER.BUN_NO, "")
    
    com = BtOpGetGreaterEqual
    Do
        
        sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K1_ODR_ORDER, Len(K1_ODR_ORDER), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                
                Exit Do
            Case Else
                Call File_Error(sts, com, "ODR_ORDER")
                Exit Do
        End Select
        
        
        If Trim(StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode)) <> Trim(GW_SIMUKE) Or _
            Trim(StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU) Or _
            Trim(StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI) Then
            Exit Do
        End If
        
        'If StrConv(ODR_ORDER_REC.USE_YM, vbUnicode) <> W_YYMM Then
        '    Exit Do
        'End If
        
        Call UniCode_Conv(ODR_ORDER_REC.USE_YM_MOTO, StrConv(ODR_ORDER_REC.USE_YM, vbUnicode))
        
        sts = BTRV(BtOpUpdate, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K1_ODR_ORDER, Len(K1_ODR_ORDER), 1)
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, com, "ODR_ORDER")
                Exit Do
        End Select
        
        com = BtOpGetNext
    Loop

    USE_YM_SAVE = False
End Function

Private Function Code_Set_Proc(Index As Integer, KBN As String, Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   コードマスタをコンボにセットする。
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim Key_Len     As Integer
Dim Option1     As Integer
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
            wkOption = Trim(StrConv(P_CODEREC.Option1, vbUnicode))
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

Private Function Grid_Err_Chk(Index As Integer, W_Aft As String) As Integer
'----------------------------------------------------------------------------
'                   グリッド入力内容エラーチェック
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim yn          As Integer
Dim W_STR       As String

    Grid_Err_Chk = True
    
    Select Case Index
        Case Col_DEL            '削除マーク
            
            If ORDR_GRID(Cor_Row, Index) Then
                If Trim(ORDR_GRID(Cor_Row, Col_FIN_DT)) <> "" Then
                    MsgBox Cor_Row & "行目　完了済み→削除不可！", vbExclamation
                    ORDR_GRID(Cor_Row, Index) = False
                    'TDBGrid1.ReBind
                    'TDBGrid1.Update
                        'TDBGrid1.MoveFirst
                    'TDBGrid1.ScrollBars = dbgAutomatic
                    Exit Function
                End If
            End If
            
        
        Case Col_ORDR_NO        '親部品　注文№
        
            GW_MOTO_ORDR = Trim(ORDR_GRID(Cor_Row, Col_KEY_ORDR))        '2009.01.15
            
            
            
            
        Case Col_BUNNO          '分納回数
            
            
            
        Case Col_OYA_ITEM       '親部品コード
            If Trim(W_Aft) = "" Then
                If Trim(ORDR_GRID(Cor_Row, Col_ORDR_NO)) <> "" Then
                    MsgBox Cor_Row & "行目　親部品　未指定エラー！", vbExclamation
                    Exit Function
                End If
            End If
            
            
            GW_MOTO_OYA = Trim(ORDR_GRID(Cor_Row, Col_KEY_OYA))          '2009.01.15

            
            ORDR_GRID(Cor_Row, Col_OYA_ITEM) = StrConv(ORDR_GRID(Cor_Row, Col_OYA_ITEM), vbUpperCase) '2017.01.16




            Call UniCode_Conv(K0_ITEM.JGYOBU, GW_JIGYOBU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, GW_NAIGAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Trim(ORDR_GRID(Cor_Row, Col_OYA_ITEM)))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound       'レコード無し
                    Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
            
            ORDR_GRID(Cor_Row, Col_OYA_NM) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
            
            
            
        Case Col_ORDR_QTY                   '注文数量
            If Trim(W_Aft) <> "" Then
                If Not IsNumeric(W_Aft) Then
                    MsgBox Cor_Row & "行目　注文数量　数値エラー！", vbExclamation
                    Exit Function
                End If
            End If
            
            
            
            
        Case Col_NOUKI                      '親部品　注文納期
            If IsDate(W_Aft) Then
                ORDR_GRID(Cor_Row, Index) = Format(W_Aft, "yyyy/mm/dd")
                
                
                
                'TDBGrid1.ReBind
                'TDBGrid1.Update
                    'TDBGrid1.MoveFirst
                'TDBGrid1.ScrollBars = dbgAutomatic
            
            
            
                '>>>>>>>>>>>>>>>>>  2016.11.25
                If ORDR_GRID(Cor_Row, Col_NOUKI) <> ORDR_GRID(Cor_Row, Col_SV_NOUKI) Then
                    If ORDR_GRID(Cor_Row, Index) < Format(Now, "YYYY/MM/DD") Then
                        MsgBox Cor_Row & "行目　親部品　注文納期　過去日付の入力不可！", vbExclamation
                        Exit Function
                    End If
                End If
                '>>>>>>>>>>>>>>>>>  2016.11.25
            Else
                If CDbl(ORDR_GRID(Cor_Row, Col_ORDR_QTY)) > 0 Then
                    MsgBox Cor_Row & "行目　親部品　注文納期　日付エラー！", vbExclamation
                    Exit Function
                End If
            End If
            
        Case Col_OK_DT                      '組立可能日
            If Trim(W_Aft) <> "" Then
                If IsDate(W_Aft) Then
                    ORDR_GRID(Cor_Row, Index) = Format(W_Aft, "yyyy/mm/dd")
                    'TDBGrid1.ReBind
                    'TDBGrid1.Update
                        'TDBGrid1.MoveFirst
                    'TDBGrid1.ScrollBars = dbgAutomatic
                Else
                    MsgBox Cor_Row & "行目　組立可能日　日付エラー！", vbExclamation
                    Exit Function
                End If
            End If
            
        Case Col_KAITO                      '親部品　回答納期
            If Trim(W_Aft) = "" Then
            
            
            Else
                If IsDate(W_Aft) Then
                    ORDR_GRID(Cor_Row, Index) = Format(W_Aft, "yyyy/mm/dd")
                    'TDBGrid1.ReBind
                    'TDBGrid1.Update
                        'TDBGrid1.MoveFirst
                    'TDBGrid1.ScrollBars = dbgAutomatic
                
                
                
                '>>>>>>>>>>>>>>>>>  2016.11.25
                If ORDR_GRID(Cor_Row, Col_KAITO) <> ORDR_GRID(Cor_Row, Col_SV_KAITO) Then
                    If ORDR_GRID(Cor_Row, Index) < Format(Now, "YYYY/MM/DD") Then
                        MsgBox Cor_Row & "行目　親部品　回答納期　過去日付の入力不可！", vbExclamation
                        Exit Function
                    End If
                End If
                '>>>>>>>>>>>>>>>>>  2016.11.25
                
                
                Else
                    MsgBox Cor_Row & "行目　親部品　回答納期　日付エラー！", vbExclamation
                    Exit Function
                End If
            End If
        Case Col_USE_YM                     '使用月
            If Trim(W_Aft) = "" Then
                ORDR_GRID(Cor_Row, Index) = Text1(ptxUSE_YY)
                    'TDBGrid1.ReBind
                    'TDBGrid1.Update
                        'TDBGrid1.MoveFirst
                    'TDBGrid1.ScrollBars = dbgAutomatic
                MsgBox Cor_Row & "行目　使用月　未設定エラー！", vbExclamation
                Exit Function
                    
            Else
                If IsDate(W_Aft & "/01") Then
                    ORDR_GRID(Cor_Row, Index) = Left(Format(W_Aft & "/01", "yyyy/mm/dd"), 7)
                    'TDBGrid1.ReBind
                    'TDBGrid1.Update
                        'TDBGrid1.MoveFirst
                    'TDBGrid1.ScrollBars = dbgAutomatic
                Else
                    MsgBox Cor_Row & "行目　使用月　日付エラー！", vbExclamation
                    Exit Function
                End If
                
                W_STR = Left(Format(W_Aft & "/01", "yyyy/mm/dd"), 7)
                If W_STR > GW_MAX_YYMM Then
                    W_STR = Left(GW_MAX_YYMM, 4) & "年" & Mid(GW_MAX_YYMM, 6, 2) & "月"
                    MsgBox Cor_Row & "行目　使用月　" & W_STR & "以降エラー！", vbExclamation
                    Exit Function
                End If
                
                W_STR = Left(Format(W_Aft & "/01", "yyyymmdd"), 6)
                If W_STR < GW_TOUGETU Then
                    W_STR = Format(DateAdd("m", -1, Left(GW_TOUGETU, 4) & "/" & Right(GW_TOUGETU, 2) & "/01"), "yyyy/mm/dd")
                    'W_STR = Left(GW_TOUGETU, 4) & "年" & Right(GW_TOUGETU, 2) & "月"
                    W_STR = Left(W_STR, 4) & "年" & Mid(W_STR, 6, 2) & "月"
                    MsgBox Cor_Row & "行目　使用月　" & W_STR & "以前エラー！", vbExclamation
                    Exit Function
                End If
                
                
            End If
            
            
            
            
        Case Col_FIN_DT                     '完了日付
            If Trim(W_Aft) = "" Then
                ORDR_GRID(Cor_Row, Col_KEY_FIN) = "0"                '2012/03/15　追加
            Else
                If IsDate(W_Aft) Then
                    ORDR_GRID(Cor_Row, Index) = Format(W_Aft, "yyyy/mm/dd")
                    ORDR_GRID(Cor_Row, Col_KEY_FIN) = "9"           '2012/03/15 追加
                    'TDBGrid1.ReBind
                    'TDBGrid1.Update
                        'TDBGrid1.MoveFirst
                    'TDBGrid1.ScrollBars = dbgAutomatic
                Else
                    MsgBox Cor_Row & "行目　完了日付　日付エラー！", vbExclamation
                    Exit Function
                End If
            End If
            
        Case Col_KEY                    'Key　№
            
            
    End Select
    
    DoEvents
    
    If Trim(W_Aft) <> "" Then
        Select Case Index
                 '注文数量      '注文納期    '回答納期   '完了日付     '親部品注文№
            Case Col_ORDR_QTY, Col_NOUKI, Col_KAITO, Col_FIN_DT, Col_ORDR_NO
                If Trim(ORDR_GRID(Cor_Row, Col_OYA_ITEM)) = "" Then
                    MsgBox Cor_Row & "行目　親部品　未指定エラー！", vbExclamation
                    Exit Function
                End If
                
            Case Else
    
        End Select
    End If
    

    Grid_Err_Chk = False

End Function
Private Function ERR_CHK(Index As Integer) As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_STR       As String


    ERR_CHK = True
    
                        '入力文字数チェック
    If LenB(StrConv(Text1(Index), vbFromUnicode)) > Text1(Index).MaxLength Then
        MsgBox "入力した項目は（桁あふれエラー）です。", vbExclamation
        Exit Function
    End If
    
    Select Case Index
        Case ptxTANTO_CD
            Lab_Dsp(plabTANTO_NM) = ""
            If Trim(Text1(Index)) = "" Then
                MsgBox "担当者を指定して下さい。", vbExclamation
                Exit Function
            End If
            
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(Index))
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound       'レコード無し
                    MsgBox "担当者　未登録！", vbExclamation
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "TANTO")
                    Exit Function
            End Select
            
            Lab_Dsp(plabTANTO_NM) = Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
            GW_TANTO = Trim(Text1(Index))
            
            GW_SIMUKE = Left(Right(Combo1(pcmbSM).Text, 4), 2)
            'GW_JIGYOBU = Right(Combo1(pcmbJI).Text, 1)
            GW_JIGYOBU = Mid(Right(Combo1(pcmbSM).Text, 4), 3, 1)
            GW_NAIGAI = Right(Combo1(pcmbSM).Text, 1)
            
            
        Case ptxUSE_YY
            If Trim(Text1(Index)) = "" Then
                MsgBox "使用年月を指定して下さい。", vbExclamation
                Exit Function
            End If
            
            
            If Not IsDate(Text1(ptxUSE_YY) & "/01") Then
                MsgBox "使用月エラー！", vbExclamation
                Exit Function
            End If
            
            W_STR = Format(Text1(ptxUSE_YY) & "/01", "yyyy/mm/dd")
            Text1(ptxUSE_YY) = Left(W_STR, 7)
            
            If W_STR > GW_MAX_YYMM Then
                W_STR = Left(GW_MAX_YYMM, 4) & "年" & Mid(GW_MAX_YYMM, 6, 2) & "月"
                MsgBox W_STR & "以降は表示出来ません！", vbExclamation
                Exit Function
            End If
            
            W_STR = Left(Text1(ptxUSE_YY), 4) & Right(Text1(ptxUSE_YY), 2)
            If W_STR < GW_TOUGETU Then
                W_STR = Format(DateAdd("m", -1, Left(GW_TOUGETU, 4) & "/" & Right(GW_TOUGETU, 2) & "/01"), "yyyy/mm/dd")
                'W_STR = Left(GW_TOUGETU, 4) & "年" & Right(GW_TOUGETU, 2) & "月"
                W_STR = Left(W_STR, 4) & "年" & Mid(W_STR, 6, 2) & "月"
                MsgBox W_STR & "以前は表示出来ません！", vbExclamation
                Exit Function
            End If
            
            
            
    End Select
    
    
    ERR_CHK = False
End Function

Private Function Data_Disp() As Integer
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer
    
Dim X_i         As Long
    
Dim W_Key       As String
Dim W_STR       As String

Dim cnt         As Integer

    Data_Disp = True
    
    row = Min_Row - 1
    Call Input_Lock                             '画面項目ロック
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "親部品　注文情報　検索中！　＜Data_Disp＞", Me.hwnd, 0)
    
    
    Set ORDR_GRID = Nothing
    
    
    'Call UniCode_Conv(K1_ODR_ORDER.SHIMUKE, GW_SIMUKE)
    'Call UniCode_Conv(K1_ODR_ORDER.JGYOBU, GW_JIGYOBU)
    'Call UniCode_Conv(K1_ODR_ORDER.NAIGAI, GW_NAIGAI)
    'Call UniCode_Conv(K1_ODR_ORDER.USE_YM, Left(Text1(ptxUSE_YY), 4) & Right(Text1(ptxUSE_YY), 2))
    
    'Call UniCode_Conv(K1_ODR_ORDER.HIN_GAI, "")
    'Call UniCode_Conv(K1_ODR_ORDER.ORDER_NO, "")
    'Call UniCode_Conv(K1_ODR_ORDER.INS_NO, "")
    'Call UniCode_Conv(K1_ODR_ORDER.BUN_NO, "")
    
    '2009/03/12
    Call UniCode_Conv(K5_ODR_ORDER.SHIMUKE, GW_SIMUKE)
    Call UniCode_Conv(K5_ODR_ORDER.JGYOBU, GW_JIGYOBU)
    Call UniCode_Conv(K5_ODR_ORDER.NAIGAI, GW_NAIGAI)
    Call UniCode_Conv(K5_ODR_ORDER.USE_YM, Left(Text1(ptxUSE_YY), 4) & Right(Text1(ptxUSE_YY), 2))
    
    Call UniCode_Conv(K5_ODR_ORDER.INS_DT, "")
    Call UniCode_Conv(K5_ODR_ORDER.INS_TM, "")
    Call UniCode_Conv(K5_ODR_ORDER.HIN_GAI, "")
    Call UniCode_Conv(K5_ODR_ORDER.ORDER_NO, "")
    Call UniCode_Conv(K5_ODR_ORDER.INS_NO, "")
    Call UniCode_Conv(K5_ODR_ORDER.BUN_NO, "")
    
    com = BtOpGetGreaterEqual
    Do
        
        'sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K1_ODR_ORDER, Len(K1_ODR_ORDER), 1)
        sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K5_ODR_ORDER, Len(K5_ODR_ORDER), 5)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                
                Exit Do
            Case Else
                Call File_Error(sts, com, "ODR_ORDER")
                GoTo Err_Exit
        End Select
        
        
        If Trim(StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode)) <> Trim(GW_SIMUKE) Or _
            Trim(StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU) Or _
            Trim(StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI) Then
            Exit Do
        End If
        
        
        DIS_USE_YM = Left(StrConv(ODR_ORDER_REC.USE_YM, vbUnicode), 4) & "/" & _
                            Mid(StrConv(ODR_ORDER_REC.USE_YM, vbUnicode), 5, 2)
        If Left(StrConv(ODR_ORDER_REC.USE_YM, vbUnicode), 4) & "/" & _
                            Mid(StrConv(ODR_ORDER_REC.USE_YM, vbUnicode), 5, 2) <> Trim(Text1(ptxUSE_YY)) Then
            
            Exit Do
    
        End If
    
    
        If CInt(StrConv(ODR_ORDER_REC.BUN_KB, vbUnicode)) = 0 Then
        
            DIS_ORDR_NO = Trim(StrConv(ODR_ORDER_REC.ORDER_NO, vbUnicode))
            
            If Trim(StrConv(ODR_ORDER_REC.BUN_NO, vbUnicode)) <> "" Then
                DIS_BUNNO = CInt(Trim(StrConv(ODR_ORDER_REC.BUN_NO, vbUnicode)))
            Else
                DIS_BUNNO = ""
            End If
            
            DIS_OYA_ITEM = Trim(StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode))
            DIS_ORDR_QTY = CDbl(Trim(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode)))
            
            If Trim(StrConv(ODR_ORDER_REC.CYUMON_DT, vbUnicode)) = "" Then
                DIS_NOUKI = ""
            Else
                DIS_NOUKI = Left(StrConv(ODR_ORDER_REC.CYUMON_DT, vbUnicode), 4) & "/" & _
                                Mid(StrConv(ODR_ORDER_REC.CYUMON_DT, vbUnicode), 5, 2) & "/" & _
                                Right(StrConv(ODR_ORDER_REC.CYUMON_DT, vbUnicode), 2)
            End If
            
            If Trim(StrConv(ODR_ORDER_REC.KUMI_OK_DT, vbUnicode)) = "" Then
                DIS_OK_DT = ""
            Else
                DIS_OK_DT = Left(StrConv(ODR_ORDER_REC.KUMI_OK_DT, vbUnicode), 4) & "/" & _
                                Mid(StrConv(ODR_ORDER_REC.KUMI_OK_DT, vbUnicode), 5, 2) & "/" & _
                                Right(StrConv(ODR_ORDER_REC.KUMI_OK_DT, vbUnicode), 2)
            End If
            
            If Trim(StrConv(ODR_ORDER_REC.KAITO_DT, vbUnicode)) = "" Then
                DIS_KAITO = ""
            Else
                DIS_KAITO = Left(StrConv(ODR_ORDER_REC.KAITO_DT, vbUnicode), 4) & "/" & _
                                Mid(StrConv(ODR_ORDER_REC.KAITO_DT, vbUnicode), 5, 2) & "/" & _
                                Right(StrConv(ODR_ORDER_REC.KAITO_DT, vbUnicode), 2)
            End If
            
            
            
            
            If Trim(StrConv(ODR_ORDER_REC.USE_YM, vbUnicode)) = "" Then
                DIS_USE_YM = ""
            Else
                DIS_USE_YM = Left(StrConv(ODR_ORDER_REC.USE_YM, vbUnicode), 4) & "/" & _
                                Mid(StrConv(ODR_ORDER_REC.USE_YM, vbUnicode), 5, 2)
            End If
            
            If Trim(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode)) = "" Then
                DIS_FIN_DT = ""
            Else
                DIS_FIN_DT = Left(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode), 4) & "/" & _
                                Mid(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode), 5, 2) & "/" & _
                                Right(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode), 2)
            End If
            
            DIS_KEY = Trim(StrConv(ODR_ORDER_REC.INS_NO, vbUnicode))
            
            
            row = row + 1
            If row > Max_Row Then
                MsgBox "最大表示行数を超えました。"
                Exit Do
            End If
                    
            If Grid_Set_Proc() Then
                GoTo Err_Exit
            End If
        
        End If
        
        com = BtOpGetNext
        
    Loop
    
    'X_I = ORDR_GRID.UpperBound(1)
    'MsgBox "Row=" & row & ",UpperBound=" & X_I
    
    
    
    Set TDBGrid1.Array = ORDR_GRID
    
    'TDBGrid1.style.Locked = True
    
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    TDBGrid1.Bookmark = 1
    
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "現在の注文情報　表示中　→　追加登録・修正入力して下さい。", Me.hwnd, 0)
    DoEvents
    
    Data_Disp = False
    
Err_Exit:
    Call Input_UnLock                             '画面項目ロック
    
End Function

Private Function Grid_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   グリッド表示（移動歴データ内容）
'               Row   行数
'               mode　FALSE:ﾁｪｯｸOFF  TRUE:ﾁｪｯｸON
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim yn         As Integer
Dim W_STR       As String

    Grid_Set_Proc = True

    ORDR_GRID.ReDim Min_Row, row, Min_Col, Max_Col

    ORDR_GRID(row, Col_No) = row                    '行№
    
    ORDR_GRID(row, Col_ORDR_NO) = DIS_ORDR_NO       '親部品　注文№
    ORDR_GRID(row, Col_BUNNO) = DIS_BUNNO           '分納回数
    ORDR_GRID(row, Col_OYA_ITEM) = DIS_OYA_ITEM     '親部品コード
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, GW_JIGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, GW_NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, DIS_OYA_ITEM)
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound       'レコード無し
            'MsgBox "品番　未登録！", vbExclamation
            'Exit Function
            Call UniCode_Conv(ITEMREC.HIN_NAME, "未登録")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "TANTO")
            Exit Function
    End Select
    ORDR_GRID(row, Col_OYA_NM) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))

    
    ORDR_GRID(row, Col_ORDR_QTY) = DIS_ORDR_QTY     '注文数量
    
    ORDR_GRID(row, Col_NOUKI) = DIS_NOUKI           '親部品　注文納期
    ORDR_GRID(row, Col_OK_DT) = DIS_OK_DT           '組立可能日
    ORDR_GRID(row, Col_KAITO) = DIS_KAITO           '親部品　回答納期
    ORDR_GRID(row, Col_USE_YM) = DIS_USE_YM         '使用月
    ORDR_GRID(row, Col_FIN_DT) = DIS_FIN_DT         '完了日付
    ORDR_GRID(row, Col_KEY) = DIS_KEY               'データＫｅｙ情報
    
    
    ORDR_GRID(row, Col_KEY_ORDR) = DIS_ORDR_NO       '親部品　注文№
    ORDR_GRID(row, Col_KEY_OYA) = DIS_OYA_ITEM       '親部品コード
    
    
    ORDR_GRID(row, Col_SV_NOUKI) = DIS_NOUKI        '親部品　注文納期       2016.11.25
    ORDR_GRID(row, Col_SV_KAITO) = DIS_KAITO        '親部品　回答納期       2016.11.25
    
    
    '2012/03/15 追加
    If Trim(DIS_FIN_DT) = "" Then
        ORDR_GRID(row, Col_KEY_FIN) = "0"
    Else
        ORDR_GRID(row, Col_KEY_FIN) = "9"
    End If
    

    Grid_Set_Proc = False

End Function

Private Function Rec_UPDT(In_Lock As Integer) As Integer
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim X_i         As Integer

Dim W_Key       As String
Dim W_No        As String
Dim W_STR       As String
Dim W_Date      As String

Dim W_Use_YM    As String

Dim W_QTY       As Double

Dim W_DIS_NO    As Long


    Rec_UPDT = True
    If In_Lock = True Then
        Call Input_Lock
    End If
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "親部品　注文情報　更新中！　＜Rec_UPDT＞", Me.hwnd, 0)
    
    
    X_i = ORDR_GRID.UpperBound(1)
    
    For Cor_Row = Min_Row To ORDR_GRID.UpperBound(1)
        
        W_STR = Trim(ORDR_GRID(Cor_Row, Col_ORDR_NO)) & Trim(ORDR_GRID(Cor_Row, Col_OYA_ITEM))
        
        W_STR = Trim(ORDR_GRID(Cor_Row, Col_OYA_ITEM))      '親品目の未設定は無視！ 2008/10/21
        If W_STR <> "" Then
        
        
            GW_MOTO_ORDR = Trim(ORDR_GRID(Cor_Row, Col_KEY_ORDR))        '2009.01.15
            GW_MOTO_OYA = Trim(ORDR_GRID(Cor_Row, Col_KEY_OYA))
        
        
        
        
            W_QTY = CDbl(Trim(ORDR_GRID(Cor_Row, Col_ORDR_QTY)))
        
            
        'If W_Qty > 0 Then           '注文数量＞０のみ対象！？
        
            GW_HINGAI = Trim(ORDR_GRID(Cor_Row, Col_OYA_ITEM))
            DIS_KEY = Trim(ORDR_GRID(Cor_Row, Col_KEY))
            
            If IsNumeric(DIS_KEY) Then
                
                DIS_KEY = Format(CInt(DIS_KEY), "0000")
                
            Else
                
                DIS_KEY = Format(CInt(Cor_Row), "0000")
            End If
                        
            
            
            ORDR_GRID(Cor_Row, Col_KEY) = DIS_KEY
            'TDBGrid1.ReBind
            'TDBGrid1.Update
                        
            DIS_ORDR_NO = Trim(ORDR_GRID(Cor_Row, Col_ORDR_NO))
            
            W_No = Trim(ORDR_GRID(Cor_Row, Col_BUNNO))
            If W_No = "" Then
                DIS_BUNNO = ""
            Else
                If IsNumeric(W_No) Then
                    
                    DIS_BUNNO = Format(CInt(W_No), "000")
                                      
                Else
                    DIS_BUNNO = ""
                End If
            End If
            
            Call UniCode_Conv(K0_ODR_ORDER.SHIMUKE, GW_SIMUKE)
            Call UniCode_Conv(K0_ODR_ORDER.JGYOBU, GW_JIGYOBU)
            Call UniCode_Conv(K0_ODR_ORDER.NAIGAI, GW_NAIGAI)
            Call UniCode_Conv(K0_ODR_ORDER.HIN_GAI, GW_HINGAI)
            Call UniCode_Conv(K0_ODR_ORDER.INS_NO, DIS_KEY)
            Call UniCode_Conv(K0_ODR_ORDER.ORDER_NO, DIS_ORDR_NO)
            Call UniCode_Conv(K0_ODR_ORDER.BUN_NO, DIS_BUNNO)
            
            
            
            Call UniCode_Conv(K0_ODR_ORDER.HIN_GAI, GW_MOTO_OYA)             '2009.01.15
            Call UniCode_Conv(K0_ODR_ORDER.ORDER_NO, GW_MOTO_ORDR)
            com = BtOpUpdate
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_ORDER, Len(K0_ODR_ORDER), 0)
                Select Case sts
                    Case BtNoErr
                        If ORDR_GRID(Cor_Row, Col_DEL) = True Then
                            com = BtOpDelete
                        Else
                            If GW_HINGAI = GW_MOTO_OYA And DIS_ORDR_NO = GW_MOTO_ORDR Then
                                com = BtOpUpdate
                            Else
                                '元を削除
                                sts = BTRV(BtOpDelete, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_TEMP1, Len(K0_ODR_TEMP1), 0)
                                
                                Call UniCode_Conv(ODR_ORDER_REC.SHIMUKE, GW_SIMUKE)
                                Call UniCode_Conv(ODR_ORDER_REC.JGYOBU, GW_JIGYOBU)
                                Call UniCode_Conv(ODR_ORDER_REC.NAIGAI, GW_NAIGAI)
                                Call UniCode_Conv(ODR_ORDER_REC.INS_NO, DIS_KEY)
                                Call UniCode_Conv(ODR_ORDER_REC.ORDER_NO, DIS_ORDR_NO)
                                Call UniCode_Conv(ODR_ORDER_REC.BUN_NO, DIS_BUNNO)
                                Call UniCode_Conv(ODR_ORDER_REC.HIN_GAI, GW_HINGAI)
                                '新を追加
                                Do
                                    Call UniCode_Conv(ODR_ORDER_REC.INS_NO, DIS_KEY)
                                    Do
                                        sts = BTRV(BtOpGetEqual, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_ORDER, Len(K0_ODR_ORDER), 0)
                                        Select Case sts
                                            Case BtNoErr
                                                Exit Do
                                            Case BtErrKeyNotFound, BtErrEOF
                                                sts = BtErrKeyNotFound
                                                Exit Do
                                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                                Sleep (500)
                                            Case Else
                                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "ODR_ORDER")
                                                GoTo Err_Exit
                                        End Select
                                    Loop
                                    If sts = BtErrKeyNotFound Then Exit Do
                                    W_DIS_NO = CLng(DIS_KEY) + 1
                                    If W_DIS_NO > 9999 Then W_No = 1
                                    DIS_KEY = Format(W_DIS_NO, "0000")
                                Loop
                                com = BtOpInsert
                                Call ODR_ORDER_CLR
                                Call UniCode_Conv(ODR_ORDER_REC.SHIMUKE, GW_SIMUKE)
                                Call UniCode_Conv(ODR_ORDER_REC.JGYOBU, GW_JIGYOBU)
                                Call UniCode_Conv(ODR_ORDER_REC.NAIGAI, GW_NAIGAI)
                                                                
                                Call UniCode_Conv(ODR_ORDER_REC.INS_NO, DIS_KEY)
                                
                                Call UniCode_Conv(ODR_ORDER_REC.ORDER_NO, DIS_ORDR_NO)
                                Call UniCode_Conv(ODR_ORDER_REC.BUN_NO, DIS_BUNNO)
                                Call UniCode_Conv(ODR_ORDER_REC.HIN_GAI, GW_HINGAI)
                                
                            End If
                        End If
                        
                        Exit Do
                    Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                        com = BtOpInsert
                        Call ODR_ORDER_CLR
                        Call UniCode_Conv(ODR_ORDER_REC.SHIMUKE, GW_SIMUKE)
                        Call UniCode_Conv(ODR_ORDER_REC.JGYOBU, GW_JIGYOBU)
                        Call UniCode_Conv(ODR_ORDER_REC.NAIGAI, GW_NAIGAI)
                        
                        
                        Call UniCode_Conv(ODR_ORDER_REC.INS_NO, DIS_KEY)
                        
                        Call UniCode_Conv(ODR_ORDER_REC.ORDER_NO, DIS_ORDR_NO)
                        Call UniCode_Conv(ODR_ORDER_REC.BUN_NO, DIS_BUNNO)
                        Call UniCode_Conv(ODR_ORDER_REC.HIN_GAI, GW_HINGAI)
                        
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
            
            If com <> BtOpDelete Then
                
                Call UniCode_Conv(ODR_ORDER_REC.HIN_GAI, GW_HINGAI)
                Call UniCode_Conv(ODR_ORDER_REC.ORDER_NO, DIS_ORDR_NO)
                
                '2008.12.16追加
                '                   追加時、Head情報を追加
                If Trim(StrConv(ODR_ORDER_REC.USE_YM_MOTO, vbUnicode)) = "" Then
                    W_Use_YM = Text1(ptxUSE_YY)
                    Call UniCode_Conv(ODR_ORDER_REC.USE_YM_MOTO, Left(W_Use_YM, 4) & Right(W_Use_YM, 2))
                End If
                
                
                W_Use_YM = ORDR_GRID(Cor_Row, Col_USE_YM)
                Call UniCode_Conv(ODR_ORDER_REC.USE_YM, Left(W_Use_YM, 4) & Right(W_Use_YM, 2))
                If W_Use_YM <> Text1(ptxUSE_YY) Then
                    Data_Out_Need = 1
                End If
                
                If CLng(ORDR_GRID(Cor_Row, Col_ORDR_QTY)) >= 0 Then
                                
                    Call UniCode_Conv(ODR_ORDER_REC.ODR_QTY, Format(CLng(ORDR_GRID(Cor_Row, Col_ORDR_QTY)), "00000"))
                Else
                    Call UniCode_Conv(ODR_ORDER_REC.ODR_QTY, Format(CLng(ORDR_GRID(Cor_Row, Col_ORDR_QTY)), "0000"))
                End If
                
                If CDbl(ORDR_GRID(Cor_Row, Col_ORDR_QTY)) < 0 Then
                    W_STR = ""
                End If
                '2008.12.17 数量の編集を変更！
                W_STR = CStr(ORDR_GRID(Cor_Row, Col_ORDR_QTY))
                Call UniCode_Conv(ODR_ORDER_REC.ODR_QTY, W_STR)
                
                
                
                Call UniCode_Conv(ODR_ORDER_REC.CYUMON_DT, Format(Trim(ORDR_GRID(Cor_Row, Col_NOUKI)), "yyyymmdd"))
                
                Call UniCode_Conv(ODR_ORDER_REC.KAITO_DT, Format(Trim(ORDR_GRID(Cor_Row, Col_KAITO)), "yyyymmdd"))
                
                If Trim(StrConv(ODR_ORDER_REC.KAITO_DT, vbUnicode)) = "" Then
                    If CLng(ORDR_GRID(Cor_Row, Col_ORDR_QTY)) < 0 Then
                        Call UniCode_Conv(ODR_ORDER_REC.KAITO_DT, Format(Now, "yyyymmdd"))
                    End If
                End If
                
'                完了日付は更新しない    2016.06.28
'                Call UniCode_Conv(ODR_ORDER_REC.FIN_DT, Format(Trim(ORDR_GRID(Cor_Row, Col_FIN_DT)), "yyyymmdd"))
                
                Call UniCode_Conv(ODR_ORDER_REC.KUMI_OK_DT, Format(Trim(ORDR_GRID(Cor_Row, Col_OK_DT)), "yyyymmdd"))
                
                Call UniCode_Conv(ODR_ORDER_REC.HIN_GAI, GW_HINGAI)
                Call UniCode_Conv(ODR_ORDER_REC.ORDER_NO, DIS_ORDR_NO)
                
                Call UniCode_Conv(ODR_ORDER_REC.UPD_TANTO, Text1(ptxTANTO_CD))
                
                Call UniCode_Conv(ODR_ORDER_REC.UPD_DT, Format(Date, "yyyymmdd"))
                Call UniCode_Conv(ODR_ORDER_REC.UPD_TM, Format(Time, "hhmmss"))
                Call UniCode_Conv(ODR_ORDER_REC.UPD_PG, Trim(App.EXEName))
                
                '2009/03/12
                If com = BtOpInsert Then
                    Call UniCode_Conv(ODR_ORDER_REC.INS_DT, Format(Date, "yyyymmdd"))
                    Call UniCode_Conv(ODR_ORDER_REC.INS_TM, Format(Time, "hhmmss"))
                End If
            End If
            
            Do
                sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_TEMP1, Len(K0_ODR_TEMP1), 0)
                Select Case sts
                    Case BtNoErr
                        
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                        Sleep (500)
                    Case Else
                        If sts <> BtErrDuplicates Then
                            Call File_Error(sts, com, "ODR_ORDER")
                            GoTo Err_Exit
                        End If
                        W_DIS_NO = CLng(StrConv(ODR_ORDER_REC.INS_NO, vbUnicode)) + 1
                        If W_DIS_NO > 9999 Then W_No = 1
                        DIS_KEY = Format(W_DIS_NO, "0000")
                        Call UniCode_Conv(ODR_ORDER_REC.INS_NO, DIS_KEY)
                        Call UniCode_Conv(K0_ODR_ORDER.INS_NO, DIS_KEY)
                End Select
            Loop
            ORDR_GRID(Cor_Row, Col_KEY) = DIS_KEY           'データＫｅｙ情報   登録順
            ORDR_GRID(Cor_Row, Col_KEY_ORDR) = DIS_ORDR_NO  'データＫｅｙ情報   注文№
            ORDR_GRID(Cor_Row, Col_KEY_OYA) = GW_HINGAI     'データＫｅｙ情報   親品番
            
            
            
            
            
            
            If REQ_UPDT(com) Then
                MsgBox "所要量Ｆ更新エラー！", vbExclamation
                GoTo Err_Exit
            End If
            
        End If
        
        
'>>>>>>>>>>　2016.11.25
        ORDR_GRID(Cor_Row, Col_SV_NOUKI) = ORDR_GRID(Cor_Row, Col_NOUKI)
        ORDR_GRID(Cor_Row, Col_SV_KAITO) = ORDR_GRID(Cor_Row, Col_KAITO)
'>>>>>>>>>>　2016.11.25

        
    Next Cor_Row
    
    TDBGrid1.ReBind
    TDBGrid1.Update
            
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "親部品　注文情報　更新終了。　＜Rec_UPDT＞", Me.hwnd, 0)
    
    
    Rec_UPDT = False
    
Err_Exit:
    If In_Lock = True Then
        Call Input_UnLock
    End If
End Function
Private Function REQ_UPDT(SYORI As Integer)
                            '       所要量Ｆの更新！

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer
    
Dim W_QTY       As Double
Dim W_STR       As String


    REQ_UPDT = True
    
    
    Key_SIMUKE = Trim(StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode))        '仕向け先
    Key_JIGYOBU = Trim(StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode))        '事業部
    Key_NAIGAI = Trim(StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode))         '国内外
    Key_HinGai = Trim(StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode))        '親品番
    Key_ORDER_NO = Trim(StrConv(ODR_ORDER_REC.ORDER_NO, vbUnicode))     '親品番　注文№
    Key_INS_NO = Trim(StrConv(ODR_ORDER_REC.INS_NO, vbUnicode))         '登録順
    Key_BUN_NO = Trim(StrConv(ODR_ORDER_REC.BUN_NO, vbUnicode))         '分納回数
    Key_USE_YM = Trim(StrConv(ODR_ORDER_REC.USE_YM, vbUnicode))         '使用月
        
        
    
    Select Case SYORI
        Case BtOpInsert
                
                
        Case BtOpUpdate, BtOpDelete
            Call UniCode_Conv(K0_ODR_REQ.SHIMUKE, Key_SIMUKE)
            Call UniCode_Conv(K0_ODR_REQ.JGYOBU, Key_JIGYOBU)
            Call UniCode_Conv(K0_ODR_REQ.NAIGAI, Key_NAIGAI)
            Call UniCode_Conv(K0_ODR_REQ.HIN_GAI, Key_HinGai)
            Call UniCode_Conv(K0_ODR_REQ.ORDER_NO, Key_ORDER_NO)
            Call UniCode_Conv(K0_ODR_REQ.INS_NO, Key_INS_NO)
            Call UniCode_Conv(K0_ODR_REQ.BUN_NO, Key_BUN_NO)
            Call UniCode_Conv(K0_ODR_REQ.KO_HIN_GAI, "")
            
            
            Call UniCode_Conv(K0_ODR_REQ.HIN_GAI, GW_MOTO_OYA)
            Call UniCode_Conv(K0_ODR_REQ.ORDER_NO, GW_MOTO_ORDR)
            
            
'2019.01.08            com = BtOpGetGreaterEqual + BtSNoWait
            com = BtOpGetGreaterEqual
            yn = 0
            Do
                Do
                    sts = BTRV(com, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                            
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                            Sleep (500)
                            yn = yn + 1
                            If yn > 500 Then
                                yn = MsgBox("他で使用中です！<所要量Ｆ>" & Chr(13) & Chr(10) & _
                                            "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                                If yn = vbNo Then Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, com, "ODR_REQUIRE")
                            Exit Function
                    End Select
                Loop
                If sts <> BtNoErr Then Exit Do
                If Trim(StrConv(ODR_REQ_R.SHIMUKE, vbUnicode)) <> Trim(Key_SIMUKE) Then Exit Do
                If Trim(StrConv(ODR_REQ_R.JGYOBU, vbUnicode)) <> Trim(Key_JIGYOBU) Then Exit Do
                If Trim(StrConv(ODR_REQ_R.NAIGAI, vbUnicode)) <> Trim(Key_NAIGAI) Then Exit Do
                
                'If Trim(StrConv(ODR_REQ_R.HIN_GAI, vbUnicode)) <> Trim(Key_HinGai) Then Exit Do
                'If Trim(StrConv(ODR_REQ_R.ORDER_NO, vbUnicode)) <> Trim(Key_ORDER_NO) Then Exit Do
                
                If Trim(StrConv(ODR_REQ_R.HIN_GAI, vbUnicode)) <> Trim(GW_MOTO_OYA) Then Exit Do
                If Trim(StrConv(ODR_REQ_R.ORDER_NO, vbUnicode)) <> Trim(GW_MOTO_ORDR) Then Exit Do
                
                
                If Trim(StrConv(ODR_REQ_R.INS_NO, vbUnicode)) <> Trim(Key_INS_NO) Then Exit Do
                If Trim(StrConv(ODR_REQ_R.BUN_NO, vbUnicode)) <> Trim(Key_BUN_NO) Then Exit Do
                
                If SYORI = BtOpDelete Then
                
                Else
'2008.05.02                    If Trim(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode)) = "" Then
                        W_QTY = CDbl(StrConv(ODR_REQ_R.REQ_QTY, vbUnicode))
'2008.05.02                    Else
'2008.05.02                        W_QTY = 0
'2008.05.02                    End If
                    
                    
                    If W_QTY >= 0 Then
                    
                        Call UniCode_Conv(ODR_REQ_R.ODR_QTY, Format(W_QTY, "00000"))
                
                    Else
                        Call UniCode_Conv(ODR_REQ_R.ODR_QTY, Format(W_QTY, "0000"))
                    End If
                    
                    Call UniCode_Conv(ODR_REQ_R.ODR_QTY, CStr(W_QTY))
                    
                    Call UniCode_Conv(ODR_REQ_R.HIN_GAI, Key_HinGai)
                    Call UniCode_Conv(ODR_REQ_R.ORDER_NO, Key_ORDER_NO)

                End If
                
                Do
                    sts = BTRV(SYORI, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                            Sleep (500)
                        Case Else
                            Call File_Error(sts, SYORI, "ODR_REQUIRE")
                            Exit Function
                    End Select
                Loop
                
                
                com = BtOpGetNext + BtSNoWait
            Loop
            If sts = BtNoErr Then
                sts = BTRV(BtOpUnlock, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
            End If
            
            
        Case Else
            
    End Select

    REQ_UPDT = False

End Function
Private Function Data_Out(OUT_Path As String, W_CNT As Long) As Integer

                            '   使用月が変更された親に対応する発注全件！


Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_YYMM      As String
Dim X_i         As Integer

Dim W_HINGAI    As String
Dim W_STR       As String
Dim W_Moto       As String
Dim W_NEW      As String
Dim F_No        As Integer
Dim W_Sw        As Integer

    Data_Out = True
    
    Call Input_Lock
    W_CNT = 0
    F_No = FreeFile
    Open Trim(OUT_Path) For Output As #F_No
    '見出し出力
    '
    Write #F_No, "親品番";
    Write #F_No, "子品番";
    Write #F_No, "注文№";
    Write #F_No, "数量";
    Write #F_No, "仕入先";
    Write #F_No, "仕入先名";
    Write #F_No, "納期";
    Write #F_No, "納期変更";
    Write #F_No, "変更前使用月";
    Write #F_No, "変更後使用月";
    Write #F_No,
    
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "子部品　発注情報　出力中！　＜Data_Out＞", Me.hwnd, 0)
    DoEvents
    
    
    W_YYMM = Left(GW_SHIMEBI, 4) & Mid(GW_SHIMEBI, 5, 2)
    
    
    Call UniCode_Conv(K1_ODR_ORDER.SHIMUKE, GW_SIMUKE)
    Call UniCode_Conv(K1_ODR_ORDER.JGYOBU, GW_JIGYOBU)
    Call UniCode_Conv(K1_ODR_ORDER.NAIGAI, GW_NAIGAI)
    Call UniCode_Conv(K1_ODR_ORDER.USE_YM, W_YYMM)
    
    Call UniCode_Conv(K1_ODR_ORDER.HIN_GAI, "")
    Call UniCode_Conv(K1_ODR_ORDER.ORDER_NO, "")
    Call UniCode_Conv(K1_ODR_ORDER.INS_NO, "")
    Call UniCode_Conv(K1_ODR_ORDER.BUN_NO, "")
    
    com = BtOpGetGreaterEqual
    Do
        
        sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K1_ODR_ORDER, Len(K1_ODR_ORDER), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                
                Exit Do
            Case Else
                Call File_Error(sts, com, "ODR_ORDER")
                Exit Do
        End Select
        
        
        If Trim(StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode)) <> Trim(GW_SIMUKE) Or _
            Trim(StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU) Or _
            Trim(StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI) Then
            Exit Do
        End If
        
        
        W_NEW = StrConv(ODR_ORDER_REC.USE_YM, vbUnicode)
        W_Moto = StrConv(ODR_ORDER_REC.USE_YM_MOTO, vbUnicode)
                                    
        If W_NEW <> W_Moto Then
        
        
            W_HINGAI = Trim(StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode))
            DIS_ORDR_NO = Trim(StrConv(ODR_ORDER_REC.ORDER_NO, vbUnicode))
            DIS_KEY = Trim(StrConv(ODR_ORDER_REC.INS_NO, vbUnicode))
            DIS_BUNNO = Trim(StrConv(ODR_ORDER_REC.BUN_NO, vbUnicode))
            
            If Trim(DIS_BUNNO) <> "" Then
                DIS_BUNNO = Format(CInt(DIS_BUNNO), "000")
            End If
            
            Call UniCode_Conv(K0_ODR_REQ.SHIMUKE, GW_SIMUKE)
            Call UniCode_Conv(K0_ODR_REQ.JGYOBU, GW_JIGYOBU)
            Call UniCode_Conv(K0_ODR_REQ.NAIGAI, GW_NAIGAI)
            Call UniCode_Conv(K0_ODR_REQ.HIN_GAI, W_HINGAI)
            Call UniCode_Conv(K0_ODR_REQ.INS_NO, DIS_KEY)
            Call UniCode_Conv(K0_ODR_REQ.ORDER_NO, DIS_ORDR_NO)
            Call UniCode_Conv(K0_ODR_REQ.BUN_NO, DIS_BUNNO)
            Call UniCode_Conv(K0_ODR_REQ.KO_HIN_GAI, "")
            com = BtOpGetGreaterEqual
            
            Do
                
                sts = BTRV(com, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                        
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "ODR_REQUIRE")
                        GoTo Err_Exit
                End Select
                If sts <> BtNoErr Then Exit Do
                If Trim(StrConv(ODR_REQ_R.SHIMUKE, vbUnicode)) <> Trim(GW_SIMUKE) Then Exit Do
                If Trim(StrConv(ODR_REQ_R.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU) Then Exit Do
                If Trim(StrConv(ODR_REQ_R.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI) Then Exit Do
                If Trim(StrConv(ODR_REQ_R.HIN_GAI, vbUnicode)) <> Trim(W_HINGAI) Then Exit Do
                If Trim(StrConv(ODR_REQ_R.INS_NO, vbUnicode)) <> Trim(DIS_KEY) Then Exit Do
                If Trim(StrConv(ODR_REQ_R.ORDER_NO, vbUnicode)) <> Trim(DIS_ORDR_NO) Then Exit Do
                If Trim(StrConv(ODR_REQ_R.BUN_NO, vbUnicode)) <> Trim(DIS_BUNNO) Then Exit Do
                
                If Trim(StrConv(ODR_REQ_R.KO_HIN_GAI, vbUnicode)) = "D061" Then
                    X_i = 0
                End If
                
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                '       対象の発注情報の検索＆出力
                '
                Call UniCode_Conv(K1_P_SHORDER.JGYOBU, StrConv(ODR_REQ_R.KO_JGYOBU, vbUnicode))
                Call UniCode_Conv(K1_P_SHORDER.NAIGAI, StrConv(ODR_REQ_R.KO_NAIGAI, vbUnicode))
                Call UniCode_Conv(K1_P_SHORDER.HIN_GAI, StrConv(ODR_REQ_R.KO_HIN_GAI, vbUnicode))
                Call UniCode_Conv(K1_P_SHORDER.ORDER_NO, "")
                Call UniCode_Conv(K1_P_SHORDER.ORDER_DT, "")
                
                com = BtOpGetGreaterEqual
                Do
                    sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K1_P_SHORDER, Len(K1_P_SHORDER), 1)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                            
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "P_SHORDER")
                            GoTo Err_Exit
                    End Select
                    If sts <> BtNoErr Then Exit Do
                    If Trim(StrConv(P_SHORDER_REC.JGYOBU, vbUnicode)) <> Trim(StrConv(ODR_REQ_R.KO_JGYOBU, vbUnicode)) Then Exit Do
                    If Trim(StrConv(P_SHORDER_REC.NAIGAI, vbUnicode)) <> Trim(StrConv(ODR_REQ_R.KO_NAIGAI, vbUnicode)) Then Exit Do
                    If Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)) <> Trim(StrConv(ODR_REQ_R.KO_HIN_GAI, vbUnicode)) Then Exit Do
                    
                    If Trim(StrConv(P_SHORDER_REC.USE_YM, vbUnicode)) = W_Moto Then
                        
                        W_Sw = True
                           
                        If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = "1" Then W_Sw = False   'キャンセル？
                        '2008.05.01
                        If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = P_KAN_ON Then W_Sw = False      '完了？
                        
                        '2008.11.26　使用月が元と同一
                        'If StrConv(P_SHORDER_REC.USE_YM, vbUnicode) <> W_Moto Then W_Sw = False
                        
                        '2008.12.16     ↑この１行はデバッグ用！　⇒　不要！！　(*_*)
                        
                        
                        
                        
                        If W_Sw Then
    
                            '親品番
                            Write #F_No, Trim(StrConv(ODR_REQ_R.HIN_GAI, vbUnicode));
                            '子品番
                            Write #F_No, Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode));
                            
                            '注文№
                            Write #F_No, Trim(StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode));
                            
                            '数量
                            Write #F_No, Trim(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode));
                            '仕入先
                            Write #F_No, Trim(StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode));
                            
                            '仕入先名
                            Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode))
                            Do
                                sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                                        Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_NAME, "")
                                        Exit Do
                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "P_SHORDER")
                                        Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_NAME, "")
                                        Exit Do
                                End Select
                            Loop
                            Write #F_No, Trim(StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode));
                            
                            
                            
                            '納期
                            If Trim(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode)) <> "" Then
                                W_STR = Left(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 4)
                                W_STR = W_STR & "/"
                                W_STR = W_STR & Mid(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 5, 2)
                                W_STR = W_STR & "/"
                                W_STR = W_STR & Right(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 2)
                                
                            Else
                                W_STR = Left(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 4)
                                W_STR = W_STR & "/"
                                W_STR = W_STR & Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 5, 2)
                                W_STR = W_STR & "/"
                                W_STR = W_STR & Right(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 2)
                            End If
                            Write #F_No, W_STR;
                            
                            
                            '納期変更：手書き欄！？
                            W_STR = ""
                            Write #F_No, W_STR;
                            
                            
                            
                            '変更前使用月
                            
                            W_STR = Left(StrConv(P_SHORDER_REC.USE_YM, vbUnicode), 4)
                           ' W_STR = W_STR & "/"
                            W_STR = W_STR & Mid(StrConv(P_SHORDER_REC.USE_YM, vbUnicode), 5, 2)
                            Write #F_No, W_STR;     '"_" & W_Str;
                            
                            '変更後使用月
                            W_STR = Left(W_NEW, 4)
                            'W_STR = W_STR & "/"
                            W_STR = W_STR & Mid(W_NEW, 5, 2)
                            Write #F_No, W_STR;         '"_" & W_Str;
                            
                            
                            '改行
                            Write #F_No,
                            
                            W_CNT = W_CNT + 1
                        End If
                    End If
                    
                    com = BtOpGetNext
                Loop
                
                com = BtOpGetNext
            Loop
            
        End If
        com = BtOpGetNext
    Loop
    
    
    Close #F_No
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "子部品　発注情報　出力終了。　＜Data_Out＞", Me.hwnd, 0)
    
    
    Data_Out = False
    
Err_Exit:
    Call Input_UnLock
End Function

Private Function Data_Out2(OUT_Path As String, W_CNT As Long) As Integer

                            '   表示中の内容を全件出力！

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_YYMM      As String
Dim X_i         As Integer

Dim W_HINGAI    As String
Dim W_STR       As String
Dim W_Moto       As String
Dim W_NEW      As String
Dim F_No        As Integer
Dim W_Sw        As Integer

    Data_Out2 = True
    
    Call Input_Lock
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "画面情報　出力中！　＜Data_Out2＞", Me.hwnd, 0)
    DoEvents
    
    
    W_CNT = 0
    F_No = FreeFile
    Open Trim(OUT_Path) For Output As #F_No
    '見出し出力
    '
    Write #F_No, "№";
    Write #F_No, "親部品注文№";
    Write #F_No, "分納";
    Write #F_No, "親品番";
    Write #F_No, "部品名";
    Write #F_No, "数量";
    Write #F_No, "注文納期";
    Write #F_No, "組立可能日";
    Write #F_No, "回答納期";
    Write #F_No, "使用月";
    Write #F_No, "完了日付";
    
    If Grid_Cor_M = True Then
        Write #F_No, "画面未更新！";
    End If
    
    Write #F_No,
    
    X_i = row
    
    For X_i = Min_Row To ORDR_GRID.UpperBound(1)
    
        W_CNT = W_CNT + 1
                    'Seq-No
        Write #F_No, CStr(W_CNT);
        
                    '親部品注文№
        Write #F_No, ORDR_GRID(X_i, Col_ORDR_NO);
        
                    '分納
        Write #F_No, ORDR_GRID(X_i, Col_BUNNO);
        
                    '親品番
        Write #F_No, ORDR_GRID(X_i, Col_OYA_ITEM);
        
                    '部品名
        Write #F_No, ORDR_GRID(X_i, Col_OYA_NM);
        
                    '数量
        Write #F_No, ORDR_GRID(X_i, Col_ORDR_QTY);
                
                    '注文納期
        Write #F_No, ORDR_GRID(X_i, Col_NOUKI);
        
                    '組立可能日
        Write #F_No, ORDR_GRID(X_i, Col_OK_DT);
        
                    '回答納期
        Write #F_No, ORDR_GRID(X_i, Col_KAITO);
        
                    '使用月
        'Write #F_No, ORDR_GRID(X_i, Col_USE_YM);    '"_" & ORDR_GRID(X_i, Col_USE_YM);
        W_STR = Left(ORDR_GRID(X_i, Col_USE_YM), 4) & Right(ORDR_GRID(X_i, Col_USE_YM), 2)
        Write #F_No, W_STR;
        
        
                    '完了日付
        Write #F_No, ORDR_GRID(X_i, Col_FIN_DT);
    
        Write #F_No,
        
        
    Next X_i
    
    
    Close #F_No
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "画面情報　出力終了。　＜Data_Out2＞", Me.hwnd, 0)
    
    
    Data_Out2 = False
    
Err_Exit:
    Call Input_UnLock
End Function

Private Function Require_Set() As Integer
'
'       子部品　展開＆組立可能日セット
'
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim X_i         As Integer
Dim W_After     As String

Dim W_HINGAI    As String
Dim W_STR       As String
Dim W_Date      As String

Dim W_No        As String
Dim W_Zan       As Double

Dim FullPath        As String
Dim c               As String * 128



    Require_Set = True
    
    Call Input_Lock
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    '   グリッド　未チェック＆未更新　→　チェック＆更新
    If Grid_Cor_M <> False Then
        hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "親部品　注文情報　展開中！＜Require_SET＞ Step-1(エラーチェック)", Me.hwnd, 0)
               
        'Set ORDR_GRID = TDBGrid1.Array
        'TDBGrid1.Update
        
        
        DoEvents
        For Cor_Row = Min_Row To ORDR_GRID.UpperBound(1)
                
            For X_i = Col_DEL To Col_FIN_DT%
                        
                W_After = ORDR_GRID(Cor_Row, X_i)
                
                
                If Grid_Err_Chk(X_i, W_After) Then
                    TDBGrid1.ReBind
                    TDBGrid1.Update
                            'TDBGrid1.MoveFirst
                    TDBGrid1.ScrollBars = dbgAutomatic
    
                    TDBGrid1.SetFocus
                    GoTo Err_Exit
                End If
                   
                
            Next X_i
                    
        Next Cor_Row
        
        TDBGrid1.ReBind
        TDBGrid1.Update
        'TDBGrid1.MoveFirst
        TDBGrid1.ScrollBars = dbgAutomatic
                    
        hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "親部品　注文情報　展開中！＜Require_SET＞ Step-2(レコード更新)", Me.hwnd, 0)
        DoEvents
        
                '更新処理
        If Rec_UPDT(False) Then
            MsgBox "更新失敗しました。", vbExclamation
            GoTo Err_Exit
        End If
               
        Grid_Cor_M = False
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "親部品　注文情報　展開中！＜Require_SET＞ Step-3(中間ファイル　削除＆再作成)", Me.hwnd, 0)
    DoEvents
    
                            '所要量展開Ｆ 占有Open → Close → KILL → 占有Open
                            
    If ODR_TEMP1_Open(BtOpenExec) Then
        MsgBox "処理を中断します。", vbExclamation
        GoTo Err_Exit
    End If
    sts = BTRV(BtOpClose, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K0_ODR_TEMP1, Len(K0_ODR_TEMP1), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_TEMP1")
        End If
    End If
    
    If ODR_TEMP1_KILL Then
        GoTo Err_Exit
    End If
    
    If ODR_TEMP1_Open(BtOpenExec) Then
        MsgBox "処理を中断します。", vbExclamation
        GoTo Err_Exit
    End If
    
                '所要量展開Ｆ 占有Open → Close → KILL → 占有Open     Part-2
    
    If ODR_TEMP2_Open(BtOpenExec) Then
        MsgBox "処理を中断します。", vbExclamation
        GoTo Err_Exit
    End If
    sts = BTRV(BtOpClose, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_TEMP2")
        End If
    End If
    
    If ODR_TEMP2_KILL Then
        GoTo Err_Exit
    End If
    
    If ODR_TEMP2_Open(BtOpenExec) Then
        MsgBox "処理を中断します。", vbExclamation
        GoTo Err_Exit
    End If
    
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "親部品　注文情報　展開中！＜Require_SET＞ Step-4(子部品　展開処理)", Me.hwnd, 0)
    DoEvents

    Call ODR_TEMP1_CLR
    
    Call UniCode_Conv(K1_ODR_ORDER.SHIMUKE, GW_SIMUKE)
    Call UniCode_Conv(K1_ODR_ORDER.JGYOBU, GW_JIGYOBU)
    Call UniCode_Conv(K1_ODR_ORDER.NAIGAI, GW_NAIGAI)
    Call UniCode_Conv(K1_ODR_ORDER.USE_YM, GW_TOUGETU)
    Call UniCode_Conv(K1_ODR_ORDER.HIN_GAI, "")
    Call UniCode_Conv(K1_ODR_ORDER.INS_NO, "")
    Call UniCode_Conv(K1_ODR_ORDER.ORDER_NO, "")
    Call UniCode_Conv(K1_ODR_ORDER.BUN_NO, "")
    com = BtOpGetGreaterEqual
    Do
        X_i = 0
        sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K1_ODR_ORDER, Len(K1_ODR_ORDER), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                
            Case Else
                Call File_Error(sts, com, "ODR_ORDER")
                GoTo Err_Exit
        End Select
        If sts <> BtNoErr Then Exit Do
        
        If Trim(StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode)) <> Trim(GW_SIMUKE) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI) Then Exit Do
        
        
        '分納の親レコードは展開対象外！
        
        If CInt(StrConv(ODR_ORDER_REC.BUN_KB, vbUnicode)) = 0 Then
            
            Key_SIMUKE = GW_SIMUKE
            Key_JIGYOBU = GW_JIGYOBU
            Key_NAIGAI = GW_NAIGAI
            Key_USE_YM = Trim(StrConv(ODR_ORDER_REC.USE_YM, vbUnicode))
            Key_HinGai = Trim(StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode))
            Key_INS_NO = Trim(StrConv(ODR_ORDER_REC.INS_NO, vbUnicode))
            Key_ORDER_NO = Trim(StrConv(ODR_ORDER_REC.ORDER_NO, vbUnicode))
            Key_BUN_NO = Trim(StrConv(ODR_ORDER_REC.BUN_NO, vbUnicode))
                
            W_HINGAI = Trim(StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode))
                
            '未完了　→　子部品展開
            'If Trim(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode)) = "" Then
                
                '2008/03/10　完了も展開！に変更。
                If OUT_TP1(W_HINGAI) Then
                    MsgBox "展開処理エラー！", vbExclamation
                    GoTo Err_Exit
                End If
            
            'End If
            
            '注文数量＜０：減算使用数
            If CDbl(Trim(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode))) < 0 Then
                If SET_O_MAINA(W_HINGAI) Then
                    MsgBox "注文数＜０の在庫計算エラー", vbExclamation
                    GoTo Err_Exit
                End If
            End If
        End If
        com = BtOpGetNext
    Loop
    
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                        '仮想在庫数の展開／集計
                        '注文数＜０は、一時的に構成子部品の在庫数を増加
                                    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "親部品　注文情報　展開中！＜Require_SET＞ Step-5(仮想在庫　集計処理)", Me.hwnd, 0)
    DoEvents
    
    
    'このSET_ALLで、在庫数、発注数の情報を作成している。
    If SET_ALL Then
        MsgBox "在庫情報設定エラー！", vbExclamation
        GoTo Err_Exit
    End If
    
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                                    '中間展開Ｆ Close → 共用Open
    sts = BTRV(BtOpClose, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K0_ODR_TEMP1, Len(K0_ODR_TEMP1), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_TEMP1")
        End If
    End If
    If ODR_TEMP1_Open(BtOpenNomal) Then
        MsgBox "処理を中断します。", vbExclamation
        GoTo Err_Exit
    End If
    
    sts = BTRV(BtOpClose, ODR_TP2_POS, ODR_TP2_R, Len(ODR_TP2_R), K0_ODR_TEMP2, Len(K0_ODR_TEMP2), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_TEMP2")
        End If
    End If
    If ODR_TEMP2_Open(BtOpenNomal) Then
        MsgBox "処理を中断します。", vbExclamation
        GoTo Err_Exit
    End If
    
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "親部品　注文情報　展開中！＜Require_SET＞ Step-6(在庫引当)", Me.hwnd, 0)
    DoEvents
    
    '       在庫、発注情報と所要量で、引当残を計算／設定。
    If ZAN_CALC() Then
        MsgBox "在庫引当計算エラー！", vbExclamation
        GoTo Err_Exit
    End If
    
    
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "親部品　注文情報　展開中！＜Require_SET＞ Step-7(引当可能日設定)", Me.hwnd, 0)
    DoEvents
    
    
    
    '               2008/09/11          全親部品を対象に変更する事！！！    (*_*)
    
    Call UniCode_Conv(K1_ODR_ORDER.SHIMUKE, GW_SIMUKE)
    Call UniCode_Conv(K1_ODR_ORDER.JGYOBU, GW_JIGYOBU)
    Call UniCode_Conv(K1_ODR_ORDER.NAIGAI, GW_NAIGAI)
    Call UniCode_Conv(K1_ODR_ORDER.USE_YM, GW_TOUGETU)
    Call UniCode_Conv(K1_ODR_ORDER.HIN_GAI, "")
    Call UniCode_Conv(K1_ODR_ORDER.INS_NO, "")
    Call UniCode_Conv(K1_ODR_ORDER.ORDER_NO, "")
    Call UniCode_Conv(K1_ODR_ORDER.BUN_NO, "")
    com = BtOpGetGreaterEqual
    Do
        sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K1_ODR_ORDER, Len(K1_ODR_ORDER), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                
            Case Else
                Call File_Error(sts, com, "ODR_ORDER")
                GoTo Err_Exit
        End Select
        If sts <> BtNoErr Then Exit Do
        
        If Trim(StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode)) <> Trim(GW_SIMUKE) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI) Then Exit Do
        
        
        '分納の親レコードは展開対象外！
        
        If CInt(StrConv(ODR_ORDER_REC.BUN_KB, vbUnicode)) = 0 Then
            
            Key_SIMUKE = GW_SIMUKE
            Key_JIGYOBU = GW_JIGYOBU
            Key_NAIGAI = GW_NAIGAI
            Key_USE_YM = Trim(StrConv(ODR_ORDER_REC.USE_YM, vbUnicode))
            Key_HinGai = Trim(StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode))
            Key_INS_NO = Trim(StrConv(ODR_ORDER_REC.INS_NO, vbUnicode))
            Key_ORDER_NO = Trim(StrConv(ODR_ORDER_REC.ORDER_NO, vbUnicode))
            Key_BUN_NO = Trim(StrConv(ODR_ORDER_REC.BUN_NO, vbUnicode))
                
            W_HINGAI = Trim(StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode))
            
            If W_HINGAI = "AD-KZ039WBW" Then
                W_Date = ""
            End If
            ''2008/05/31 完成しているオーダーは、セットしない！
            If Trim(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode)) = "" Then
                
                W_Date = Trim(StrConv(ODR_ORDER_REC.KUMI_OK_DT, vbUnicode))
                
                If OK_DT_SRCH(W_Date) Then
                    MsgBox "組立可能日設定でエラー！", vbExclamation
                    GoTo Err_Exit
                End If
                
                Call UniCode_Conv(ODR_ORDER_REC.KUMI_OK_DT, W_Date)
                X_i = 0
                Do
                    sts = BTRV(BtOpUpdate, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K1_ODR_ORDER, Len(K1_ODR_ORDER), 1)
                    Select Case sts
                        Case BtNoErr
                                Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                                Sleep (500)
                                X_i = X_i + 1
                                If X_i > 1000 Then
                                    MsgBox "親品番注文Ｆ　書込みタイムアウトエラー！", vbExclamation
                                    GoTo Err_Exit
                                End If
                        Case Else
                                Call File_Error(sts, BtOpUpdate, "ODR_ORDER")
                                GoTo Err_Exit
                    End Select
                Loop
                
            End If
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            
        End If
        
        
        com = BtOpGetNext
    Loop
    
    
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "親部品　注文情報　展開中！＜Require_SET＞ Step-8(所要量Ｆ更新①)", Me.hwnd, 0)
    DoEvents
    
    '2008.12.26
    '               KILL & Createに変更！
    '
    sts = BTRV(BtOpClose, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_REQ")
        End If
    End If
    
    
                                            '所要量Ｆ フルパス取込み
    sts = GetIni("FILE", ODR_REQUIRE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_REQUIRE]読み込みエラー")
        GoTo Err_Exit
    End If
    FullPath = RTrim(c)
    
                                '所要量展開Ｆ
    If ODR_REQUIRE_Open(BtOpenExec) Then
        GoTo Err_Exit
    End If
    sts = BTRV(BtOpClose, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
    
    Kill FullPath
    
    
                                '所要量展開Ｆ
    If ODR_REQUIRE_Open(BtOpenNomal) Then
        GoTo Err_Exit
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    
    Call UniCode_Conv(K3_ODR_REQ.USE_YM, StrConv(ODR_TP1_R.USE_YM, vbUnicode))       '使用月（YYYYMM)
    Call UniCode_Conv(K3_ODR_REQ.KO_JGYOBU, "")     '子　事業部
    Call UniCode_Conv(K3_ODR_REQ.KO_NAIGAI, "")     '子　国内外
    Call UniCode_Conv(K3_ODR_REQ.KO_HIN_GAI, "")    '子品番
    Call UniCode_Conv(K3_ODR_REQ.SHIMUKE, "")       '仕向け先
    Call UniCode_Conv(K3_ODR_REQ.JGYOBU, "")        '事業部
    Call UniCode_Conv(K3_ODR_REQ.NAIGAI, "")        '国内外
    Call UniCode_Conv(K3_ODR_REQ.HIN_GAI, "")       '親品番
    Call UniCode_Conv(K3_ODR_REQ.ORDER_NO, "")      '親品番　注文№
    Call UniCode_Conv(K3_ODR_REQ.INS_NO, "")        '登録順
    Call UniCode_Conv(K3_ODR_REQ.BUN_NO, "")        '分納回数
    
    
    com = BtOpGetGreaterEqual
    Do
        Do
'2019.01.08            sts = BTRV(com + BtSNoWait, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
            sts = BTRV(com, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    yn = MsgBox("他で使用中です！<所要量Ｆ>" & Chr(13) & Chr(10) & _
                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If yn = vbNo Then GoTo Err_Exit
                Case Else
                    Call File_Error(sts, com, "ODR_REQUIRE")
                    GoTo Err_Exit
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        If StrConv(ODR_REQ_R.USE_YM, vbUnicode) <> StrConv(ODR_TP1_R.USE_YM, vbUnicode) Then Exit Do
        
        sts = BTRV(BtOpDelete, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
        
        
        com = BtOpGetNext
    Loop
    
    com = BtOpGetFirst
    Do
        sts = BTRV(com, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K0_ODR_TEMP1, Len(K0_ODR_TEMP1), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
            
            Case Else
                Call File_Error(sts, com, "ODR_TEMP1")
                GoTo Err_Exit
        End Select
        If sts <> BtNoErr Then Exit Do
        
        
        Call UniCode_Conv(K0_ODR_REQ.SHIMUKE, StrConv(ODR_TP1_R.SHIMUKE, vbUnicode))
        Call UniCode_Conv(K0_ODR_REQ.JGYOBU, StrConv(ODR_TP1_R.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ODR_REQ.NAIGAI, StrConv(ODR_TP1_R.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ODR_REQ.HIN_GAI, StrConv(ODR_TP1_R.HIN_GAI, vbUnicode))
        Call UniCode_Conv(K0_ODR_REQ.ORDER_NO, StrConv(ODR_TP1_R.ORDER_NO, vbUnicode))
        Call UniCode_Conv(K0_ODR_REQ.INS_NO, StrConv(ODR_TP1_R.INS_NO, vbUnicode))
        Call UniCode_Conv(K0_ODR_REQ.BUN_NO, StrConv(ODR_TP1_R.BUN_NO, vbUnicode))
        Call UniCode_Conv(K0_ODR_REQ.KO_HIN_GAI, StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode))
        
        
        com = BtOpUpdate
        Do
'2019.01.08            sts = BTRV(BtOpGetEqual + BtSNoWait, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
            sts = BTRV(BtOpGetEqual, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound       'レコード無し
                    com = BtOpInsert
                    Call ODR_REQUIRE_CLR
                    
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    yn = MsgBox("他で使用中です！<所要量Ｆ>" & Chr(13) & Chr(10) & _
                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If yn = vbNo Then GoTo Err_Exit
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "ODR_REQUIRE")
                    GoTo Err_Exit
            End Select
        Loop
            
            
        Call UniCode_Conv(ODR_REQ_R.SHIMUKE, StrConv(ODR_TP1_R.SHIMUKE, vbUnicode))
        Call UniCode_Conv(ODR_REQ_R.JGYOBU, StrConv(ODR_TP1_R.JGYOBU, vbUnicode))
        Call UniCode_Conv(ODR_REQ_R.NAIGAI, StrConv(ODR_TP1_R.NAIGAI, vbUnicode))
        Call UniCode_Conv(ODR_REQ_R.HIN_GAI, StrConv(ODR_TP1_R.HIN_GAI, vbUnicode))
        Call UniCode_Conv(ODR_REQ_R.ORDER_NO, StrConv(ODR_TP1_R.ORDER_NO, vbUnicode))
        Call UniCode_Conv(ODR_REQ_R.INS_NO, StrConv(ODR_TP1_R.INS_NO, vbUnicode))
        Call UniCode_Conv(ODR_REQ_R.BUN_NO, StrConv(ODR_TP1_R.BUN_NO, vbUnicode))
        Call UniCode_Conv(ODR_REQ_R.KO_HIN_GAI, StrConv(ODR_TP1_R.KO_HIN_GAI, vbUnicode))
        Call UniCode_Conv(ODR_REQ_R.KO_NAIGAI, StrConv(ODR_TP1_R.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(ODR_REQ_R.KO_JGYOBU, StrConv(ODR_TP1_R.KO_JGYOBU, vbUnicode))
        
        Call UniCode_Conv(ODR_REQ_R.CYUMON_DT, StrConv(ODR_TP1_R.CYUMON_DT, vbUnicode))
        Call UniCode_Conv(ODR_REQ_R.USE_YM, StrConv(ODR_TP1_R.USE_YM, vbUnicode))
        
        W_STR = CStr(CDbl(StrConv(ODR_TP1_R.ALL_QTY, vbUnicode)))
        Call UniCode_Conv(ODR_REQ_R.REQ_QTY, W_STR)
        
        W_STR = CStr(CDbl(StrConv(ODR_TP1_R.NED_QTY, vbUnicode)))
        Call UniCode_Conv(ODR_REQ_R.ODR_QTY, W_STR)
        
        W_STR = CStr(CDbl(StrConv(ODR_TP1_R.FUSOKU_QTY, vbUnicode)))
        Call UniCode_Conv(ODR_REQ_R.FUSOKU_QTY, W_STR)
        
        
        Call UniCode_Conv(ODR_REQ_R.OK_DT, StrConv(ODR_TP1_R.OK_DT, vbUnicode))
        
        Call UniCode_Conv(ODR_REQ_R.UPD_DT, Format(Date, "yyyymmdd"))
        Call UniCode_Conv(ODR_REQ_R.UPD_TM, Format(Time, "hhmmss"))
        
        Do
            sts = BTRV(com, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    Sleep (500)
                Case Else
                    Call File_Error(sts, com, "ODR_REQUIRE")
                    GoTo Err_Exit
            End Select
        Loop
        
        
        com = BtOpGetNext
    Loop
    
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Key_USE_YM = Left(Text1(ptxUSE_YY), 4) & Right(Text1(ptxUSE_YY), 2)
    yn = 1
    If yn <> 0 Then
        hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "親部品　注文情報　展開中！＜Require_SET＞ Step-9(所要量Ｆ更新②)", Me.hwnd, 0)
        DoEvents
        
        com = BtOpGetFirst
        
        Call UniCode_Conv(K0_ODR_REQ.SHIMUKE, GW_SIMUKE)
        Call UniCode_Conv(K0_ODR_REQ.JGYOBU, GW_JIGYOBU)
        Call UniCode_Conv(K0_ODR_REQ.NAIGAI, GW_NAIGAI)
        Call UniCode_Conv(K0_ODR_REQ.ORDER_NO, "")
        Call UniCode_Conv(K0_ODR_REQ.HIN_GAI, "")
        Call UniCode_Conv(K0_ODR_REQ.INS_NO, "")
        Call UniCode_Conv(K0_ODR_REQ.BUN_NO, "")
        
'2019.01.08        com = BtOpGetGreaterEqual + BtSNoWait
        com = BtOpGetGreaterEqual
        Do
            Do
                sts = BTRV(com, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                        
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                        yn = MsgBox("他で使用中です！<所要量Ｆ>" & Chr(13) & Chr(10) & _
                                    "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                        If yn = vbNo Then GoTo Err_Exit
                    Case Else
                        Call File_Error(sts, com, "ODR_REQUIRE")
                        GoTo Err_Exit
                End Select
            Loop
            If sts <> BtNoErr Then Exit Do
            If Trim(StrConv(ODR_REQ_R.SHIMUKE, vbUnicode)) <> GW_SIMUKE Then Exit Do
            If Trim(StrConv(ODR_REQ_R.JGYOBU, vbUnicode)) <> GW_JIGYOBU Then Exit Do
            If Trim(StrConv(ODR_REQ_R.NAIGAI, vbUnicode)) <> GW_NAIGAI Then Exit Do
            
            If StrConv(ODR_REQ_R.USE_YM, vbUnicode) = Key_USE_YM Then
    
                Call UniCode_Conv(K0_ODR_TEMP1.SHIMUKE, StrConv(ODR_REQ_R.SHIMUKE, vbUnicode))
                Call UniCode_Conv(K0_ODR_TEMP1.JGYOBU, StrConv(ODR_REQ_R.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_ODR_TEMP1.NAIGAI, StrConv(ODR_REQ_R.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_ODR_TEMP1.HIN_GAI, StrConv(ODR_REQ_R.HIN_GAI, vbUnicode))
                Call UniCode_Conv(K0_ODR_TEMP1.ORDER_NO, StrConv(ODR_REQ_R.ORDER_NO, vbUnicode))
                Call UniCode_Conv(K0_ODR_TEMP1.INS_NO, StrConv(ODR_REQ_R.INS_NO, vbUnicode))
                Call UniCode_Conv(K0_ODR_TEMP1.BUN_NO, StrConv(ODR_REQ_R.BUN_NO, vbUnicode))
                Call UniCode_Conv(K0_ODR_TEMP1.KO_HIN_GAI, StrConv(ODR_REQ_R.KO_HIN_GAI, vbUnicode))
                
                Do
                    sts = BTRV(BtOpGetEqual, ODR_TP1_POS, ODR_TP1_R, Len(ODR_TP1_R), K0_ODR_TEMP1, Len(K0_ODR_TEMP1), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrKeyNotFound       'レコード無し
                            
                            
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                            yn = MsgBox("他で使用中です！<中間所要量Ｆ>" & Chr(13) & Chr(10) & _
                                        "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                            If yn = vbNo Then GoTo Err_Exit
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "ODR_TEMP1")
                            GoTo Err_Exit
                    End Select
                Loop
            
                If sts <> BtNoErr Then
                    Do
                        sts = BTRV(BtOpDelete, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                                Sleep (500)
                            Case Else
                                Call File_Error(sts, BtOpDelete, "ODR_REQUIRE")
                                GoTo Err_Exit
                        End Select
                    Loop
                
                End If
            
            End If
            
            com = BtOpGetNext + BtSNoWait
        Loop
        
        If sts <> BtNoErr Then
            sts = BTRV(BtOpUnlock, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
        End If
        
    End If
    
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "親部品　注文情報　展開終了！　＜Require_SET＞", Me.hwnd, 0)
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    DoEvents
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "使用月単位の月初在庫　計算中！＜GESSYO_SET＞ Step-10", Me.hwnd, 0)
    DoEvents
    
     
    If GESSYO_SET Then
        MsgBox "使用月単位月初在庫　計算失敗！", vbExclamation
        GoTo Err_Exit
    End If
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "使用月単位の月初在庫　計算終了！　＜GESSYO_SET＞", Me.hwnd, 0)
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    DoEvents
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "発注検討Ｆ　計算／出力中！＜OUT_KENTO＞ Step-LAST", Me.hwnd, 0)
    DoEvents
    
     
    If OUT_KENTO Then
        sts = BTRV(BtOpClose, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
        If sts Then
            If sts <> BtErrNoOpen Then
                Call File_Error(sts, BtOpClose, "ODR_KENTO")
            End If
        End If
        MsgBox "発注検討Ｆ　計算／出力失敗！", vbExclamation
        GoTo Err_Exit
    End If

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "発注検討Ｆ　計算／出力終了！　＜OUT_KENTO＞", Me.hwnd, 0)
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    DoEvents
    Require_Set = False
    
Err_Exit:
    Call Input_UnLock
    
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

Private Sub TEMP_DEL()
Dim sts     As Integer
Dim com     As Integer
Dim yn      As Integer
Dim W_QTY   As Double

    Do
        Call UniCode_Conv(K2_ODR_ORDER.ODR_QTY, "")
        Do
            sts = BTRV(BtOpGetGreaterEqual + BtSNoWait, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K2_ODR_ORDER, Len(K2_ODR_ORDER), 2)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    yn = MsgBox("他で使用中です！<注文Ｆ>" & Chr(13) & Chr(10) & _
                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If yn = vbNo Then Exit Sub
                Case Else
                    Call File_Error(sts, BtOpGetGreaterEqual, "ODR_ORDER")
                    Exit Sub
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        
        W_QTY = CDbl(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode))
        If W_QTY > 0 Then Exit Do
        
        Do
            sts = BTRV(BtOpDelete, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K2_ODR_ORDER, Len(K2_ODR_ORDER), 2)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    Sleep (500)
                Case Else
                    Call File_Error(sts, BtOpDelete, "ODR_ORDER")
                    Exit Sub
            End Select
        Loop
        
    Loop
    
    

End Sub

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    ODR10101.MousePointer = vbHourglass

    Call Ctrl_Lock(ODR10101)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(ODR10101)


    ODR10101.MousePointer = vbDefault

End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    Select Case Index
        Case pcmbSM                 '仕向け先
            GW_SIMUKE = Left(Right(Combo1(pcmbSM).Text, 4), 2)
            'GW_JIGYOBU = Right(Combo1(pcmbJI).Text, 1)
            GW_JIGYOBU = Mid(Right(Combo1(pcmbSM).Text, 4), 3, 1)
            GW_NAIGAI = Right(Combo1(pcmbSM).Text, 1)

        Case Else
    
    End Select
    
    Call Tab_Ctrl(Shift)        '移動
    

End Sub

Private Sub Command1_Click(Index As Integer)
Dim sts     As Integer
Dim yn      As Integer
Dim X_i     As Integer
Dim X_j     As Integer

Dim W_After     As String

Dim W_Date  As String
Dim wYY     As String * 4
Dim wMM     As String * 2
Dim wDD     As String * 2

Dim W_PC        As String
Dim W_DT        As String
Dim c           As String
Dim W_Path      As String
Dim W_CNT       As Long

Dim W_STR       As String

    Select Case Index
    
        Case FuncCOR
        
            If IsNull(TDBGrid1.Bookmark) Then Exit Sub
            
            If Grid_Cor_M <> True Then
                Exit Sub
            End If
            
            X_i = ORDR_GRID.UpperBound(1)
            
            'Set ORDR_GRID = TDBGrid1.Array
            
            
            
            'TDBGrid1.Update
    
    
            For X_j = Min_Row To ORDR_GRID.UpperBound(1)
                Cor_Row = X_j
                For X_i = Col_DEL To Col_FIN_DT
                    
                    W_STR = Trim(ORDR_GRID(X_j, Col_ORDR_NO)) & Trim(ORDR_GRID(X_j, Col_OYA_ITEM))
                    
                    W_STR = Trim(ORDR_GRID(X_j, Col_OYA_ITEM)) '親品目の未設定は無視！ 2008/10/21
                    If W_STR <> "" Then
                    
                        W_After = ORDR_GRID(X_j, X_i)
                        Cor_Row = X_j
                        If Grid_Err_Chk(X_i, W_After) Then
                            TDBGrid1.ReBind
                            TDBGrid1.Update
                                'TDBGrid1.MoveFirst
                            TDBGrid1.ScrollBars = dbgAutomatic
                            TDBGrid1.SetFocus
                            Exit Sub
                        End If
                    
                    End If
                    
                Next X_i
                
            Next X_j
            'TDBGrid1.ReBind
            'TDBGrid1.Update
                'TDBGrid1.MoveFirst
            'TDBGrid1.ScrollBars = dbgAutomatic
                    
            yn = MsgBox("更新しますか？", vbYesNo + vbDefaultButton1 + vbQuestion, "確認入力")
            yn = vbYes
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            '更新処理
            If Rec_UPDT(True) Then
                MsgBox "更新失敗しました。", vbExclamation
                
                Exit Sub
            End If
            row = ORDR_GRID.UpperBound(1)
            Grid_Cor_M = False
            Grid_Req_M = True
            
            '2008/09/27             「データ出力」を有効にする為、再表示を止めた。
            'If Data_Disp() Then
            '    MsgBox "指定条件の注文情報で、表示失敗！", vbExclamation
            'End If
            
            DoEvents
            
            If ODR10101.MousePointer <> vbDefault Then
                Call Input_UnLock
            End If
            
            If Data_Out_Need = 1 Then
            
                yn = MsgBox("使用月変更行　データ出力しますか？", vbYesNo + vbDefaultButton1 + vbQuestion, "確認入力")
                
                If yn = vbYes Then
                
                                                    '変更テキストＦ フルパス取込み
                    sts = GetIni("FILE", "HENKOU", "SYS_ODR1010", c)
                    If sts <> False Then
                        Call Log_Out(LOG_F, "SYS_ODR1010.INI [HENKOU]読み込みエラー")
                        Exit Sub
                    End If
                    W_Path = RTrim(c)
                    
                    c = Space(255)
                    If GetComputerNameA(c, 255) <> 0 Then
                        W_PC = Left(c, InStr(c, vbNullChar) - 1)
                    Else
                        W_PC = "000"
                    End If
                    
                    W_DT = Right(Format(Date, "yyyymmdd"), 6)
                    
                    W_DT = W_DT & "_" & Left(Text1(ptxUSE_YY), 4) & Right(Text1(ptxUSE_YY), 2)
                    
                    
                    X_i = InStr(1, W_Path, "*") - 1
                    If X_i <= 0 Then
                        X_i = Len(Trim(W_Path)) - 4
                    End If
                    
                    W_Path = Left(W_Path, X_i) & "_" & W_PC & "_" & W_DT & ".CSV" 'Right(FullPath, 4)
        
        
        
                    Set ORDR_GRID = TDBGrid1.Array
                    
                    If Data_Out(W_Path, W_CNT) Then
                        MsgBox "出力失敗！", vbExclamation
                        Command1(FuncEND).SetFocus
                        Exit Sub
                    End If
                    MsgBox W_Path & Chr(13) & Chr(10) & _
                                "　" & W_CNT & "件 " & "出力しました！"
                Data_Out_Need = 0
                End If
                
            End If
            
            
            Text1(ptxTOP%).SetFocus
            Call Text1_GotFocus(ptxTOP%)
            
            
            Exit Sub
            
        Case FuncREQ
        
            If IsNull(TDBGrid1.Bookmark) Then Exit Sub
            'If row <= 0 Then
            '    MsgBox "表示月の注文データありません。" & Chr(13) & Chr(10) & _
            '            "　展開処理不能です。", vbExclamation
            '    Exit Sub
            'End If
            
            
            If Grid_Cor_M = True Then
                MsgBox "更新処理が未実行です！！" & Chr(13) & Chr(10) & _
                        "　更新して下さい。", vbExclamation
                Exit Sub
            End If
            
            'yn = MsgBox("展開処理は時間がかかります！！" & Chr(13) & Chr(10) & _
            '             "　展開処理しますか？", vbYesNo + vbDefaultButton2 + vbQuestion, "確認入力")
                      
                      
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            '2010.06.06 下記を追加！
'            If GetIni("PR00030", "LAST_SHIME_DT01", "P_SYS", c) Then           '2016.01.12
            If GetIni("PR00030", "LAST_SHIME_DT01", "PR00030", c) Then          '2016.01.12
                GW_TOUGETU = Left(Format(Date, "yyyymmdd"), 6)
                GW_SHIMEBI = Format(Date, "yyyymmdd")
            Else
                GW_TOUGETU = Left(Format(Trim(c), "yyyymmdd"), 6)
                
                GW_SHIMEBI = Format(Trim(c), "yyyymmdd")
                
            
            End If
            
            wYY = Left(GW_TOUGETU, 4)
            wMM = Right(GW_TOUGETU, 2)
            wDD = Right(GW_SHIMEBI, 2)
            Text1(ptxSHIME_DT) = Right(wYY, 2) & "/" & wMM & "/" & wDD
            
            W_Date = Left(GW_SHIMEBI, 4) & "/" & Mid(GW_SHIMEBI, 5, 2) & "/" & Right(GW_SHIMEBI, 2)
            GW_MAX_YYMM = Left(Format(DateAdd("m", 20, W_Date), "yyyy/mm/dd"), 7)
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< 2010/06/06ここまで
            
            
            yn = MsgBox("展開処理しますか？", vbYesNo + vbDefaultButton2 + vbQuestion, "確認入力")
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            'MsgBox "★展開★します！　(`_´ )ゞ"
            If Require_Set() Then
                Call Input_UnLock
                MsgBox "展開処理でエラー！", vbExclamation
                Command1(FuncEND).SetFocus
                
                Exit Sub
            End If
                
            Grid_Req_M = False
            
            If Data_Disp() Then
                MsgBox "指定条件の注文情報で、表示失敗！", vbExclamation
            End If
            
            DoEvents
            '2012/03/15 北村さんからの要望（電話など）で下記を追加
'Private Const Col_BUNNO% = 3                '分納回数
'Private Const Col_OYA_ITEM% = 4             '親部品コード
'Private Const Col_OYA_NM% = 5               '親部品コード
'Private Const Col_ORDR_QTY% = 6             '注文数量
'Private Const Col_NOUKI% = 7                '親部品　注文納期
'Private Const Col_OK_DT% = 8                '組立可能日
'Private Const Col_KAITO% = 9                '親部品　回答納期
'Private Const Col_USE_YM% = 10              '使用月
'Private Const Col_FIN_DT% = 11              '完了日付
            'Call TDBGrid1_HeadClick_YES(Col_NOUKI)
            Call TDBGrid1_HeadClick_YES(Col_FIN_DT%)
            DoEvents
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  ココマデ
            
            
            If ODR10101.MousePointer <> vbDefault Then
                Call Input_UnLock
            End If
            
            Text1(ptxTOP%).SetFocus
            Call Text1_GotFocus(ptxTOP%)
            'MsgBox "☆展開処理☆しました！(･･;)"
            
        Case FuncOUT
            If IsNull(TDBGrid1.Bookmark) Then Exit Sub
            
            If row <= 0 Then
                Exit Sub
            End If
            
            'If Grid_Cor_M = True Then
                
                'Set ORDR_GRID = TDBGrid1.Array
                
                'TDBGrid1.Update
                
        
                For X_j = Min_Row To ORDR_GRID.UpperBound(1)
                
                    For X_i = Col_DEL To Col_FIN_DT%
                        
                        W_After = ORDR_GRID(X_j, X_i)
                        Cor_Row = X_j
                        If Grid_Err_Chk(X_i, W_After) Then
                            TDBGrid1.ReBind
                            TDBGrid1.Update
                                'TDBGrid1.MoveFirst
                            TDBGrid1.ScrollBars = dbgAutomatic
                            TDBGrid1.SetFocus
                            Exit Sub
                        End If
                    
                    Next X_i
                    
                Next X_j
                
            'End If
            
            'TDBGrid1.ReBind
            'TDBGrid1.Update
            '    'TDBGrid1.MoveFirst
            'TDBGrid1.ScrollBars = dbgAutomatic
                    
            yn = MsgBox("画面表示内容　データ出力しますか？", vbYesNo + vbDefaultButton1 + vbQuestion, "確認入力")
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            'If SDC_FLD_GET("SYS_ODR1010", "OUT_1010", W_Path) Then
            '    Text1(ptxTOP).SetFocus
            '    Call Text1_GotFocus(ptxTOP)
            '    Exit Sub
            'End If
                                            '変更テキストＦ フルパス取込み
            sts = GetIni("FILE", "ZENKEN", "SYS_ODR1010", c)
            If sts <> False Then
                Call Log_Out(LOG_F, "SYS_ODR1010.INI [ZENKEN]読み込みエラー")
                Exit Sub
            End If
            W_Path = RTrim(c)
        
        
            
            c = Space(255)
            If GetComputerNameA(c, 255) <> 0 Then
                W_PC = Left(c, InStr(c, vbNullChar) - 1)
            Else
                W_PC = "000"
            End If
            
'            W_DT = Right(Format(Date, "yyyymmdd"), 6)  '2016.01.18
            W_DT = Format(Date, "yyyymmdd") & Mid(Format(Now, "yyyymmddhhmm"), 9, 4)         '2016.01.18
            
            W_DT = W_DT & "_" & Left(Text1(ptxUSE_YY), 4) & Right(Text1(ptxUSE_YY), 2)
            
            
            X_i = InStr(1, W_Path, "*") - 1
            If X_i <= 0 Then
                X_i = Len(Trim(W_Path)) - 4
            End If
            
            W_Path = Left(W_Path, X_i) & "_" & W_PC & "_" & W_DT & ".CSV" 'Right(FullPath, 4)



            Set ORDR_GRID = TDBGrid1.Array
            
            If Data_Out2(W_Path, W_CNT) Then
                MsgBox "出力失敗！", vbExclamation
                Command1(FuncEND).SetFocus
                Exit Sub
            End If
            MsgBox W_Path & Chr(13) & Chr(10) & _
                        "　" & W_CNT & "件 " & "出力しました！"
            
            Command1(FuncEND).SetFocus
            Exit Sub
            
        Case FuncEND
            If Grid_Cor_M = True Then
                yn = MsgBox("更新されていません！！" & Chr(13) & Chr(10) & _
                            "　終了しますか？", vbYesNo + vbDefaultButton2 + vbQuestion, "確認入力")
            Else
                If Grid_Req_M = True Then
                    yn = MsgBox("展開処理されていません！！" & Chr(13) & Chr(10) & _
                                "　終了しますか？", vbYesNo + vbDefaultButton2 + vbQuestion, "確認入力")
                'yn = MsgBox("終了しますか？", vbYesNo + vbDefaultButton1 + vbQuestion, "確認入力")
                'yn = vbYes
                Else
                    yn = vbYes
                End If
            End If
            
            
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            
'2008.04.10            Call TEMP_DEL
            
            
            Unload Me
    
    End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)




    If Shift = vbShiftMask Then
        If KeyCode = vbKeyZ Then
        
            If TDBGrid1.Columns(Col_SV_NOUKI).Visible = True Then
                TDBGrid1.Columns(Col_SV_NOUKI).Visible = False
                TDBGrid1.Columns(Col_SV_KAITO).Visible = False

            Else


                TDBGrid1.Columns(Col_SV_NOUKI).Visible = True
                TDBGrid1.Columns(Col_SV_KAITO).Visible = True

            End If
        End If
    End If

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

Dim W_STR   As String

Dim X_i     As Integer
Dim X_j     As Integer
Dim X_K     As Integer
Dim X_L     As Integer


Dim wYY     As String * 4
Dim wMM     As String * 2
Dim wDD     As String * 2

Init_F_10101 = 0

'コモンコントロールを初期化する
cc.dwSize = Len(cc)
cc.dwICC = ICC_BAR_CLASSES

'ステータスウィンドウを作成する
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "親部品　注文情報登録", Me.hwnd, 0)
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
    Show
    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。", vbExclamation
        End
    End If
    
    
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
    
                                '管理マスタＯＰＥＮ
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '移動歴ＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '親品番注文Ｆ　ＯＰＥＮ
    If ODR_ORDER_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '所要量展開Ｆ
    If ODR_REQUIRE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '注文中間Ｆ
    'If ODR_TEMP1_Open(BtOpenNomal) Then
    '    Unload Me
    'End If
                                
                                
                                '構成マスタＯＰＥＮ
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '資材発注ファイルＯＰＥＮ
    If P_SHORDER_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '資材受入履歴ファイルＯＰＥＮ
    If P_SHUKEIRE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '受払先マスターＯＰＥＮ
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫マスタＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '半製品管理ＯＰＥＮ
    If ODR_HANSEIHIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '月初在庫ＯＰＥＮ
    'If ODR_ZAIKO_Open(BtOpenNomal) Then
    '    Unload Me
    'End If
    
    
    
    
    Load ODR10102
    Load ODR10103
    
'テキストを設定する
    Text1(ptxUSE_YY) = Left(Format(Date, "yyyy/mm/dd"), 7)
    
    
    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc
    
    '仕向け先のセット
    If Code_Set_Proc(pcmbSM, P_KBN04_CD, 0) Then
        Unload Me
    End If
    Combo1(pcmbSM).ListIndex = 0
'事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止します。"
        End
    End If
    
    If SET_JGYOBU_T Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止します。"
        End
    End If
    
    GW_SIMUKE = Left(Right(Combo1(pcmbSM).Text, 4), 2)
    'GW_JIGYOBU = Right(Combo1(pcmbJI).Text, 1)
    GW_JIGYOBU = Mid(Right(Combo1(pcmbSM).Text, 4), 3, 1)
    GW_NAIGAI = Right(Combo1(pcmbSM).Text, 1)
    
    'GW_SIMUKE = "01"
    'GW_JIGYOBU = "B"
    'GW_NAIGAI = "1"
    
    GW_HINGAI = ""
    
    
    '2008.07.02 下記を追加！
'    If GetIni("PR00030", "LAST_SHIME_DT01", "P_SYS", c) Then           '2016.01.12
    If GetIni("PR00030", "LAST_SHIME_DT01", "PR00030", c) Then          '2016.01.12
        GW_TOUGETU = Left(Format(Date, "yyyymmdd"), 6)
        GW_SHIMEBI = Format(Date, "yyyymmdd")
    Else
        GW_TOUGETU = Left(Format(Trim(c), "yyyymmdd"), 6)
        
        GW_SHIMEBI = Format(Trim(c), "yyyymmdd")
        
    
    End If
    
    wYY = Left(GW_TOUGETU, 4)
    wMM = Right(GW_TOUGETU, 2)
    wDD = Right(GW_SHIMEBI, 2)
        'If wDD <= "20" Then
        
        'Else
            
        '    wMM = Format(CInt(wMM) + 1, "00")
        
        '    If wMM > "12" Then
        '        wYY = Format(CInt(wYY) + 1, "0000")
        '        wMM = "01"
        '    End If
        'End If
    W_Date = Left(GW_SHIMEBI, 4) & "/" & Mid(GW_SHIMEBI, 5, 2) & "/" & Right(GW_SHIMEBI, 2)
    
    Text1(ptxUSE_YY) = wYY & "/" & wMM
    Text1(ptxSHIME_DT) = Right(wYY, 2) & "/" & wMM & "/" & wDD
    
    GW_MAX_YYMM = Left(Format(DateAdd("m", 20, W_Date), "yyyy/mm/dd"), 7)
    
    
    
    '2008/09    在訂±区分の取得
    'If GetIni("ZAITEI", "PLUS", "SYS_ODR1010", c) Then
    '    GW_PURA = "+"
    'Else
    '    GW_PURA = Trim(c)
    'End If
    'If GetIni("ZAITEI", "MINUS", "SYS_ODR1010", c) Then
    '    GW_MAINA = "-"
    'Else
    '    GW_MAINA = Trim(c)
    'End If
    
    '2009/03/04 要因をテーブルにした！　(*_*;
    Erase GW_PURA
    Erase GW_MAINA
    If GetIni("ZAITEI", "PLUS", "SYS_ODR1010", c) Then
        
    Else
        W_STR = Trim(c)
        X_j = Len(W_STR) / 3 '+ 1
        X_K = 0
        For X_i = 1 To X_j
            X_L = (X_i - 1) * 3 + 1
            GW_PURA(X_K) = Mid(W_STR, X_L, 2)
            X_K = X_K + 1
        Next X_i
    End If
    
    If GetIni("ZAITEI", "MINUS", "SYS_ODR1010", c) Then
        
    Else
        W_STR = Trim(c)
        X_j = Len(W_STR) / 3 '+ 1
        X_K = 0
        For X_i = 1 To X_j
            X_L = (X_i - 1) * 3 + 1
            GW_MAINA(X_K) = Mid(W_STR, X_L, 2)
            X_K = X_K + 1
        Next X_i
    End If
    
    
    If USE_YM_SAVE Then
        MsgBox "初期化失敗！", vbExclamation
        Unload Me
    End If
    
    'Combo1(pcmbSM).SetFocus
       
    Text1(ptxTOP).SetFocus
       
    Grid_Cor_M = False
    Grid_Req_M = False
    Data_Out_Need = 0
    
    
    ODR10101.Caption = ODR10101.Caption & Last_Update$  '2016.12.03
    
    
    row = Min_Row - 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim yn      As Integer

    If UnloadMode = 1 Then Exit Sub
    
    If Grid_Cor_M = True Then
        yn = MsgBox("更新されていません！！" & Chr(13) & Chr(10) & _
                    "　終了しますか？", vbYesNo + vbDefaultButton2 + vbQuestion, "確認入力")
    Else
        yn = MsgBox("終了しますか？", vbYesNo + vbDefaultButton1 + vbQuestion, "確認入力")
        'yn = vbYes
    End If
    
    If yn = vbNo Then
        Cancel = 1
        Exit Sub
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts     As Integer

    sts = BTRV(BtOpClose, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ITEM, Len(K0_ODR_ZK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_ZAIKO")
        End If
    End If


    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "IDO")
        End If
    End If

    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ITEM")
        End If
    End If

    sts = BTRV(BtOpClose, ODR_HANSEIHIN_POS, ODR_HANSEIHIN_O_REC, Len(ODR_HANSEIHIN_O_REC), K0_ODR_HANSEIHIN, Len(K0_ODR_HANSEIHIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_HANSEIHIN")
        End If
    End If

    sts = BTRV(BtOpClose, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_KENTO")
        End If
    End If
    
    sts = BTRV(BtOpClose, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_ORDER, Len(K0_ODR_ORDER), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_ORDER")
        End If
    End If
       
    sts = BTRV(BtOpClose, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_REQ")
        End If
    End If
    
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
    
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "P_CODE")
        End If
    End If
    
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "P_COMPO")
        End If
    End If
    
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "P_KANRI")
        End If
    End If

    
    sts = BTRV(BtOpClose, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "P_SHORDER")
        End If
    End If
    
    sts = BTRV(BtOpClose, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K0_P_SHUKEIRE, Len(K0_P_SHUKEIRE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "P_SHUKEIRE")
        End If
    End If
    
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "P_UKEHARAI")
        End If
    End If
    
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "TANTO")
        End If
    End If
    

    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ZAIKO")
        End If
    End If


    End
End Sub

Private Sub SHORI_Click(Index As Integer)
Dim yn      As Integer


    Select Case Index
    
        Case 0      '更新
            Call Command1_Click(FuncCOR)
            
        Case 1      '展開
            Call Command1_Click(FuncREQ)
            
        Case 2      'データ出力
            Call Command1_Click(FuncOUT)
            
        Case 3      '画面印刷
            yn = MsgBox("画面印刷しますか？", vbYesNo + vbDefaultButton2 + vbQuestion, "確認入力")
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)
        
            
        
        Case 4      '終了
            Call Command1_Click(FuncEND)
    
    End Select


End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
Dim W_STR       As String
    
Dim W_Before    As String
Dim W_After     As String

    If IsNull(TDBGrid1.Bookmark) Then Exit Sub

    If TDBGrid1.Bookmark <= 0 Then Exit Sub
    
    If Not IsNumeric(TDBGrid1.Bookmark) Then Exit Sub
    
    Cor_Row = TDBGrid1.Bookmark
    
    'W_Before = Trim(ORDR_GRID(Cor_Row, ColIndex))
    W_After = Trim(TDBGrid1.Text)
    
    Set ORDR_GRID = TDBGrid1.Array
    TDBGrid1.Update
    
    'If W_Before <> W_After Then
    '    Grid_Cor_M = True
    'End If
    
    'If Grid_Err_Chk(ColIndex, W_Before, W_After) Then
        
    '    Exit Sub
    'End If
    

End Sub

Private Sub TDBGrid1_BeforeInsert(Cancel As Integer)

    ORDR_GRID.ReDim Min_Row, ORDR_GRID.Count(1), Min_Col, Max_Col
    
End Sub

Private Sub TDBGrid1_Change()

    Grid_Cor_M = True

End Sub

Private Sub TDBGrid1_DblClick()
Dim W_ORDR  As String
Dim W_STR   As String

    If IsNull(TDBGrid1.Bookmark) Then Exit Sub
    
    If TDBGrid1.Bookmark = -1 Then
    
    Else
        Set ORDR_GRID = TDBGrid1.Array
        
        '       分納の可否チェック
        '
        '       ①親部品注文№が指定してある事！
        '       ②未完了の事！
        '       ③分納の親（基）情報を指示する事！?
        '
        
        
'        W_ORDR = ORDR_GRID(TDBGrid1.Bookmark, Col_ORDR_NO) '親部品注文№
'        If Trim(W_ORDR) = "" Then Exit Sub
        If Option1(0).Value Then
            If Trim(ORDR_GRID(TDBGrid1.Bookmark, Col_FIN_DT)) <> "" Then
                MsgBox "完了済み　分納指示不可！", vbExclamation
                Exit Sub
            End If
        End If
        
        'W_Str = ORDR_GRID(TDBGrid1.Bookmark, Col_BUNNO%)    '分納回数
        'If Trim(W_Str) = "" Then
        '    W_Str = "0"
        'End If
        'If CInt(Trim(W_Str)) <> 0 Then
        '    MsgBox "分納指示　不可！", vbExclamation
        '    Exit Sub
        'End If
        
        
        '           分納指示画面に移行！
        Key_SIMUKE = GW_SIMUKE
        Key_JIGYOBU = GW_JIGYOBU
        Key_NAIGAI = GW_NAIGAI
        
        W_STR = Trim(ORDR_GRID(TDBGrid1.Bookmark, Col_USE_YM))
        If Trim(W_STR) = "" Then
            W_STR = Left(Text1(ptxUSE_YY), 4) & Right(Text1(ptxUSE_YY), 2)
        Else
            W_STR = Left(ORDR_GRID(TDBGrid1.Bookmark, Col_USE_YM), 4) & Right(ORDR_GRID(TDBGrid1.Bookmark, Col_USE_YM), 2)
        End If
        Key_USE_YM = W_STR
        
        Key_INS_NO = Trim(ORDR_GRID(TDBGrid1.Bookmark, Col_KEY))
        
        Key_ORDER_NO = W_ORDR
        
        Key_BUN_NO = Trim(ORDR_GRID(TDBGrid1.Bookmark, Col_BUNNO))
                
        DIS_ORDR_NO = ORDR_GRID(TDBGrid1.Bookmark, Col_ORDR_NO)    '親部品注文№
        DIS_BUNNO = ORDR_GRID(TDBGrid1.Bookmark, Col_BUNNO%)        '分納回数
        DIS_OYA_ITEM = ORDR_GRID(TDBGrid1.Bookmark, Col_OYA_ITEM)  '親部品コード
        GW_HINGAI = DIS_OYA_ITEM
        
        DIS_ORDR_QTY = ORDR_GRID(TDBGrid1.Bookmark, Col_ORDR_QTY)  '注文数量
        DIS_NOUKI = ORDR_GRID(TDBGrid1.Bookmark, Col_NOUKI)        '親部品　注文納期
        DIS_OK_DT = ORDR_GRID(TDBGrid1.Bookmark, Col_OK_DT)         '組立可能日
        DIS_KAITO = ORDR_GRID(TDBGrid1.Bookmark, Col_KAITO)        '親部品　回答納期
        DIS_USE_YM = ORDR_GRID(TDBGrid1.Bookmark, Col_USE_YM)      '使用月
        DIS_FIN_DT = ORDR_GRID(TDBGrid1.Bookmark, Col_FIN_DT)      '完了日付
        DIS_KEY = ORDR_GRID(TDBGrid1.Bookmark, Col_KEY)            'データＫｅｙ情報
    
        DoEvents
        If Option1(0).Value Then

            ODR10102.Show vbModal
        Else
            ODR10103.Show vbModal
        End If
        
        
        If ODR10102_Return = True Then Exit Sub     'キャンセル
        
        '分納分を反映して再表示する。
        
        If Data_Disp Then
            MsgBox "指定条件の注文情報で、表示失敗！", vbExclamation
            Call Input_UnLock                             '画面項目ロック
            Call Text1_GotFocus(ptxTOP)
            Text1(ptxTOP%).SetFocus
            Exit Sub
        End If
        
    End If


End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
Dim yn          As Integer
Dim W_Index     As Integer

Dim X_i         As Long

    If IsNull(TDBGrid1.Bookmark) Then Exit Sub
    'TDBGrid1.Bookmark = -1
    W_Index = ColIndex
    
    If row <= 1 Then Exit Sub
    
    
    
    yn = MsgBox("並べ換えますか？", vbYesNo + vbExclamation, "確認入力")
    If yn <> vbYes Then Exit Sub
    
    
    'Set ORDR_GRID = TDBGrid1.Array
    
    Select Case ColIndex
        Case Col_ORDR_NO           '親部品　注文№
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING, _
                        Col_BUNNO, XORDER_ASCEND, XTYPE_STRING
                        
                
        Case Col_OYA_ITEM          '親部品コード
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_BUNNO, XORDER_ASCEND, XTYPE_STRING
        
        Case Col_OYA_NM            '親部品名
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_OYA_NM, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_BUNNO, XORDER_ASCEND, XTYPE_STRING
        
        Case Col_NOUKI             '親部品　注文納期
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_NOUKI, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING
        
        Case Col_ORDR_QTY             '親部品　注文数量
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_ORDR_QTY, XORDER_ASCEND, XTYPE_DOUBLE, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING
        
        Case Col_OK_DT             '組立可能日
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_OK_DT, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING
        
        
        Case Col_KAITO             '親部品　回答納期
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_KAITO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING
        
        Case Col_USE_YM            '使用月
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_USE_YM, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING
        
        Case Col_FIN_DT            '完了日付
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_FIN_DT, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING
        
        Case Col_DEL               '削除マーク
            'MsgBox "削除マークでＳＯＲＴ？　(･･;)" & Chr(13) & Chr(10) & _
            '       "　そりゃないでしょ！！　(^_^;)", vbExclamation
            
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_DEL, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING
        Case Col_BUNNO             '分納回数
            'MsgBox "分納回数でＳＯＲＴ？　(･･;)" & Chr(13) & Chr(10) & _
            '       "意味不明の並び順で、訳が分からなくなります！！　(^_^;)", vbExclamation
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_BUNNO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING
            
            
            
        Case Else
            MsgBox "並べ替指定 除外項目！", vbExclamation
            Exit Sub
        
        
    End Select

    For X_i = Min_Row To ORDR_GRID.UpperBound(1)
        ORDR_GRID(X_i, Col_No) = X_i
    Next X_i

    Set TDBGrid1.Array = ORDR_GRID
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    TDBGrid1.Bookmark = 1
    
    DoEvents
    
End Sub

Private Sub TDBGrid1_HeadClick_YES(ByVal ColIndex As Integer)
'                               2012/03/14 当該SUBを追加
'                               目的 : 展開後の表示順を変更する。
'
Dim yn          As Integer
Dim W_Index     As Integer

Dim X_i         As Long

    If IsNull(TDBGrid1.Bookmark) Then Exit Sub
    'TDBGrid1.Bookmark = -1
    W_Index = ColIndex
    
    If row <= 1 Then Exit Sub
    
    
    
'    yn = MsgBox("並べ換えますか？", vbYesNo + vbExclamation, "確認入力")
'    If yn <> vbYes Then Exit Sub
    
    
    'Set ORDR_GRID = TDBGrid1.Array
    
    For X_i = Min_Row To ORDR_GRID.UpperBound(1)
        If Trim(ORDR_GRID(X_i, Col_OK_DT)) = "" Then
            ORDR_GRID(X_i, Col_OK_DT) = "9999"
        End If
    Next X_i
    
    Select Case ColIndex
        Case Col_ORDR_NO           '親部品　注文№
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING, _
                        Col_BUNNO, XORDER_ASCEND, XTYPE_STRING
                        
                
        Case Col_OYA_ITEM          '親部品コード
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_BUNNO, XORDER_ASCEND, XTYPE_STRING
        
        Case Col_OYA_NM            '親部品名
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_OYA_NM, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_BUNNO, XORDER_ASCEND, XTYPE_STRING
        
        Case Col_NOUKI             '親部品　注文納期
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_NOUKI, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING
        
        Case Col_ORDR_QTY             '親部品　注文数量
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_ORDR_QTY, XORDER_ASCEND, XTYPE_DOUBLE, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING
        
        Case Col_OK_DT             '組立可能日
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_OK_DT, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING
        
        
        Case Col_KAITO             '親部品　回答納期
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_KAITO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING
        
        Case Col_USE_YM            '使用月
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_USE_YM, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING
        
        Case Col_FIN_DT            '完了日付
'            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
'                        Col_FIN_DT, XORDER_ASCEND, XTYPE_STRING, _
'                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
'                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING

            '2012/03/15 下記に変更。
            '               完了日付（状況）、組立可能日、注文納期、親品番、親注文№
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_KEY_FIN, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OK_DT, XORDER_ASCEND, XTYPE_STRING, _
                        Col_NOUKI, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING
        
        
        Case Col_DEL               '削除マーク
            'MsgBox "削除マークでＳＯＲＴ？　(･･;)" & Chr(13) & Chr(10) & _
            '       "　そりゃないでしょ！！　(^_^;)", vbExclamation
            
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_DEL, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING
        Case Col_BUNNO             '分納回数
            'MsgBox "分納回数でＳＯＲＴ？　(･･;)" & Chr(13) & Chr(10) & _
            '       "意味不明の並び順で、訳が分からなくなります！！　(^_^;)", vbExclamation
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_BUNNO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ORDR_NO, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM, XORDER_ASCEND, XTYPE_STRING
            
            
            
        Case Else
            MsgBox "並べ替指定 除外項目！", vbExclamation
            Exit Sub
        
        
    End Select

    For X_i = Min_Row To ORDR_GRID.UpperBound(1)
        If Trim(ORDR_GRID(X_i, Col_OK_DT)) = "9999" Then
            ORDR_GRID(X_i, Col_OK_DT) = ""
        End If
    Next X_i


    For X_i = Min_Row To ORDR_GRID.UpperBound(1)
        ORDR_GRID(X_i, Col_No) = X_i
    Next X_i

    Set TDBGrid1.Array = ORDR_GRID
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    TDBGrid1.Bookmark = 1
    
    DoEvents
    
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
    
    If Index = ptxTOP And Init_F_10101 = 0 Then
        If Data_Disp Then
            MsgBox "指定条件の注文情報で、表示失敗！", vbExclamation
            Call Text1_GotFocus(ptxTOP%)
            Text1(ptxTOP%).SetFocus
            Exit Sub
        End If
        '2012/03/15 下記を北村氏からの要望により追加
        DoEvents
        Call TDBGrid1_HeadClick_YES(Col_FIN_DT%)
        DoEvents
        '>>>>>>>>>>>>>>>>>  ココマデ
        

        Init_F_10101 = 1
        Call Text1_GotFocus(ptxUSE_YY)
        Text1(ptxUSE_YY).SetFocus
        Exit Sub
    End If
    
    If Index = ptxUSE_YY Then
        If Data_Disp Then
            MsgBox "指定条件の注文情報で、表示失敗！", vbExclamation
            Call Text1_GotFocus(ptxTOP%)
            Text1(ptxTOP%).SetFocus
            Exit Sub
        End If
        '2012/03/15 下記を北村氏からの要望により追加
        DoEvents
        Call TDBGrid1_HeadClick_YES(Col_FIN_DT%)
        DoEvents
        '>>>>>>>>>>>>>>>>>  ココマデ
        
        TDBGrid1.SetFocus
        
        Exit Sub
    End If
    
    Call Tab_Ctrl(Shift)    '移動
    
End Sub

