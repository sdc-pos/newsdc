VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form F1300101 
   Caption         =   "送り状発行データ処理"
   ClientHeight    =   12060
   ClientLeft      =   2025
   ClientTop       =   -5145
   ClientWidth     =   16050
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
   ScaleHeight     =   12060
   ScaleWidth      =   16050
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      Index           =   0
      Left            =   2520
      OLEDragMode     =   1  '自動
      OLEDropMode     =   1  '手動
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   8055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終　了"
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
      Index           =   7
      Left            =   14520
      TabIndex        =   10
      ToolTipText     =   "処理を終了します"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "日本通運"
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
      Index           =   6
      Left            =   11520
      TabIndex        =   9
      ToolTipText     =   "日本通運向けデータを作成"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "第一貨物"
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
      Index           =   5
      Left            =   9840
      TabIndex        =   8
      ToolTipText     =   "第一貨物向けデータを作成"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "久留米"
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
      Index           =   4
      Left            =   8160
      TabIndex        =   7
      ToolTipText     =   "久留米向けデータを作成"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "福山"
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
      Index           =   3
      Left            =   6480
      TabIndex        =   6
      ToolTipText     =   "福山向けデータを作成"
      Top             =   0
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Index           =   2
      Left            =   5040
      Locked          =   -1  'True
      OLEDragMode     =   1  '自動
      OLEDropMode     =   1  '手動
      TabIndex        =   1
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Index           =   1
      Left            =   1680
      Locked          =   -1  'True
      OLEDragMode     =   1  '自動
      OLEDropMode     =   1  '手動
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "着店設定"
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
      Index           =   2
      Left            =   3720
      TabIndex        =   5
      ToolTipText     =   "着店の設定を行います"
      Top             =   0
      Width           =   1380
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   9735
      Left            =   480
      TabIndex        =   11
      Top             =   1800
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   17171
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "対象"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "№"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "出荷日"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "集約送り先CD"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "送り先CD"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "送り先名"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "売伝"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "伝票番号"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "品番"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "数量"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "注文№"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "得意先CD"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "得意先名"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "備考"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "運送会社"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "営業所CD(便)"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "住所"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "郵便番号"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "TEL"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "件名管理№"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "品番管理№"
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "着店ｺｰﾄﾞ"
      Columns(21).DataField=   ""
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "ID_NO"
      Columns(22).DataField=   ""
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   23
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   714
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=23"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1402"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1296"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1005"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=900"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1667"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1561"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=512"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2910"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2805"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=512"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=4339"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=4233"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=3281"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=3175"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=512"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=953"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=847"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=512"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=2090"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=1984"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=3281"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=3175"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(9).Width=1429"
      Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=1323"
      Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=514"
      Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(51)=   "Column(10).Width=1746"
      Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=1640"
      Splits(0)._ColumnProps(54)=   "Column(10)._ColStyle=516"
      Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(56)=   "Column(11).Width=1931"
      Splits(0)._ColumnProps(57)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(11)._WidthInPix=1826"
      Splits(0)._ColumnProps(59)=   "Column(11)._ColStyle=516"
      Splits(0)._ColumnProps(60)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(61)=   "Column(12).Width=3281"
      Splits(0)._ColumnProps(62)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(12)._WidthInPix=3175"
      Splits(0)._ColumnProps(64)=   "Column(12)._ColStyle=516"
      Splits(0)._ColumnProps(65)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(66)=   "Column(13).Width=4710"
      Splits(0)._ColumnProps(67)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(13)._WidthInPix=4604"
      Splits(0)._ColumnProps(69)=   "Column(13)._ColStyle=516"
      Splits(0)._ColumnProps(70)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(71)=   "Column(14).Width=2487"
      Splits(0)._ColumnProps(72)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(14)._WidthInPix=2381"
      Splits(0)._ColumnProps(74)=   "Column(14)._ColStyle=516"
      Splits(0)._ColumnProps(75)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(76)=   "Column(15).Width=2355"
      Splits(0)._ColumnProps(77)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(78)=   "Column(15)._WidthInPix=2249"
      Splits(0)._ColumnProps(79)=   "Column(15)._ColStyle=516"
      Splits(0)._ColumnProps(80)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(81)=   "Column(16).Width=8229"
      Splits(0)._ColumnProps(82)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(83)=   "Column(16)._WidthInPix=8123"
      Splits(0)._ColumnProps(84)=   "Column(16)._ColStyle=516"
      Splits(0)._ColumnProps(85)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(86)=   "Column(17).Width=2090"
      Splits(0)._ColumnProps(87)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(88)=   "Column(17)._WidthInPix=1984"
      Splits(0)._ColumnProps(89)=   "Column(17)._ColStyle=516"
      Splits(0)._ColumnProps(90)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(91)=   "Column(18).Width=3281"
      Splits(0)._ColumnProps(92)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(93)=   "Column(18)._WidthInPix=3175"
      Splits(0)._ColumnProps(94)=   "Column(18)._ColStyle=516"
      Splits(0)._ColumnProps(95)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(96)=   "Column(19).Width=3281"
      Splits(0)._ColumnProps(97)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(98)=   "Column(19)._WidthInPix=3175"
      Splits(0)._ColumnProps(99)=   "Column(19)._ColStyle=516"
      Splits(0)._ColumnProps(100)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(101)=   "Column(20).Width=1773"
      Splits(0)._ColumnProps(102)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(103)=   "Column(20)._WidthInPix=1667"
      Splits(0)._ColumnProps(104)=   "Column(20)._ColStyle=516"
      Splits(0)._ColumnProps(105)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(106)=   "Column(21).Width=1958"
      Splits(0)._ColumnProps(107)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(108)=   "Column(21)._WidthInPix=1852"
      Splits(0)._ColumnProps(109)=   "Column(21)._ColStyle=516"
      Splits(0)._ColumnProps(110)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(111)=   "Column(22).Width=1720"
      Splits(0)._ColumnProps(112)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(113)=   "Column(22)._WidthInPix=1614"
      Splits(0)._ColumnProps(114)=   "Column(22)._ColStyle=516"
      Splits(0)._ColumnProps(115)=   "Column(22).Order=23"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=67,.parent=1,.bgcolor=&HFFFF00&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=68,.parent=2,.alignment=2,.bold=0,.fontsize=900"
      _StyleDefs(27)  =   ":id=68,.italic=0,.underline=0,.strikethrough=0,.charset=128"
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
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=118,.parent=67"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=115,.parent=68"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=116,.parent=69"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=117,.parent=71"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=82,.parent=67,.alignment=3"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=79,.parent=68"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=80,.parent=69"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=81,.parent=71"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=94,.parent=67,.alignment=0"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=91,.parent=68"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=92,.parent=69"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=93,.parent=71"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=98,.parent=67,.alignment=0"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=95,.parent=68"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=96,.parent=69"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=97,.parent=71"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=20,.parent=67"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=17,.parent=68"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=18,.parent=69"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=19,.parent=71"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=102,.parent=67,.alignment=0"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=99,.parent=68"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=100,.parent=69"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=101,.parent=71"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=114,.parent=67,.alignment=0"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=111,.parent=68"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=112,.parent=69"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=113,.parent=71"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=16,.parent=67"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=13,.parent=68"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=14,.parent=69"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=15,.parent=71"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=24,.parent=67"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=21,.parent=68"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=22,.parent=69"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=23,.parent=71"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=28,.parent=67,.alignment=1"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=25,.parent=68"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=26,.parent=69"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=27,.parent=71"
      _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=32,.parent=67"
      _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=29,.parent=68"
      _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=30,.parent=69"
      _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=31,.parent=71"
      _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=46,.parent=67"
      _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=43,.parent=68"
      _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=44,.parent=69"
      _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=45,.parent=71"
      _StyleDefs(86)  =   "Splits(0).Columns(12).Style:id=50,.parent=67"
      _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=47,.parent=68"
      _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=48,.parent=69"
      _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=49,.parent=71"
      _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=54,.parent=67"
      _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=51,.parent=68"
      _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=52,.parent=69"
      _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=53,.parent=71"
      _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=58,.parent=67"
      _StyleDefs(95)  =   "Splits(0).Columns(14).HeadingStyle:id=55,.parent=68"
      _StyleDefs(96)  =   "Splits(0).Columns(14).FooterStyle:id=56,.parent=69"
      _StyleDefs(97)  =   "Splits(0).Columns(14).EditorStyle:id=57,.parent=71"
      _StyleDefs(98)  =   "Splits(0).Columns(15).Style:id=62,.parent=67"
      _StyleDefs(99)  =   "Splits(0).Columns(15).HeadingStyle:id=59,.parent=68"
      _StyleDefs(100) =   "Splits(0).Columns(15).FooterStyle:id=60,.parent=69"
      _StyleDefs(101) =   "Splits(0).Columns(15).EditorStyle:id=61,.parent=71"
      _StyleDefs(102) =   "Splits(0).Columns(16).Style:id=66,.parent=67"
      _StyleDefs(103) =   "Splits(0).Columns(16).HeadingStyle:id=63,.parent=68"
      _StyleDefs(104) =   "Splits(0).Columns(16).FooterStyle:id=64,.parent=69"
      _StyleDefs(105) =   "Splits(0).Columns(16).EditorStyle:id=65,.parent=71"
      _StyleDefs(106) =   "Splits(0).Columns(17).Style:id=86,.parent=67"
      _StyleDefs(107) =   "Splits(0).Columns(17).HeadingStyle:id=83,.parent=68"
      _StyleDefs(108) =   "Splits(0).Columns(17).FooterStyle:id=84,.parent=69"
      _StyleDefs(109) =   "Splits(0).Columns(17).EditorStyle:id=85,.parent=71"
      _StyleDefs(110) =   "Splits(0).Columns(18).Style:id=90,.parent=67"
      _StyleDefs(111) =   "Splits(0).Columns(18).HeadingStyle:id=87,.parent=68"
      _StyleDefs(112) =   "Splits(0).Columns(18).FooterStyle:id=88,.parent=69"
      _StyleDefs(113) =   "Splits(0).Columns(18).EditorStyle:id=89,.parent=71"
      _StyleDefs(114) =   "Splits(0).Columns(19).Style:id=106,.parent=67"
      _StyleDefs(115) =   "Splits(0).Columns(19).HeadingStyle:id=103,.parent=68"
      _StyleDefs(116) =   "Splits(0).Columns(19).FooterStyle:id=104,.parent=69"
      _StyleDefs(117) =   "Splits(0).Columns(19).EditorStyle:id=105,.parent=71"
      _StyleDefs(118) =   "Splits(0).Columns(20).Style:id=110,.parent=67"
      _StyleDefs(119) =   "Splits(0).Columns(20).HeadingStyle:id=107,.parent=68"
      _StyleDefs(120) =   "Splits(0).Columns(20).FooterStyle:id=108,.parent=69"
      _StyleDefs(121) =   "Splits(0).Columns(20).EditorStyle:id=109,.parent=71"
      _StyleDefs(122) =   "Splits(0).Columns(21).Style:id=122,.parent=67"
      _StyleDefs(123) =   "Splits(0).Columns(21).HeadingStyle:id=119,.parent=68"
      _StyleDefs(124) =   "Splits(0).Columns(21).FooterStyle:id=120,.parent=69"
      _StyleDefs(125) =   "Splits(0).Columns(21).EditorStyle:id=121,.parent=71"
      _StyleDefs(126) =   "Splits(0).Columns(22).Style:id=126,.parent=67"
      _StyleDefs(127) =   "Splits(0).Columns(22).HeadingStyle:id=123,.parent=68"
      _StyleDefs(128) =   "Splits(0).Columns(22).FooterStyle:id=124,.parent=69"
      _StyleDefs(129) =   "Splits(0).Columns(22).EditorStyle:id=125,.parent=71"
      _StyleDefs(130) =   "Named:id=33:Normal"
      _StyleDefs(131) =   ":id=33,.parent=0"
      _StyleDefs(132) =   "Named:id=34:Heading"
      _StyleDefs(133) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(134) =   ":id=34,.wraptext=-1"
      _StyleDefs(135) =   "Named:id=35:Footing"
      _StyleDefs(136) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(137) =   "Named:id=36:Selected"
      _StyleDefs(138) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(139) =   "Named:id=37:Caption"
      _StyleDefs(140) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(141) =   "Named:id=38:HighlightRow"
      _StyleDefs(142) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(143) =   "Named:id=39:EvenRow"
      _StyleDefs(144) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(145) =   "Named:id=40:OddRow"
      _StyleDefs(146) =   ":id=40,.parent=33"
      _StyleDefs(147) =   "Named:id=41:RecordSelector"
      _StyleDefs(148) =   ":id=41,.parent=34"
      _StyleDefs(149) =   "Named:id=42:FilterBar"
      _StyleDefs(150) =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "郵便番号"
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
      Left            =   2040
      TabIndex        =   4
      ToolTipText     =   "郵便番号の再設定を行います(現在、未使用)"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "読込"
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
      TabIndex        =   3
      ToolTipText     =   "出荷予定データを表示します"
      Top             =   0
      Width           =   1380
   End
   Begin VB.Label Label2 
      Caption         =   "件"
      Height          =   255
      Index           =   2
      Left            =   13680
      TabIndex        =   17
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "便"
      Height          =   255
      Index           =   1
      Left            =   5400
      TabIndex        =   16
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "出荷日"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   15
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblFILE_NAME 
      Caption         =   "ファイル名"
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblDisp_Count 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   12240
      TabIndex        =   13
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "読込件数"
      Height          =   255
      Index           =   1
      Left            =   11160
      TabIndex        =   12
      Top             =   1320
      Width           =   975
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "読込"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "郵便番号"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "着点設定"
         Index           =   2
      End
      Begin VB.Menu SHORI 
         Caption         =   "福山"
         Index           =   3
      End
      Begin VB.Menu SHORI 
         Caption         =   "久留米"
         Index           =   4
      End
      Begin VB.Menu SHORI 
         Caption         =   "第一貨物"
         Index           =   5
      End
      Begin VB.Menu SHORI 
         Caption         =   "日本通運"
         Index           =   6
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   7
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "F1300101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim SYUKA           As New XArrayDB



Private Const ptxFILE_NAME% = 0         'ﾌｧｲﾙ名
Private Const ptxSYUKA_YMD% = 1         '出荷日
Private Const ptxINS_BIN% = 2           '便


Private Const Min_Row% = 1              '最小行数

Dim Max_Row    As Integer               'グリッド最大表示件数

Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 22             '最大列数

Private Const colOBJECT% = 0            '対象
Private Const colSYUKA_NO% = 1          '№
Private Const colSYUKA_YMD% = 2         '出荷日
Private Const colCOL_OKURISAKI_CD% = 3  '送り先集約CD
Private Const colOKURISAKI_CD% = 4      '送り先CD
Private Const colOKURISAKI% = 5         '送り先名
Private Const colURIDEN% = 6            '売伝
Private Const colDEN_NO% = 7            '伝票番号
Private Const colHINBAN% = 8            '品番
Private Const colSURYO% = 9             '数量
Private Const colCYU_NO% = 10           '注文№
Private Const colTOKUI_CODE% = 11       '得意先CD
Private Const colTOKUI_NAME% = 12       '得意先名
Private Const colBIKOU% = 13            '備考
Private Const colUNSOU% = 14            '運送会社
Private Const colINS_BIN% = 15          '便（営業所CD）
Private Const colJYUSHO% = 16           '住所
Private Const colYUBIN_NO% = 17         '郵便番号
Private Const colTEL_NO% = 18           '電話番号
Private Const colSEK_KEN_NO% = 19       '件管№　　　■管理№(上)
Private Const colSEK_HIN_NO% = 20       '品管№　　　■管理№(下)
Private Const colTYAKUTEN% = 21         '福山　着店ｺｰﾄﾞ

Private Const colID_NO% = 22            'ID_NO


Private EXCEL_DATA  As Variant


Dim FUKUYAMA_CSV            As String
Dim KURUME_CSV              As String
Dim DAIICHI_CSV             As String
Dim NITTSU_CSV              As String




Private Type fukuyama_tbl_tag
    YUBIN_NO    As String * 7
    CODE        As String * 7
    JYUSHO      As String
    TYAKUTEN    As String * 3
End Type

Dim FUKUYAMA_TBL()  As fukuyama_tbl_tag


Dim TITLE_CSV()     As String




Private Type Hinban_Tbl_tag
    DEN_NO      As String
    Hinban      As String * 16
    SURYO       As Long
End Type

Dim csvOKURISAKI            As String
Dim csvTOKUI_NAME           As String
Dim csvYUBIN_NO             As String
Dim csvJYUSHO               As String
Dim csvTEL_NO               As String
Dim csvHinban_Tbl()         As Hinban_Tbl_tag
Dim csvTYAKUTEN             As String
Dim csvURIDEN               As String

'久留米運輸(県)
Dim KURUME_SELECT()         As String
'日本通運(送り先集約CD)
Dim NITTSU_SELECT_COL_OKURISAKI_CD() _
                            As String
'日本通運(送り先CD)
Dim NITTSU_SELECT_OKURISAKI_CD() _
                            As String
'第一貨物(送り先集約CD)
Dim DAIICHI_SELECT_COL_OKURISAKI_CD() _
                            As String


Dim INPUT_MODE              As Integer      'データ取込みモード



'Private Const LAST_UPDATE_DAY$ = "[F130010] 2017.04.06 13:00"
Private Const LAST_UPDATE_DAY$ = "[F130010] 2017.04.14 09:45"





Private Sub Command1_Click(Index As Integer)

Dim sWk         As String
Dim i           As Long
Dim j           As Long


    Select Case Index



        Case 0          '読込み


            '取込みﾃﾞｰﾀ表示
            
            If INPUT_MODE = 0 Then
                If List_Disp_Proc() Then
                    Unload Me
                End If
            Else
                If List_Disp_Pervasive_Proc() Then
                    Unload Me
                End If
            End If



            If SYUKA.Count(1) > 0 Then
                
                Command1(2).Enabled = True
                SHORI(2).Enabled = True
            
                Command1(3).Enabled = True
                SHORI(3).Enabled = True
            
                Command1(4).Enabled = True
                SHORI(4).Enabled = True
            
                Command1(5).Enabled = True
                SHORI(5).Enabled = True
                
                Command1(6).Enabled = True
                SHORI(6).Enabled = True
            
            
            Else
                
                Command1(2).Enabled = False
                SHORI(2).Enabled = False
            
                Command1(3).Enabled = False
                SHORI(3).Enabled = False
            
                Command1(4).Enabled = False
                SHORI(4).Enabled = False
            
                Command1(5).Enabled = False
                SHORI(5).Enabled = False
                
                Command1(6).Enabled = False
                SHORI(6).Enabled = False
            
            
            End If


        Case 1          '郵便番設定



        Case 2          '着店設定

            Call fukuyama_tyakuten_Set


        Case 3          '福山   出力
        
            Call FUKUYAMA_CSV_OUT

        Case 4 To 6     '久留米～第一貨物   出力
        
            Call OKURI_CSV_OUT(Index)


        Case 7          '終了

            Unload Me
    
    End Select



'    Command1(Index).SetFocus


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
    
    Select Case KeyCode
        Case vbKeyF12
            Unload Me
    End Select
    
    
    

End Sub

Private Sub Form_Load()


Dim c           As String * 512

Dim wkVariant   As Variant
Dim i           As Long


    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[送り状発行]", Me.hwnd, 0)
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
                                '福山CSV
    If GetIni(App.EXEName, "FUKUYAMA", App.EXEName, c) Then
        Beep
        MsgBox "福山向けCSVﾌｧｲﾙ名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    FUKUYAMA_CSV = RTrim(c)
    Command1(3).ToolTipText = Command1(3).ToolTipText & " (" & FUKUYAMA_CSV & ")"
                                
                                '久留米CSV
    If GetIni(App.EXEName, "KURUME", App.EXEName, c) Then
        Beep
        MsgBox "久留米向けCSVﾌｧｲﾙ名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    KURUME_CSV = RTrim(c)
    Command1(4).ToolTipText = Command1(4).ToolTipText & " (" & KURUME_CSV & ")"
                                '第一貨物CSV
    If GetIni(App.EXEName, "DAIICHI", App.EXEName, c) Then
        Beep
        MsgBox "第一貨物向けCSVﾌｧｲﾙ名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    DAIICHI_CSV = RTrim(c)
    Command1(5).ToolTipText = Command1(5).ToolTipText & " (" & DAIICHI_CSV & ")"
                                '日本通運CSV
    If GetIni(App.EXEName, "NITTSU", App.EXEName, c) Then
        Beep
        MsgBox "日本通運向けCSVﾌｧｲﾙ名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    NITTSU_CSV = RTrim(c)
    Command1(6).ToolTipText = Command1(6).ToolTipText & " (" & NITTSU_CSV & ")"

    '福山着店ｺｰﾄﾞ獲得
    If fukuyama_tyakuten() Then
        End
    End If




    'CSVタイトル獲得
    If GetIni(App.EXEName, "TITLE_CSV", App.EXEName, c) Then
        c = ""
    End If
    wkVariant = Split(Trim(c), ",", -1)
            
    For i = 0 To UBound(wkVariant)
        ReDim Preserve TITLE_CSV(0 To i)
    
        TITLE_CSV(i) = wkVariant(i)
    
    
    
    Next i
    




    '久留米運輸選択条件　(県)
    If GetIni(App.EXEName, "KURUME_SELECT", App.EXEName, c) Then
        c = "＊"
    End If
    wkVariant = Split(Trim(c), ",", -1)
    For i = 0 To UBound(wkVariant)
        
        ReDim Preserve KURUME_SELECT(0 To i)
        KURUME_SELECT(i) = wkVariant(i)
    
    Next i
    
    '日本通運選択条件　(送り先集約CD)
    If GetIni(App.EXEName, "NITTSU_SELECT_COL_OKURISAKI_CD", App.EXEName, c) Then
        c = "＊"
    End If
    wkVariant = Split(Trim(c), ",", -1)
    For i = 0 To UBound(wkVariant)
        ReDim Preserve NITTSU_SELECT_COL_OKURISAKI_CD(0 To i)
    
        NITTSU_SELECT_COL_OKURISAKI_CD(i) = wkVariant(i)
    
    Next i
    
    '日本通運選択条件　(送り先CD)
    If GetIni(App.EXEName, "NITTSU_SELECT_OKURISAKI_CD", App.EXEName, c) Then
        c = "＊"
    End If
    wkVariant = Split(Trim(c), ",", -1)
    For i = 0 To UBound(wkVariant)
        ReDim Preserve NITTSU_SELECT_OKURISAKI_CD(0 To i)
    
        NITTSU_SELECT_OKURISAKI_CD(i) = wkVariant(i)
    
    Next i
    
    '第一貨物選択条件　(送り先集約CD)
    If GetIni(App.EXEName, "DAIICHI_SELECT_COL_OKURISAKI_CD", App.EXEName, c) Then
        c = "＊"
    End If
    wkVariant = Split(Trim(c), ",", -1)
    For i = 0 To UBound(wkVariant)
        ReDim Preserve DAIICHI_SELECT_COL_OKURISAKI_CD(0 To i)
    
        DAIICHI_SELECT_COL_OKURISAKI_CD(i) = wkVariant(i)
    
    Next i
    
    
    '取込みﾓｰﾄﾞ
    If GetIni(App.EXEName, "INPUT_MODE", App.EXEName, c) Then
        c = "0"
    End If
    If Trim(c) <> "0" And Trim(c) <> "1" Then
        c = "0"
    End If
    INPUT_MODE = Val(Trim(c))


    Select Case INPUT_MODE
    
        Case 0
            Text1(ptxSYUKA_YMD).Locked = True
            Text1(ptxINS_BIN).Locked = True
    
        Case 1
            
            Text1(ptxSYUKA_YMD).Locked = False
            Text1(ptxINS_BIN).Locked = False
            
            
            Text1(ptxSYUKA_YMD).Text = Format(Now, "YYYY/MM/DD")
            Text1(ptxSYUKA_YMD).SetFocus
            
            lblFILE_NAME.Visible = False
            Text1(ptxFILE_NAME).Visible = False
    End Select


    If INPUT_MODE = 1 Then
                                    '出荷予定ＯＰＥＮ
        If Y_SYU_Open(BtOpenNomal) Then
            Unload Me
        End If
                                    '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)ＯＰＥＮ
        If Y_SYU_H_Open(BtOpenNomal) Then
            Unload Me
        End If
    
    End If
    
    



    F1300101.Caption = F1300101.Caption & " " & LAST_UPDATE_DAY

End Sub


Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
    If lblFILE_NAME.Visible = False Then
        Exit Sub
    End If
    
    Text1(ptxFILE_NAME).Text = Trim(Data.Files(1))


    Command1(0).Value = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
Dim sts     As Integer
    
    If INPUT_MODE = 1 Then
                                                '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)ＣＬＯＳＥ
        sts = BTRV(BtOpClose, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
        If sts Then
            If sts <> BtErrNoOpen Then
                Call File_Error(sts, BtOpClose, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
            End If
        End If
                                                '出荷予定ＣＬＯＳＥ
        sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        If sts Then
            If sts <> BtErrNoOpen Then
                Call File_Error(sts, BtOpClose, "出荷予定")
            End If
        End If
    
        sts = BTRV(BtOpReset, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
        If sts Then
            Call File_Error(sts, BtOpReset, "")
        End If
    
    End If


    Set F1300101 = Nothing



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


Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    Call Tab_Ctrl(Shift)


End Sub

Private Sub Text1_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Text1(ptxFILE_NAME).Text = Trim(Data.Files(1))
    
    Command1(0).Value = True

End Sub





Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   「出荷予定データ」取込＆表示処理
'----------------------------------------------------------------------------




Dim sts             As Integer
Dim Ret             As String

Dim HS_SMEISAINo    As Long
Dim HS_SMEISAI_OP   As Boolean

Dim FileName        As String

Dim c               As String * 128

Dim i               As Integer

Dim Input_Buffer    As String
Dim Pos             As Integer
        
Dim Skip_Flg        As Boolean

Dim Input_Wk        As Variant

Dim SYUKA_NO        As String
Dim SYUKA_YMD       As String
Dim COL_OKURISAKI_CD _
                    As String
Dim OKURISAKI_CD    As String
Dim OKURISAKI       As String
Dim URIDEN          As String
Dim DEN_NO          As String
Dim Hinban          As String
Dim SURYO           As String
Dim CYU_NO          As String
Dim TOKUI_CODE      As String
Dim TOKUI_NAME      As String
Dim BIKOU           As String
Dim UNSOU           As String
Dim INS_BIN         As String
Dim JYUSHO          As String
Dim YUBIN_NO        As String
Dim TEL_NO          As String
Dim SEK_KEN_NO      As String
Dim SEK_HIN_NO      As String

Dim ans             As Integer
Dim Row             As Long


    List_Disp_Proc = True      '


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[送り状発行]データ取込開始", Me.hwnd, 0)

    Call Input_Lock


    '出荷明細ファイル名取り込み & ＯＰＥＮ
    FileName = Trim(Text1(ptxFILE_NAME).Text)
    HS_SMEISAI_OP = False
    
    On Error GoTo Exit_Proc
    
    HS_SMEISAINo = FreeFile
    Open FileName For Input As #HS_SMEISAINo
    HS_SMEISAI_OP = True
    
    
    
    
    Set SYUKA = Nothing
    
    Row = Min_Row - 1
    lblDisp_Count.Caption = ""
    
    
    

    Do While Not EOF(HS_SMEISAINo)
        
        
        
        
        Line Input #HS_SMEISAINo, Input_Buffer

        Input_Wk = Split(Input_Buffer, vbTab, -1)

        SYUKA_NO = ""
        SYUKA_YMD = ""
        COL_OKURISAKI_CD = ""
        OKURISAKI_CD = ""
        OKURISAKI = ""
        URIDEN = ""
        DEN_NO = ""
        Hinban = ""
        SURYO = ""
        CYU_NO = ""
        TOKUI_CODE = ""
        TOKUI_NAME = ""
        BIKOU = ""
        UNSOU = ""
        INS_BIN = ""
        JYUSHO = ""
        YUBIN_NO = ""
        TEL_NO = ""
        SEK_KEN_NO = ""
        SEK_HIN_NO = ""
        
        
        
        '出荷№
        If UBound(Input_Wk) > 0 Then
            SYUKA_NO = Input_Wk(1)
        End If
        
        
        If Not IsNumeric(SYUKA_NO) Then
        Else
            '出荷日
            If UBound(Input_Wk) > 1 Then
                
                If Mid(Format(Now, "YYYYMMDD"), 5, 2) = "12" Then
                    If Mid(CStr(Input_Wk(2)), 1, 2) = "01" Then
                        SYUKA_YMD = Format(CLng(Mid(Format(Now, "YYYYMMDD"), 1, 4) + 1), "0000") & "/" & Input_Wk(2)
                    Else
                        SYUKA_YMD = Mid(Format(Now, "YYYYMMDD"), 1, 4) & "/" & Input_Wk(2)
                    End If
                Else
                    SYUKA_YMD = Mid(Format(Now, "YYYYMMDD"), 1, 4) & "/" & Input_Wk(2)
                End If
            End If
            '集約送り先CD
            If UBound(Input_Wk) > 3 Then
                COL_OKURISAKI_CD = Trim(Input_Wk(4))
            
                If UBound(Input_Wk) > 4 Then
                    OKURISAKI_CD = Trim(Input_Wk(5))
            
                End If
            
            End If
            '送り先名
            If UBound(Input_Wk) > 6 Then
                If Trim(Input_Wk(7)) = "" Then
                Else
                    OKURISAKI = Trim(Input_Wk(7))
                End If
            End If
            
            
            '売伝
            If UBound(Input_Wk) > 7 Then
                URIDEN = Input_Wk(8)
            End If
            '伝票番号
            If UBound(Input_Wk) > 9 Then
                
                If Len(Input_Wk(10)) > 7 Then
                
                    DEN_NO = Left(Input_Wk(10), 7)
                Else
                    DEN_NO = Trim(Input_Wk(10))
                End If
            End If
            '品番
            If UBound(Input_Wk) > 11 Then
                Hinban = Trim(Input_Wk(12))
            End If
            '数量
            If UBound(Input_Wk) > 12 Then
                SURYO = Trim(Input_Wk(13))
            End If
            '注文№
            If UBound(Input_Wk) > 14 Then
                CYU_NO = Trim(Input_Wk(15))
            End If
            '得意先ｺｰﾄﾞ
            If UBound(Input_Wk) > 15 Then
                TOKUI_CODE = Trim(Input_Wk(16))
            End If
            '得意先名
            If UBound(Input_Wk) > 16 Then
                TOKUI_NAME = Trim(Input_Wk(17))
            End If
            '備考
            If UBound(Input_Wk) > 18 Then
                BIKOU = Trim(Input_Wk(19))
            End If
            '運送会社
            If UBound(Input_Wk) > 20 Then
                UNSOU = Trim(Input_Wk(21))
            End If
            '便 '2007.01.16
            If UBound(Input_Wk) > 21 Then
                INS_BIN = Trim(Input_Wk(22))
            End If
            
            
            '住所 '2009.11.19
            If UBound(Input_Wk) > 22 Then
                JYUSHO = Trim(Input_Wk(23))
            End If
            
            
            '郵便番号 '2010.04.05
            If UBound(Input_Wk) > 23 Then
                YUBIN_NO = Trim(Input_Wk(24))
            End If
            
            
            '電話番号 '2010.04.05
            If UBound(Input_Wk) > 24 Then
                TEL_NO = Trim(Input_Wk(25))
            End If
            
            
            '件管№　　　■管理№(上)   2011.04.30
            If UBound(Input_Wk) > 25 Then
                SEK_KEN_NO = Trim(Input_Wk(26))
            End If
            
            '品管№　　　■管理№(下)   2011.04.30
            If UBound(Input_Wk) > 26 Then
                SEK_HIN_NO = Trim(Input_Wk(27))
            End If
            
            'ｴﾗｰﾁｪｯｸ
            Skip_Flg = False
            
            If Trim(SYUKA_YMD) = "" Or _
                Trim(DEN_NO) = "" Or _
                Trim(Hinban) = "" Or _
                Trim(SURYO) = "" Then
                
                Skip_Flg = True
        
            Else
        
                If Not IsDate(SYUKA_YMD) Then
                    Skip_Flg = True
                Else
                    SYUKA_YMD = (Format(SYUKA_YMD, "YYYYMMDD"))
                End If
        
                If Not IsNumeric(SURYO) Then
                    Skip_Flg = True
                Else
                    If CLng(SURYO) = 0 Then
                        Skip_Flg = True
                    End If
                End If
        
        
            End If
        
            If Not Skip_Flg Then
                
                    
                    
                Row = Row + 1
                SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
                
                SYUKA(Row, colSYUKA_NO) = SYUKA_NO
                SYUKA(Row, colSYUKA_YMD) = SYUKA_YMD
                SYUKA(Row, colCOL_OKURISAKI_CD) = COL_OKURISAKI_CD
                SYUKA(Row, colOKURISAKI_CD) = OKURISAKI_CD
                SYUKA(Row, colOKURISAKI) = OKURISAKI
                SYUKA(Row, colURIDEN) = URIDEN
                SYUKA(Row, colDEN_NO) = DEN_NO
                SYUKA(Row, colHINBAN) = Hinban
                SYUKA(Row, colSURYO) = Format(Val(SURYO), "#0")
                SYUKA(Row, colCYU_NO) = CYU_NO
                SYUKA(Row, colTOKUI_CODE) = TOKUI_CODE
                SYUKA(Row, colTOKUI_NAME) = TOKUI_NAME
                SYUKA(Row, colBIKOU) = BIKOU
                SYUKA(Row, colUNSOU) = UNSOU
                SYUKA(Row, colINS_BIN) = INS_BIN
                SYUKA(Row, colJYUSHO) = JYUSHO
                SYUKA(Row, colYUBIN_NO) = YUBIN_NO
                SYUKA(Row, colTEL_NO) = TEL_NO
                SYUKA(Row, colSEK_KEN_NO) = SEK_KEN_NO
                SYUKA(Row, colSEK_HIN_NO) = SEK_HIN_NO
                    
                    
                If Row = 1 Then
                    Text1(ptxSYUKA_YMD) = SYUKA_YMD
                    Text1(ptxINS_BIN) = INS_BIN
                End If
                    
                    
            End If
        
        End If

    Loop

    Set TDBGrid1.Array = SYUKA
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst










    lblDisp_Count.Caption = Format(Row, "#0")

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[送り状発行]データ取込終了", Me.hwnd, 0)



    Close #HS_SMEISAINo
            
    Call Input_UnLock
            
            
    List_Disp_Proc = False      '
    
    
    Exit Function
    
Exit_Proc:
    
    
''''MsgBox Err.Number
    
    
    If HS_SMEISAI_OP Then
        Close #HS_SMEISAINo
    End If
    
    
    
    Select Case Err.Number
        
        '52 ファイル名または番号が不正です。
        '53 ファイルが見つかりません。
        '54 ファイル モードが不正です。
        '55 ファイルは既に開かれています。
        '57 デバイス I/O エラーです。
        '59 レコード長が一致しません。
        '61 ディスクの空き容量が不足しています。
        '62 ファイルにこれ以上データがありません。
        '63 レコード番号が不正です。
        '68 デバイスが準備されていません。
        '70 書き込みできません。
        '71 ディスクが準備されていません。
        '75 パス名が無効です。
        '76 パスが見つかりません。
        Case 52, 53, 54, 55, 57, 59, 61, 62, 63, 68, 70, 71, 75, 76
            
            
            MsgBox "指定のファイルが見つかりません。" & Chr(13) & Chr(10) & "正しいファイル名を入力してください。"
            
            
            
            List_Disp_Proc = False      '





            


        Case Else
            MsgBox Err.Description
    
    End Select
    
    Call Input_UnLock
    DoEvents
    
    On Error GoTo 0
    
    
End Function



Private Function List_Disp_Pervasive_Proc() As Integer
'----------------------------------------------------------------------------
'                   「出荷予定データ」取込＆表示処理
'----------------------------------------------------------------------------




Dim sts             As Integer
Dim Ret             As String
Dim com             As Integer

Dim c               As String * 128

Dim i               As Integer


Dim ans             As Integer
Dim Row             As Long


Dim SYUKA_NO        As String
Dim SYUKA_YMD       As String
Dim COL_OKURISAKI_CD _
                    As String
Dim OKURISAKI_CD    As String
Dim OKURISAKI       As String
Dim URIDEN          As String
Dim DEN_NO          As String
Dim Hinban          As String
Dim SURYO           As String
Dim CYU_NO          As String
Dim TOKUI_CODE      As String
Dim TOKUI_NAME      As String
Dim BIKOU           As String
Dim UNSOU           As String
Dim INS_BIN         As String
Dim JYUSHO          As String
Dim YUBIN_NO        As String
Dim TEL_NO          As String
Dim SEK_KEN_NO      As String
Dim SEK_HIN_NO      As String


Dim skip_f          As Integer


    List_Disp_Pervasive_Proc = True      '



    If Not IsDate(Text1(ptxSYUKA_YMD).Text) Then
        MsgBox "出荷日を正しく入力して下さい"
        Text1(ptxSYUKA_YMD).SetFocus
        List_Disp_Pervasive_Proc = False     '
        Exit Function
    End If

    If Trim(Text1(ptxINS_BIN).Text) = "" Then
        MsgBox "便№を正しく入力して下さい"
        Text1(ptxINS_BIN).SetFocus
        List_Disp_Pervasive_Proc = False      '
        Exit Function
    End If

    If Not IsNumeric(Text1(ptxINS_BIN).Text) Then
        MsgBox "便№を正しく入力して下さい"
        Text1(ptxINS_BIN).SetFocus
        List_Disp_Pervasive_Proc = False      '
        Exit Function
    End If



hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[送り状発行]データ取込開始", Me.hwnd, 0)

    Call Input_Lock


    
    
    
    
    Set SYUKA = Nothing
    
    Row = Min_Row - 1
    lblDisp_Count.Caption = ""
    
    
        
        
'    Call UniCode_Conv(K9_Y_SYU_H.SYUKA_YMD, Format(Text1(ptxSYUKA_YMD).Text, "YYYYMMDD"))
'    Call UniCode_Conv(K9_Y_SYU_H.INS_BIN, Format(Val(Text1(ptxINS_BIN).Text), "00"))
'    Call UniCode_Conv(K9_Y_SYU_H.SYUKA_NO, "")
        
    
    Call UniCode_Conv(K3_Y_SYU_H.SYUKA_YMD, Format(Text1(ptxSYUKA_YMD).Text, "YYYYMMDD"))
    
    com = BtOpGetGreaterEqual
    
    
    Do
    
    
'        sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K9_Y_SYU_H, Len(K9_Y_SYU_H), 9)
        sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K3_Y_SYU_H, Len(K3_Y_SYU_H), 3)

        Select Case sts
            Case BtNoErr
        
                If StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode) <> Format(Text1(ptxSYUKA_YMD).Text, "YYYYMMDD") Then
                    Exit Do
                End If
            
                skip_f = 0
                If StrConv(Y_SYU_HREC.INS_BIN, vbUnicode) <> Format(Val(Text1(ptxINS_BIN).Text), "00") Then
                    skip_f = 1
                End If
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                List_Disp_Pervasive_Proc = SYS_ERR
                Exit Function
        End Select
    
If skip_f = 0 Then
    
        SYUKA_NO = Format(Val(StrConv(Y_SYU_HREC.SYUKA_NO, vbUnicode)), "000")
        SYUKA_YMD = StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode)
        COL_OKURISAKI_CD = StrConv(Y_SYU_HREC.COL_OKURISAKI_CD, vbUnicode)
        OKURISAKI_CD = StrConv(Y_SYU_HREC.OKURISAKI_CD, vbUnicode)
        OKURISAKI = StrConv(Y_SYU_HREC.OKURISAKI, vbUnicode)
        URIDEN = StrConv(Y_SYU_HREC.URIDEN, vbUnicode)
        DEN_NO = StrConv(Y_SYU_HREC.DEN_NO, vbUnicode)
    
        
        Call UniCode_Conv(K0_Y_SYU.JGYOBU, StrConv(Y_SYU_HREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, StrConv(Y_SYU_HREC.ID_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)

        Select Case sts
            Case BtNoErr
                Hinban = StrConv(Y_SYUREC.HIN_NO, vbUnicode)
                SURYO = StrConv(Y_SYUREC.SURYO, vbUnicode)
            Case BtErrKeyNotFound
                Hinban = ""
                SURYO = "0"
            Case Else
                Call File_Error(sts, BtOpGetEqual, "出荷予定")
                Exit Function
        End Select
        CYU_NO = StrConv(Y_SYU_HREC.ODER_NO, vbUnicode)
        TOKUI_CODE = StrConv(Y_SYU_HREC.MUKE_CODE, vbUnicode)
        TOKUI_NAME = StrConv(Y_SYU_HREC.MUKE_NAME, vbUnicode)
        BIKOU = StrConv(Y_SYU_HREC.BIKOU, vbUnicode)
        UNSOU = StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)
        INS_BIN = StrConv(Y_SYU_HREC.INS_BIN, vbUnicode)
        JYUSHO = StrConv(Y_SYU_HREC.JYUSHO, vbUnicode)
        YUBIN_NO = StrConv(Y_SYU_HREC.YUBIN_NO, vbUnicode)
        TEL_NO = StrConv(Y_SYU_HREC.TEL_NO, vbUnicode)
        SEK_KEN_NO = StrConv(Y_SYU_HREC.SEK_KEN_NO, vbUnicode)
        SEK_HIN_NO = StrConv(Y_SYU_HREC.SEK_HIN_NO, vbUnicode)
    
        Row = Row + 1
        SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
        
        SYUKA(Row, colSYUKA_NO) = SYUKA_NO
        SYUKA(Row, colSYUKA_YMD) = SYUKA_YMD
        SYUKA(Row, colCOL_OKURISAKI_CD) = COL_OKURISAKI_CD
        SYUKA(Row, colOKURISAKI_CD) = OKURISAKI_CD
        SYUKA(Row, colOKURISAKI) = OKURISAKI
        
        If URIDEN = "1" Then
            SYUKA(Row, colURIDEN) = "有"
        Else
            SYUKA(Row, colURIDEN) = ""
        End If
        
        
        SYUKA(Row, colDEN_NO) = DEN_NO
        SYUKA(Row, colHINBAN) = Hinban
        SYUKA(Row, colSURYO) = Format(Val(SURYO), "#0")
        SYUKA(Row, colCYU_NO) = CYU_NO
        SYUKA(Row, colTOKUI_CODE) = TOKUI_CODE
        SYUKA(Row, colTOKUI_NAME) = TOKUI_NAME
        SYUKA(Row, colBIKOU) = BIKOU
        SYUKA(Row, colUNSOU) = UNSOU
        SYUKA(Row, colINS_BIN) = INS_BIN
        SYUKA(Row, colJYUSHO) = JYUSHO
        SYUKA(Row, colYUBIN_NO) = YUBIN_NO
        SYUKA(Row, colTEL_NO) = TEL_NO
        SYUKA(Row, colSEK_KEN_NO) = SEK_KEN_NO
        SYUKA(Row, colSEK_HIN_NO) = SEK_HIN_NO
                    

        SYUKA(Row, colID_NO) = StrConv(Y_SYU_HREC.ID_NO, vbUnicode)

End If
    
    
        com = BtOpGetNext
    Loop


    If Row > (Min_Row - 1) Then

        SYUKA.QuickSort Min_Row, SYUKA.UpperBound(1), colSYUKA_NO, 0, XTYPE_STRING
    
    End If
        
        
        


    Set TDBGrid1.Array = SYUKA
    
    
    
    
    
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst

    lblDisp_Count.Caption = Format(Row, "#0")

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[送り状発行]データ取込終了", Me.hwnd, 0)



            
    Call Input_UnLock
    DoEvents
            
            
    List_Disp_Pervasive_Proc = False      '
    
End Function



Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    F1300101.MousePointer = vbHourglass

    Call Ctrl_Lock(F1300101)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(F1300101)


    F1300101.MousePointer = vbDefault

End Sub









Private Function fukuyama_tyakuten() As Integer
'----------------------------------------------------------------------------
'                   福山　着店ｺｰﾄﾞ読込み
'----------------------------------------------------------------------------
Dim FileNo          As Long
Dim FileName        As String
Dim FILE_OP         As Boolean

Dim c               As String * 128

Dim Input_Buffer    As String
Dim Input_Wk        As Variant

Dim i               As Long

    fukuyama_tyakuten = True

    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[送り状発行]着店情報取込開始", Me.hwnd, 0)



    If GetIni(App.EXEName, "fukuyama-tyakuten", App.EXEName, c) Then
        Beep
        MsgBox "福山着店ｺｰﾄﾞ用CSVﾌｧｲﾙ名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    FileName = RTrim(c)

    FILE_OP = False

    On Error GoTo Exit_Proc
                            
    FileNo = FreeFile
    Open FileName For Input As #FileNo

    FILE_OP = True


    i = -1
    Do While Not EOF(FileNo)
        DoEvents
        
        Line Input #FileNo, Input_Buffer

        Input_Wk = Split(Input_Buffer, ",", -1)
    
        i = i + 1
        ReDim Preserve FUKUYAMA_TBL(0 To i)
            
        FUKUYAMA_TBL(i).YUBIN_NO = Input_Wk(0)
        FUKUYAMA_TBL(i).CODE = Input_Wk(1)
        FUKUYAMA_TBL(i).JYUSHO = Input_Wk(2)
        FUKUYAMA_TBL(i).TYAKUTEN = Input_Wk(3)
            
            
            
    Loop

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[送り状発行]着店情報取込終了", Me.hwnd, 0)

    Call Input_UnLock


    fukuyama_tyakuten = False
    Exit Function

Exit_Proc:
    
    
''''MsgBox Err.Number
    
    
    If FILE_OP Then
        Close #FileNo
    End If
    
    
    
    Select Case Err.Number
        
        '52 ファイル名または番号が不正です。
        '53 ファイルが見つかりません。
        '54 ファイル モードが不正です。
        '55 ファイルは既に開かれています。
        '57 デバイス I/O エラーです。
        '59 レコード長が一致しません。
        '61 ディスクの空き容量が不足しています。
        '62 ファイルにこれ以上データがありません。
        '63 レコード番号が不正です。
        '68 デバイスが準備されていません。
        '70 書き込みできません。
        '71 ディスクが準備されていません。
        '75 パス名が無効です。
        '76 パスが見つかりません。
        Case 52, 53, 54, 55, 57, 59, 61, 62, 63, 68, 70, 71, 75, 76
            MsgBox "指定のファイルが見つかりません。" & Chr(13) & Chr(10) & "福山　着店データ(fukuyama-tyakuten=)"

        Case Else
            MsgBox Err.Description
    
    '2011.12.03
    End Select
    
    Call Input_UnLock
    
    
    On Error GoTo 0
End Function

Private Function fukuyama_tyakuten_Set() As Integer
'----------------------------------------------------------------------------
'                   福山　着店ｺｰﾄﾞ　設定
'----------------------------------------------------------------------------
Dim i               As Long
Dim j               As Long

Dim wkTYAKUTEN      As String * 3

Dim sts             As Integer

    fukuyama_tyakuten_Set = True

   Call Input_Lock

Debug.Print Now

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[送り状発行]着店コード設定開始", Me.hwnd, 0)


    wkTYAKUTEN = ""
    For i = 1 To SYUKA.UpperBound(1)
        DoEvents
        
        For j = 0 To UBound(FUKUYAMA_TBL)
            
            DoEvents
            
'            If Trim(SYUKA(i, colJYUSHO)) = Trim(FUKUYAMA_TBL(j).JYUSHO) Then
'                wkTYAKUTEN = FUKUYAMA_TBL(j).TYAKUTEN
'                Exit For
'            End If
'
'            If Trim(Replace(SYUKA(i, colJYUSHO), "　", "")) = Trim(Replace(FUKUYAMA_TBL(j).JYUSHO, "　", "")) Then
'                wkTYAKUTEN = FUKUYAMA_TBL(j).TYAKUTEN
'                Exit For
'            End If
'
'            If Trim(Replace(SYUKA(i, colJYUSHO), " ", "")) = Trim(Replace(FUKUYAMA_TBL(j).JYUSHO, " ", "")) Then
'                wkTYAKUTEN = FUKUYAMA_TBL(j).TYAKUTEN
'                Exit For
'            End If
        
        
            If Trim(SYUKA(i, colYUBIN_NO)) = Trim(FUKUYAMA_TBL(j).YUBIN_NO) Then
         
                wkTYAKUTEN = FUKUYAMA_TBL(j).TYAKUTEN
                Exit For
            End If
        
        
        
        
        
        
        Next j
        SYUKA(i, colTYAKUTEN) = wkTYAKUTEN
    
    
        Call UniCode_Conv(K4_Y_SYU_H.ID_NO, SYUKA(i, colID_NO))
    
        sts = BTRV(BtOpGetEqual, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)

        Select Case sts
            Case BtNoErr
        
                Call UniCode_Conv(Y_SYU_HREC.TYAKUTEN, wkTYAKUTEN)
                sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
        
        
                Select Case sts
                
                    Case BtNoErr
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                        Exit Function
                
                End Select
        
        
        
            
            
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                Exit Function
        End Select
    
    
Debug.Print i
    
    Next i

    Set TDBGrid1.Array = SYUKA
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst


   Call Input_UnLock

    fukuyama_tyakuten_Set = False


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[送り状発行]着店コード設定終了", Me.hwnd, 0)


Debug.Print Now



End Function

Private Sub FUKUYAMA_CSV_OUT()
'----------------------------------------------------------------------------
'                   福山向け　送り状CSV処理
'----------------------------------------------------------------------------
Dim i                       As Integer
Dim j                       As Integer


Dim FileNo                  As Long
    
Dim svCOL_OKURISAKI_CD      As String
Dim svOKURISAKI_CD          As String
   
Dim Fsw                     As Boolean
   
   
   
   
   
   
   Call Input_Lock

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[送り状発行]福山通運用送り状データ出力開始", Me.hwnd, 0)


    FileNo = FreeFile
    Open (FUKUYAMA_CSV) For Output As FileNo



    'タイトル出力
    For i = 0 To UBound(TITLE_CSV) - 1
        Write #FileNo, TITLE_CSV(i),
    Next i
    Write #FileNo, TITLE_CSV(i)


    svCOL_OKURISAKI_CD = ""
    svOKURISAKI_CD = ""

    
    
    For i = 1 To SYUKA.UpperBound(1)
        If (Trim(svCOL_OKURISAKI_CD) = "" And Trim(svOKURISAKI_CD) = "") Then
                
            svCOL_OKURISAKI_CD = SYUKA(i, colCOL_OKURISAKI_CD)
            svOKURISAKI_CD = SYUKA(i, colOKURISAKI_CD)
                
            csvOKURISAKI = SYUKA(i, colOKURISAKI)
            csvTOKUI_NAME = SYUKA(i, colTOKUI_NAME)
            csvYUBIN_NO = SYUKA(i, colYUBIN_NO)
            csvJYUSHO = SYUKA(i, colJYUSHO)
            csvTEL_NO = SYUKA(i, colTEL_NO)
            csvTYAKUTEN = SYUKA(i, colTYAKUTEN)
            csvURIDEN = SYUKA(i, colURIDEN)
        
        
            Erase csvHinban_Tbl
            Fsw = True
                
        End If
    
        If Trim(svCOL_OKURISAKI_CD) <> Trim(SYUKA(i, colCOL_OKURISAKI_CD)) Or _
            Trim(svOKURISAKI_CD) <> Trim(SYUKA(i, colOKURISAKI_CD)) Then
    
            Call csv_make(FileNo)
    
    
    
            svCOL_OKURISAKI_CD = SYUKA(i, colCOL_OKURISAKI_CD)
            svOKURISAKI_CD = SYUKA(i, colOKURISAKI_CD)
                
            csvOKURISAKI = SYUKA(i, colOKURISAKI)
            csvTOKUI_NAME = SYUKA(i, colTOKUI_NAME)
            csvYUBIN_NO = SYUKA(i, colYUBIN_NO)
            csvJYUSHO = SYUKA(i, colJYUSHO)
            csvTEL_NO = SYUKA(i, colTEL_NO)
            csvTYAKUTEN = SYUKA(i, colTYAKUTEN)
            csvURIDEN = SYUKA(i, colURIDEN)
        
            Erase csvHinban_Tbl
            Fsw = True
    
    
        End If
    
        
        If Fsw Then
            ReDim Preserve csvHinban_Tbl(0 To 0)
            csvHinban_Tbl(0).DEN_NO = SYUKA(i, colDEN_NO)
            csvHinban_Tbl(0).Hinban = SYUKA(i, colHINBAN)
            csvHinban_Tbl(0).SURYO = 0
    
            Fsw = False
        
        
        End If
    
        For j = 0 To UBound(csvHinban_Tbl)
        
            If Trim(csvHinban_Tbl(j).Hinban) = Trim(SYUKA(i, colHINBAN)) Then
                csvHinban_Tbl(j).SURYO = csvHinban_Tbl(j).SURYO + Val(SYUKA(i, colSURYO))
                Exit For
            End If
        
        Next j
    
        If j > UBound(csvHinban_Tbl) Then
            ReDim Preserve csvHinban_Tbl(0 To j)
            csvHinban_Tbl(j).DEN_NO = SYUKA(i, colDEN_NO)
            csvHinban_Tbl(j).Hinban = SYUKA(i, colHINBAN)
            csvHinban_Tbl(j).SURYO = Val(SYUKA(i, colSURYO))
        End If
    
    Next i


    If Trim(svCOL_OKURISAKI_CD) <> "" Or Trim(svOKURISAKI_CD) <> "" Then
        Call csv_make(FileNo)
    End If



    Close #FileNo




        

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[送り状発行]福山通運用送り状データ出力終了", Me.hwnd, 0)


   Call Input_UnLock



End Sub
Private Sub OKURI_CSV_OUT(unsou_SELECT As Integer)
'----------------------------------------------------------------------------
'                   送り状CSV出力
'----------------------------------------------------------------------------
Dim i                       As Integer
Dim j                       As Integer


Dim FileNo                  As Long
Dim FileName               As String
    
Dim svCOL_OKURISAKI_CD      As String
Dim svOKURISAKI_CD          As String
   
Dim Fsw                     As Boolean
Dim Onsw                    As Boolean
   
Dim UNSOU_NAME              As String
   
   
   
   Call Input_Lock


    Select Case unsou_SELECT
        Case 4
            UNSOU_NAME = "久留米運輸"
            FileName = KURUME_CSV
        Case 5
            UNSOU_NAME = "第一貨物"
            FileName = DAIICHI_CSV
        Case 6
            UNSOU_NAME = "日本通運"
            FileName = NITTSU_CSV
    End Select



    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[送り状発行]" & UNSOU_NAME & "用送り状データ出力開始", Me.hwnd, 0)


    FileNo = FreeFile
    Open (FileName) For Output As FileNo



    'タイトル出力
    For i = 0 To UBound(TITLE_CSV) - 1
        Write #FileNo, TITLE_CSV(i),
    Next i
    Write #FileNo, TITLE_CSV(i)


    svCOL_OKURISAKI_CD = ""
    svOKURISAKI_CD = ""

    
    
    For i = 1 To SYUKA.UpperBound(1)
        
        Onsw = False
        
        Select Case unsou_SELECT
        
            Case 4                  '久留米運輸
                
                For j = 0 To UBound(KURUME_SELECT)
                
                    If InStr(1, SYUKA(i, colJYUSHO), KURUME_SELECT(j)) Then
                        Onsw = True
                        Exit For
                    End If
                                    
                Next j
                
                For j = 0 To UBound(NITTSU_SELECT_COL_OKURISAKI_CD)
                
                    If Trim(NITTSU_SELECT_COL_OKURISAKI_CD(j)) = Trim(SYUKA(i, colCOL_OKURISAKI_CD)) Then
                        Onsw = False
                        Exit For
                    End If
                Next j
                
                For j = 0 To UBound(NITTSU_SELECT_COL_OKURISAKI_CD)
                
                    If Trim(NITTSU_SELECT_OKURISAKI_CD(j)) = Trim(SYUKA(i, colOKURISAKI_CD)) Then
                        Onsw = False
                        Exit For
                    End If
                Next j
                
                For j = 0 To UBound(DAIICHI_SELECT_COL_OKURISAKI_CD)
                
                    If Trim(DAIICHI_SELECT_COL_OKURISAKI_CD(j)) = Trim(SYUKA(i, colCOL_OKURISAKI_CD)) Then
                        Onsw = False
                        Exit For
                    End If
                Next j
            Case 5                  '第一貨物
        
        
                For j = 0 To UBound(DAIICHI_SELECT_COL_OKURISAKI_CD)
                
                    If Trim(DAIICHI_SELECT_COL_OKURISAKI_CD(j)) = Trim(SYUKA(i, colCOL_OKURISAKI_CD)) Then
                        Onsw = True
                        Exit For
                    End If
                
                Next j
        
        
            Case 6                  '日本運輸
            
            
                For j = 0 To UBound(NITTSU_SELECT_COL_OKURISAKI_CD)
                
                    If Trim(NITTSU_SELECT_COL_OKURISAKI_CD(j)) = Trim(SYUKA(i, colCOL_OKURISAKI_CD)) Then
                        Onsw = True
                        Exit For
                    End If
                
                Next j
            
                For j = 0 To UBound(NITTSU_SELECT_OKURISAKI_CD)
                
                    If Trim(NITTSU_SELECT_OKURISAKI_CD(j)) = Trim(SYUKA(i, colOKURISAKI_CD)) Then
                        Onsw = True
                        Exit For
                    End If
                
                Next j
            
            
                For j = 0 To UBound(DAIICHI_SELECT_COL_OKURISAKI_CD)
                
                    If Trim(DAIICHI_SELECT_COL_OKURISAKI_CD(j)) = Trim(SYUKA(i, colCOL_OKURISAKI_CD)) Then
                        Onsw = False
                        Exit For
                    End If
                
                Next j
            
            
            
        
        End Select
        
        
        If Onsw Then
        
            If (Trim(svCOL_OKURISAKI_CD) = "" And Trim(svOKURISAKI_CD) = "") Then
                    
                svCOL_OKURISAKI_CD = SYUKA(i, colCOL_OKURISAKI_CD)
                svOKURISAKI_CD = SYUKA(i, colOKURISAKI_CD)
                    
                csvOKURISAKI = SYUKA(i, colOKURISAKI)
                csvTOKUI_NAME = SYUKA(i, colTOKUI_NAME)
                csvYUBIN_NO = SYUKA(i, colYUBIN_NO)
                csvJYUSHO = SYUKA(i, colJYUSHO)
                csvTEL_NO = SYUKA(i, colTEL_NO)
                csvTYAKUTEN = ""
            
                Erase csvHinban_Tbl
                Fsw = True
                    
            End If
        
            If Trim(svCOL_OKURISAKI_CD) <> Trim(SYUKA(i, colCOL_OKURISAKI_CD)) Or _
                Trim(svOKURISAKI_CD) <> Trim(SYUKA(i, colOKURISAKI_CD)) Then
        
                Call csv_make(FileNo)
        
        
        
                svCOL_OKURISAKI_CD = SYUKA(i, colCOL_OKURISAKI_CD)
                svOKURISAKI_CD = SYUKA(i, colOKURISAKI_CD)
                    
                csvOKURISAKI = SYUKA(i, colOKURISAKI)
                csvTOKUI_NAME = SYUKA(i, colTOKUI_NAME)
                csvYUBIN_NO = SYUKA(i, colYUBIN_NO)
                csvJYUSHO = SYUKA(i, colJYUSHO)
                csvTEL_NO = SYUKA(i, colTEL_NO)
                csvTYAKUTEN = ""
            
                Erase csvHinban_Tbl
                Fsw = True
        
        
            End If
        
            
            If Fsw Then
                ReDim Preserve csvHinban_Tbl(0 To 0)
                csvHinban_Tbl(0).DEN_NO = SYUKA(i, colDEN_NO)
                csvHinban_Tbl(0).Hinban = SYUKA(i, colHINBAN)
                csvHinban_Tbl(0).SURYO = 0
        
                Fsw = False
            
            
            End If
        
            For j = 0 To UBound(csvHinban_Tbl)
            
                If Trim(csvHinban_Tbl(j).Hinban) = Trim(SYUKA(i, colHINBAN)) Then
                    csvHinban_Tbl(j).SURYO = csvHinban_Tbl(j).SURYO + Val(SYUKA(i, colSURYO))
                    Exit For
                End If
            
            Next j
        
            If j > UBound(csvHinban_Tbl) Then
                ReDim Preserve csvHinban_Tbl(0 To j)
                csvHinban_Tbl(j).DEN_NO = SYUKA(i, colDEN_NO)
                csvHinban_Tbl(j).Hinban = SYUKA(i, colHINBAN)
                csvHinban_Tbl(j).SURYO = Val(SYUKA(i, colSURYO))
            End If
        End If
    Next i


    If Trim(svCOL_OKURISAKI_CD) <> "" Then
        Call csv_make(FileNo)
    End If



    Close #FileNo




        

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[送り状発行]" & UNSOU_NAME & "用送り状データ出力終了", Me.hwnd, 0)


   Call Input_UnLock
End Sub

Private Sub csv_make(FileNo As Long)
'----------------------------------------------------------------------------
'                   送り状CSV出力
'----------------------------------------------------------------------------
Dim i           As Integer

Dim wkSuryo     As String



    Write #FileNo, , , RTrim(csvOKURISAKI), ,
'    Write #FileNo, csvTOKUI_NAME,
    Write #FileNo, csvYUBIN_NO,
    Write #FileNo, RTrim(csvJYUSHO),
    Write #FileNo, , , RTrim(csvTEL_NO),
    Write #FileNo, ,

    If UBound(csvHinban_Tbl) > 4 Then
        Write #FileNo, "別明細有り", , , , ,
    Else
        
        
        
        
        For i = 0 To UBound(csvHinban_Tbl)
        
        
            wkSuryo = Format(csvHinban_Tbl(i).SURYO, "#0")
            If wkSuryo < 2 Then
                wkSuryo = " " & wkSuryo
            End If
        
        
        
        
            If i = 0 And csvURIDEN = "有" Then
                Write #FileNo, csvHinban_Tbl(i).DEN_NO & " " & csvHinban_Tbl(i).Hinban & " " & wkSuryo & " " & "●",
            Else
                Write #FileNo, csvHinban_Tbl(i).DEN_NO & " " & csvHinban_Tbl(i).Hinban & " " & wkSuryo,
            End If
        
        Next i
    
    End If
    
    If UBound(csvHinban_Tbl) = 0 Then
        Write #FileNo, , , , ,
    End If


    If UBound(csvHinban_Tbl) = 1 Then
        Write #FileNo, , , ,
    End If
    
    If UBound(csvHinban_Tbl) = 2 Then
        Write #FileNo, , ,
    End If
    
    If UBound(csvHinban_Tbl) = 3 Then
        Write #FileNo, ,
    End If
    
    
    
    Write #FileNo, , , , , , , , , , , , , , , , , , , , ,


    Write #FileNo, Trim(csvTYAKUTEN)

End Sub



