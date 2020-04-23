VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form F1030211 
   BackColor       =   &H00FFFFFF&
   Caption         =   "伝票番号指定出庫表印刷"
   ClientHeight    =   11670
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   16830
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
   ScaleHeight     =   11670
   ScaleWidth      =   16830
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Command1 
      Caption         =   "全て選択"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10440
      TabIndex        =   35
      Top             =   840
      Width           =   1380
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   6780
      MaxLength       =   10
      TabIndex        =   10
      Top             =   600
      Width           =   1275
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   1530
      MaxLength       =   4
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   2250
      MaxLength       =   2
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   2730
      MaxLength       =   2
      TabIndex        =   6
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   3810
      MaxLength       =   4
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   4530
      MaxLength       =   2
      TabIndex        =   8
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   5010
      MaxLength       =   2
      TabIndex        =   9
      Top             =   600
      Width           =   375
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   2
      Left            =   8100
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   3
      Top             =   120
      Width           =   3360
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   6780
      MaxLength       =   8
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      Left            =   4410
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   1575
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   120
      Width           =   1095
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
      Left            =   10440
      TabIndex        =   23
      Top             =   10560
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
      Left            =   9600
      TabIndex        =   22
      Top             =   10560
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
      Left            =   8760
      TabIndex        =   21
      Top             =   10560
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "印　刷"
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
      TabIndex        =   20
      Top             =   10560
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "検 索"
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
      TabIndex        =   19
      Top             =   10560
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
      Left            =   5760
      TabIndex        =   18
      Top             =   10560
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
      Left            =   4920
      TabIndex        =   17
      Top             =   10560
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
      Left            =   4080
      TabIndex        =   16
      Top             =   10560
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
      Left            =   2760
      TabIndex        =   15
      Top             =   10560
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
      Left            =   1920
      TabIndex        =   14
      Top             =   10560
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
      Left            =   1080
      TabIndex        =   13
      Top             =   10560
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
      Left            =   240
      TabIndex        =   12
      Top             =   10560
      Width           =   855
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   8895
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   16260
      _ExtentX        =   28681
      _ExtentY        =   15690
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   4
      Columns(1)._MaxComboItems=   5
      Columns(1).ValueItems(0)._DefaultItem=   0
      Columns(1).ValueItems(0).Value=   ""
      Columns(1).ValueItems(0).Value.vt=   8
      Columns(1).ValueItems(0).DisplayValue=   "1"
      Columns(1).ValueItems(0).DisplayValue.vt=   8
      Columns(1).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(1).ValueItems.Count=   1
      Columns(1).Caption=   "選"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "注文区分（ｺｰﾄﾞ）"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "注区"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "向け先ｺｰﾄﾞ"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "出荷先"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "伝票日付"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "ID№"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "伝票№"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "品番(外部)"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "出荷数"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "出荷済数"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "印"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "品番（内部）"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "品名"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "標準棚番"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   16
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   714
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=16"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4366"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
      Splits(0)._ColumnProps(4)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=635"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=503"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1614"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1482"
      Splits(0)._ColumnProps(14)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=1217"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1085"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=4366"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=4233"
      Splits(0)._ColumnProps(23)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(25)=   "Column(5).Width=4048"
      Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=3916"
      Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(29)=   "Column(6).Width=2249"
      Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=2117"
      Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(33)=   "Column(7).Width=2672"
      Splits(0)._ColumnProps(34)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(7)._WidthInPix=2540"
      Splits(0)._ColumnProps(36)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(37)=   "Column(8).Width=2355"
      Splits(0)._ColumnProps(38)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(8)._WidthInPix=2223"
      Splits(0)._ColumnProps(40)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(41)=   "Column(9).Width=2672"
      Splits(0)._ColumnProps(42)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(9)._WidthInPix=2540"
      Splits(0)._ColumnProps(44)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(45)=   "Column(10).Width=2249"
      Splits(0)._ColumnProps(46)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(10)._WidthInPix=2117"
      Splits(0)._ColumnProps(48)=   "Column(10)._ColStyle=2"
      Splits(0)._ColumnProps(49)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(50)=   "Column(11).Width=2249"
      Splits(0)._ColumnProps(51)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(11)._WidthInPix=2117"
      Splits(0)._ColumnProps(53)=   "Column(11)._ColStyle=2"
      Splits(0)._ColumnProps(54)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(55)=   "Column(12).Width=979"
      Splits(0)._ColumnProps(56)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(12)._WidthInPix=847"
      Splits(0)._ColumnProps(58)=   "Column(12)._ColStyle=1"
      Splits(0)._ColumnProps(59)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(60)=   "Column(13).Width=2672"
      Splits(0)._ColumnProps(61)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(13)._WidthInPix=2540"
      Splits(0)._ColumnProps(63)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(64)=   "Column(14).Width=4366"
      Splits(0)._ColumnProps(65)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(66)=   "Column(14)._WidthInPix=4233"
      Splits(0)._ColumnProps(67)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(68)=   "Column(15).Width=4366"
      Splits(0)._ColumnProps(69)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(70)=   "Column(15)._WidthInPix=4233"
      Splits(0)._ColumnProps(71)=   "Column(15).Order=16"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bgcolor=&HFFFF&,.bold=0,.fontsize=1200"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=35"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=31,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ＭＳ ゴシック"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35,.bgcolor=&H80FF80&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36,.bgcolor=&HFF80&"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=45,.parent=2,.namedParent=47"
      _StyleDefs(23)  =   "FilterBarStyle:id=48,.parent=1,.namedParent=50"
      _StyleDefs(24)  =   "Splits(0).Style:id=111,.parent=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=120,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=112,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=113,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=114,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=116,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=115,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=117,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=118,.parent=9,.bgcolor=&HFFFF&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=119,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=46,.parent=45"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=49,.parent=48"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=44,.parent=111"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=41,.parent=112"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=42,.parent=113"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=43,.parent=115"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=128,.parent=111,.alignment=2,.locked=0"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=125,.parent=112,.alignment=3"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=126,.parent=113,.alignment=3"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=127,.parent=115"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=132,.parent=111"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=129,.parent=112"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=130,.parent=113"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=131,.parent=115"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=136,.parent=111"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=133,.parent=112"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=134,.parent=113"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=135,.parent=115"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=140,.parent=111"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=137,.parent=112"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=138,.parent=113"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=139,.parent=115"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=144,.parent=111"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=141,.parent=112"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=142,.parent=113"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=143,.parent=115"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=148,.parent=111"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=145,.parent=112"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=146,.parent=113"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=147,.parent=115"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=152,.parent=111"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=149,.parent=112"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=150,.parent=113"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=151,.parent=115"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=156,.parent=111"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=153,.parent=112"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=154,.parent=113"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=155,.parent=115"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=160,.parent=111"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=157,.parent=112"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=158,.parent=113"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=159,.parent=115"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=164,.parent=111,.alignment=1,.locked=0"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=161,.parent=112,.alignment=3"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=162,.parent=113,.alignment=3"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=163,.parent=115"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=168,.parent=111,.alignment=1,.locked=0"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=165,.parent=112,.alignment=3"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=166,.parent=113,.alignment=3"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=167,.parent=115"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=172,.parent=111,.alignment=2,.locked=0"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=169,.parent=112,.alignment=3"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=170,.parent=113,.alignment=3"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=171,.parent=115"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=176,.parent=111"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=173,.parent=112"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=174,.parent=113"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=175,.parent=115"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=180,.parent=111"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=177,.parent=112"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=178,.parent=113"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=179,.parent=115"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=14,.parent=111"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=11,.parent=112"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=12,.parent=113"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=13,.parent=115"
      _StyleDefs(100) =   "Named:id=29:Normal"
      _StyleDefs(101) =   ":id=29,.parent=0,.bgcolor=&HFF00&"
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
      _StyleDefs(114) =   ":id=35,.parent=29,.bgcolor=&HFFFF&"
      _StyleDefs(115) =   "Named:id=36:OddRow"
      _StyleDefs(116) =   ":id=36,.parent=29"
      _StyleDefs(117) =   "Named:id=47:RecordSelector"
      _StyleDefs(118) =   ":id=47,.parent=30"
      _StyleDefs(119) =   "Named:id=50:FilterBar"
      _StyleDefs(120) =   ":id=50,.parent=29"
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "　印刷件数"
      Height          =   255
      Index           =   9
      Left            =   12840
      TabIndex        =   42
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "読込み件数"
      Height          =   255
      Index           =   8
      Left            =   12840
      TabIndex        =   41
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Height          =   255
      Index           =   1
      Left            =   9720
      TabIndex        =   40
      Top             =   11160
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label4 
      Height          =   255
      Index           =   0
      Left            =   7440
      TabIndex        =   39
      Top             =   11160
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   14265
      TabIndex        =   38
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   14265
      TabIndex        =   37
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9840
      TabIndex        =   36
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "伝票番号"
      Height          =   255
      Index           =   11
      Left            =   5745
      TabIndex        =   34
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "出荷予定日"
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   33
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   1
      Left            =   2130
      TabIndex        =   32
      Top             =   720
      Width           =   120
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   2
      Left            =   2610
      TabIndex        =   31
      Top             =   720
      Width           =   120
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   240
      Index           =   5
      Left            =   3450
      TabIndex        =   30
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   6
      Left            =   4410
      TabIndex        =   29
      Top             =   720
      Width           =   120
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   7
      Left            =   4890
      TabIndex        =   28
      Top             =   720
      Width           =   120
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "出荷先"
      Height          =   255
      Index           =   10
      Left            =   5985
      TabIndex        =   27
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "注文区分"
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   26
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "印刷区分"
      Height          =   255
      Index           =   3
      Left            =   210
      TabIndex        =   25
      Top             =   240
      Width           =   975
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
      TabIndex        =   24
      Top             =   11040
      Width           =   180
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
Attribute VB_Name = "F1030211"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxMUKE_CODE% = 0             '向け先コード（手入力）

Private Const ptxS_DEN_DT_YY% = 1           '開始　出荷予定日　年
Private Const ptxS_DEN_DT_MM% = 2           '開始　出荷予定日　月
Private Const ptxS_DEN_DT_DD% = 3           '開始　出荷予定日　日
Private Const ptxE_DEN_DT_YY% = 4           '終了　出荷予定日　年
Private Const ptxE_DEN_DT_MM% = 5           '終了　出荷予定日　月
Private Const ptxE_DEN_DT_DD% = 6           '終了　出荷予定日　日
Private Const ptxDEN_NO% = 7                '伝票番号

Private Const Text_Max% = 7                 '画面項目別最大ｲﾝﾃﾞｯｸｽ

Private Const pcmbPRINT_KBN% = 0            '印刷区分
Private Const pcmbCyu_Kbn% = 1              '注文区分
Private Const pcmbMUKE_Code% = 2            '向け先


Dim SYUKA As New XArrayDB

Private Const Min_Row% = 1              '最小行数

Dim Max_Row     As Long                 'グリッド最大表示件数

Dim SYUKA_DATA  As String               '出荷データフルパス


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 15             '最大列数           '2014.01.07 14->15



Private Const ColDummy% = 0             'ダミー

Private Const ColSEL% = 1               '選
Private Const ColCyu_Kbn% = 2           '注文区分名称
Private Const ColCyu_Kbn_Name% = 3      '注文区分名称
Private Const ColMUKE_Code% = 4         '出荷先ｺｰﾄﾞ（非表示）
Private Const ColMUKE_Name% = 5         '出荷先名
Private Const ColDEN_DT% = 6            '伝票日付
Private Const ColID_NO% = 7             '伝票ＩＤ
Private Const ColDEN_NO% = 8            '伝票№
Private Const ColHIN_GAI% = 9           '品目（外部）
Private Const ColSURYO% = 10            '出荷数（予定）
Private Const ColJITU_SURYO% = 11       '出荷数（実績）
Private Const ColPrint% = 12            '出庫表印刷マーク
Private Const ColHIN_NAI% = 13          '品目（内部）
Private Const ColHIN_Name% = 14         '品名

Private Const ColST_TANABAN% = 15       '標準棚番       2014.01.07



Private Sort_Tbl(Min_Col To Max_Col) _
                As Integer                  'ｿｰﾄの制御 0:昇順 1:降順    2013.03.26


Private Const Print_KBN0$ = "新規　"
Private Const Print_KBN1$ = "再印刷"
Private Const Print_KBN_SIN$ = "0"
Private Const Print_KBN_SAI$ = "1"

Private KASO_NYUKA_SOKO As String * 2       '仮想　入荷倉庫番号
Private KASO_SYOHN_SOKO As String * 2       '仮想　商品化倉庫番号
Private KASO_NAI_SOKO As String * 2         '仮想　内職倉庫番号


Private Const LMAX% = 56                    '頁内最大行数
Private Const MGN_L% = 10                   '左余白（桁数：１から）
Private Const MGN_U% = 1                    '上余白（行数：１から）

Dim Pdate As String                         '印刷開始日付（ﾍｯﾀﾞｰ用）
Dim Ptime As String                         '印刷開始時刻（ﾍｯﾀﾞｰ用）


Dim NormalFont As New StdFont               '印刷フォント
Dim Code39Font As New StdFont               '印刷フォント

Dim NON_MUKE_CODE() As String * 8           '除外向け先コード
Dim NON_MUKE_FLG    As Boolean

Dim ALL_Check       As Boolean              '全件対象

Dim Print_Cnt       As Long




'Private Const Last_Update_Day$ = "(F103021 2014.01.08 10:00)"
'Private Const Last_Update_Day$ = "(F103021 2019.06.05 15:45)"
Private Const Last_Update_Day$ = "(F103021 2019.06.05 16:42)"


Private Sub Combo_Click(Index As Integer)
    Select Case Index
        Case pcmbCyu_Kbn
            
            
            Text(ptxMUKE_CODE).SetFocus
    End Select

End Sub

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Select Case Index
        Case pcmbCyu_Kbn
            Text(ptxMUKE_CODE).SetFocus
        Case pcmbMUKE_Code
            
            Text(ptxMUKE_CODE).Text = Trim(Right(Combo(Index).Text, 16))
            
            
            
            
            If List_Disp_Proc Then
                Unload Me
            End If
    End Select

End Sub


Private Sub Command_Click(Index As Integer)

Dim ans As Integer

    Select Case Index
        
        Case 7                              '検索
            If List_Disp_Proc() Then
                Unload Me
            End If
        
                    
'            ALL_Check = False               '2013.03.26
'            Command1.Caption = "全て選択"   '2013.03.26
        
        
        
        Case 8                              '印刷
            
            
            
            ans = MsgBox("「出庫表」印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                
                TDBGrid1.Update
                
                
                If Print_Proc(1) Then
                    Unload Me
                End If
            
                            
                            
                ALL_Check = False
                If List_Disp_Proc() Then
                    Unload Me
                End If
                
                Command1.Caption = "全て選択"   '2013.03.26
            
            
            End If
        
        
        Case 11                             '終了
            Unload Me
        Case Else
            Beep
    End Select
    
End Sub

Private Sub Command1_Click()
    
    If Not ALL_Check Then
        ALL_Check = True
        Command1.Caption = "全て解除"
    
    Else
        ALL_Check = False
        Command1.Caption = "全て選択"
    
    End If

    If List_Disp_Proc() Then
        Unload Me
    End If

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

                    '最大表示件数の獲得
    If GetIni(App.EXEName, "LISTMAX", "SYS", c) Then
        Beep
        MsgBox "最大表示件数の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    Max_Row = CLng(RTrim(c))
                                
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
            F1030211.Caption = "伝票番号指定出庫表印刷（" + RTrim(JGYOBU_T(i).NAME) + ")" & Last_Update_Day
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)
                                '入荷仮想倉庫取り込み
    If GetIni(App.EXEName, "KASO_NYUKA_SOKO", "SYS", c) Then
        c = ""
    End If
    KASO_NYUKA_SOKO = RTrim(c)
                                '商品化仮想倉庫取り込み
    If GetIni(App.EXEName, "KASO_SYOHN_SOKO", "SYS", c) Then
        c = ""
    End If
    KASO_SYOHN_SOKO = RTrim(c)
                                '内職仮想倉庫取り込み
    If GetIni(App.EXEName, "KASO_NAI_SOKO", "SYS", c) Then
        c = ""
    End If
    KASO_NAI_SOKO = RTrim(c)
                                
                                '除外向け先コード獲得
    i = 0
    NON_MUKE_FLG = False
    Do
        If GetIni(App.EXEName, "MUKE" & Format(i + 1, "00"), "SYS", c) Then
            Exit Do
        End If
    
        If RTrim(c) = "NON" Then
            Exit Do
        End If
    
        ReDim Preserve NON_MUKE_CODE(0 To i)
    
        NON_MUKE_CODE(i) = RTrim(c)
        NON_MUKE_FLG = True
    
        i = i + 1
    Loop

                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '向け先管理マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定ＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫ＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '印刷フォント設定
    With NormalFont
        .NAME = F1030211.FontName
        .Size = 10
    End With
                                '印刷フォント設定（バーコード）
    With Code39Font
        .NAME = Label1.FontName
        .Size = Label1.FontSize
    End With

    ALL_Check = False

'向け先設定
    If MTS_Set_Proc() Then
        Unload Me
    End If

                                '画面初期設定
    Combo(pcmbPRINT_KBN).AddItem "      " & "   " & " "
    Combo(pcmbPRINT_KBN).AddItem Print_KBN0 & "   " & Print_KBN_SIN
    Combo(pcmbPRINT_KBN).AddItem Print_KBN1 & "   " & Print_KBN_SAI
    Combo(pcmbPRINT_KBN).ListIndex = 0

'ｺﾝﾎﾞ初期設定
    
    Combo(pcmbCyu_Kbn).AddItem "全て" & "   " & " "
    Combo(pcmbCyu_Kbn).AddItem CYU_KBN_1 & "   " & CYU_KBN_TUK
    Combo(pcmbCyu_Kbn).AddItem CYU_KBN_2 & "   " & CYU_KBN_SPO
    Combo(pcmbCyu_Kbn).AddItem CYU_KBN_3 & "   " & CYU_KBN_HJU
    Combo(pcmbCyu_Kbn).AddItem CYU_KBN_4 & "   " & CYU_KBN_TOK
    Combo(pcmbCyu_Kbn).AddItem CYU_KBN_E & "   " & CYU_KBN_BOU
    Combo(pcmbCyu_Kbn).AddItem CYU_KBN_T & "   " & CYU_KBN_KIN
    Combo(pcmbCyu_Kbn).ListIndex = 0

    Combo(pcmbPRINT_KBN).SetFocus



'出荷予定日初期表示     2012.06.15
    Text(ptxS_DEN_DT_YY).Text = Mid(Format(Now, "YYYYMMDD"), 1, 4)
    Text(ptxS_DEN_DT_MM).Text = Mid(Format(Now, "YYYYMMDD"), 5, 2)
    Text(ptxS_DEN_DT_DD).Text = Mid(Format(Now, "YYYYMMDD"), 7, 2)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫マスタ")
        End If
    End If
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
                                            '出荷予定ＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定")
        End If
    End If
                                            '在庫ＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定")
        End If
    End If
    
    sts = BTRV(BtOpReset, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
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
    F1030211.Caption = "伝票番号指定出庫表印刷（" + RTrim(JGYOBU_T(Index).NAME) + ")" & Last_Update_Day
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
    
    
    Combo(pcmbMUKE_Code).Clear
    
    Edit = "全出荷先" & "   "
    Edit = Edit & "                "
    Combo(pcmbMUKE_Code).AddItem Edit
    
    
    
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
        
        Edit = StrConv(MTSREC.MUKE_NAME, vbUnicode) & "   "
        Edit = Edit & StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode)
        
        
        Combo(pcmbMUKE_Code).AddItem Edit
    
        com = BtOpGetNext
    
    Loop

    If Combo(pcmbMUKE_Code).ListCount <= 0 Then
    Else
        Combo(pcmbMUKE_Code).ListIndex = 0
    End If

    Call Input_UnLock

    MTS_Set_Proc = False
End Function


Private Function List_Disp_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim Row         As Long
    
Dim Skip_Flg    As Boolean
    
    
    
Dim wkDEN_No    As String
    
    
    
Label4(0).Caption = Format(Now, "HHMMSS")
    
    F1030211.MousePointer = vbHourglass
    
    
    List_Disp_Proc = True
                                    
'    Call Input_Lock
                                    
    For i = ptxS_DEN_DT_YY To ptxE_DEN_DT_DD
    
        If IsNumeric(Trim(Text(i).Text)) Then
        
        
            Text(i).Text = Right(Format(CInt(Text(i).Text), "0000"), Text(i).MaxLength)
        
        End If
    
    
    Next i
                                    
    Text(ptxMUKE_CODE).Text = Trim(Right(Combo(pcmbMUKE_Code).Text, 16))
                                    
    If Trim(Text(ptxMUKE_CODE).Text) = "" Then
        Call UniCode_Conv(MTSREC.MUKE_CODE, "")
        Call UniCode_Conv(MTSREC.SS_CODE, "")
    Else
        Call UniCode_Conv(K2_MTS.MUKE_CODE, Text(ptxMUKE_CODE).Text)
        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K2_MTS, Len(K2_MTS), 2)
        Select Case sts
            Case BtNoErr
                            
            Case BtErrKeyNotFound
                            
                Call UniCode_Conv(K3_MTS.SS_CODE, Text(ptxMUKE_CODE).Text)
                                                    
                sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K3_MTS, Len(K3_MTS), 3)
                Select Case sts
                    Case BtNoErr
                                    
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(MTSREC.MUKE_CODE, "")
                        Call UniCode_Conv(MTSREC.SS_CODE, "")
                                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                        Unload Me
                End Select

            Case Else
                Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                Unload Me
        End Select
    End If

    For i = 0 To Combo(pcmbMUKE_Code).ListCount - 1 '向け先

        If Right(Combo(pcmbMUKE_Code).List(i), 16) = StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode) Then
            Combo(pcmbMUKE_Code).ListIndex = i
            Exit For
        End If
    

    Next
                                    
    '空読み
'    sts = BTRV(BtOpGetFirst, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K2_Y_SYU, Len(K2_Y_SYU), 2)
'
'    Select Case sts
'        Case BtNoErr
'            Skip_Flg = False
'        Case BtErrEOF
'            Skip_Flg = True
'        Case Else
'            Unload Me
'    End Select
                                    
                                        'テーブルリセット
    Set SYUKA = Nothing
    
    'ｿｰﾄ情報の初期化                2013.03.26
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0             'ﾃﾞﾌｫﾙﾄ昇順
    Next i


    

    
    
                                    '出荷予定読み込み開始
    Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU) '事業部
                                                    '注文区分
    Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, "")
                                                    '向け先
    
    
    Row = Min_Row - 1
        
        
    
    com = BtOpGetGreaterEqual
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>  2012.06.15
    
    Call UniCode_Conv(K10_Y_SYU.KAN_KBN, KAN_KBN_UN)
    Call UniCode_Conv(K10_Y_SYU.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K10_Y_SYU.KEY_SYUKA_YMD, Text(ptxS_DEN_DT_YY).Text & Text(ptxS_DEN_DT_MM).Text & Text(ptxS_DEN_DT_DD).Text)
    Call UniCode_Conv(K10_Y_SYU.PRINT_YMD, "")
    Call UniCode_Conv(K10_Y_SYU.DEN_NO, "")



'>>>>>>>>>>>>>>>>>>>>>>>>>  2012.06.15
    
    
    
    Do
'2012.06.15        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K10_Y_SYU, Len(K10_Y_SYU), 10)      '2012.06.15
    
        Select Case sts
            Case BtNoErr
        
            Case BtErrEOF
                Exit Do
            Case Else
                Unload Me
        End Select
                                
                                
                                
                                '完了区分 KEYﾌﾞﾚｰｸ  2012.06.15
        If StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = KAN_KBN_FIN Then
            
'Call LOG_OUT(LOG_F, "1=" & StrConv(Y_SYUREC.ID_NO, vbUnicode))
            Exit Do
        End If
                                
                                
                                
                                '事業部 KEYﾌﾞﾚｰｸ
        If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
            
'Call LOG_OUT(LOG_F, "1=" & StrConv(Y_SYUREC.ID_NO, vbUnicode))
            Exit Do
        End If
        
        Skip_Flg = False
                                
                                '注文区分 KEYﾌﾞﾚｰｸ
        If Right(Combo(pcmbCyu_Kbn).Text, 1) <> " " Then
            If StrConv(Y_SYUREC.CYU_KBN, vbUnicode) <> Right(Combo(pcmbCyu_Kbn).Text, 1) Then
                Skip_Flg = True
'Call LOG_OUT(LOG_F, "2=" & StrConv(Y_SYUREC.ID_NO, vbUnicode))
            End If
        End If
                                
                                '向け先 KEYﾌﾞﾚｰｸ
        If Trim(Text(ptxMUKE_CODE).Text) <> "" Then
            If Trim(Right(Combo(pcmbMUKE_Code).Text, 16)) <> "" Then
                If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) <> Trim(Left(Right(Combo(pcmbMUKE_Code).Text, 16), 8)) Or _
                    Trim(StrConv(Y_SYUREC.SS_CODE, vbUnicode)) <> Trim(Right(Combo(pcmbMUKE_Code).Text, 8)) Then
                    Skip_Flg = True
'Call LOG_OUT(LOG_F, "3=" & StrConv(Y_SYUREC.ID_NO, vbUnicode))
                End If
            End If
        
        Else
            If NON_MUKE_FLG Then
                For i = 0 To UBound(NON_MUKE_CODE)
                    If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = Trim(NON_MUKE_CODE(i)) Then
                        Skip_Flg = True
'Call LOG_OUT(LOG_F, "4=" & StrConv(Y_SYUREC.ID_NO, vbUnicode))
                        Exit For
                    End If
                Next i
            End If
        End If
            
        
        
                                '処理完了済
        If CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) = CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) Then
            Skip_Flg = True
'Call LOG_OUT(LOG_F, "5=" & StrConv(Y_SYUREC.ID_NO, vbUnicode))
        End If
                                
                                
                                '印刷区分
        If Trim(Right(Combo(pcmbPRINT_KBN).Text, 1)) <> "" Then
            If Right(Combo(pcmbPRINT_KBN).Text, 1) = Print_KBN_SIN Then
                If IsDate(Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 1, 4) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 5, 2) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 7, 2)) Then
                    Skip_Flg = True
'Call LOG_OUT(LOG_F, "6=" & StrConv(Y_SYUREC.ID_NO, vbUnicode) & "-" & StrConv(Y_SYUREC.PRINT_YMD, vbUnicode))
                End If
            Else
                If Not IsDate(Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 1, 4) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 5, 2) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 7, 2)) Then
                    Skip_Flg = True
                    Skip_Flg = True
'Call LOG_OUT(LOG_F, "7=" & StrConv(Y_SYUREC.ID_NO, vbUnicode))
                End If
            End If
        End If
        
                                '伝票日付範囲(開始)
        If StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) < (Text(ptxS_DEN_DT_YY).Text & Text(ptxS_DEN_DT_MM).Text & Text(ptxS_DEN_DT_DD).Text) Then
            Skip_Flg = True
'Call LOG_OUT(LOG_F, "8=" & StrConv(Y_SYUREC.ID_NO, vbUnicode))
        End If
                                '伝票日付範囲(終了)
        If Trim(Text(ptxE_DEN_DT_YY).Text & Text(ptxE_DEN_DT_MM).Text & Text(ptxE_DEN_DT_DD).Text) <> "" Then
            If StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) > (Text(ptxE_DEN_DT_YY).Text & Text(ptxE_DEN_DT_MM).Text & Text(ptxE_DEN_DT_DD).Text) Then
                Skip_Flg = True
'Call LOG_OUT(LOG_F, "9=" & StrConv(Y_SYUREC.ID_NO, vbUnicode))
            End If
        End If
                                '伝票番号
        If Trim(Text(ptxDEN_NO).Text) <> "" Then
            If Trim(StrConv(Y_SYUREC.DEN_NO, vbUnicode)) <> Trim(Text(ptxDEN_NO)) Then
                Skip_Flg = True
'Call LOG_OUT(LOG_F, "10=" & StrConv(Y_SYUREC.ID_NO, vbUnicode))
            End If
        Else
'''--->伝票№桁数制限を廃止
'''            If IsNumeric(Trim(StrConv(Y_SYUREC.DEN_NO, vbUnicode))) Then
'''                wkDEN_No = Trim(Format(CLng(StrConv(Y_SYUREC.DEN_NO, vbUnicode))))
'''            Else
'''                wkDEN_No = Trim(StrConv(Y_SYUREC.DEN_NO, vbUnicode))
'''            End If
'''            If Len(wkDEN_No) > 5 Then
'''                Skip_Flg = True
'''            End If
        
        End If
        
        If Skip_Flg Then
        Else
            Row = Row + 1
            If Row > Max_Row Then
                Beep
                MsgBox "最大表示行数を超えました。"
                Exit Do
            End If
                    
            
            
            If Grid_Set_Proc(Row) Then
                Unload Me
            End If
        End If
        
        com = BtOpGetNext
        
        DoEvents
    Loop
                                
                                
    If Row = (Min_Row - 1) Then
                                'データなし
        Command1.Enabled = False
        ALL_Check = False
    Else
                                'DBテーブルリンク   2013.03.26 初期の並び順を品番順に変更
'        SYUKA.QuickSort Min_Row, (SYUKA.UpperBound(1)), ColCyu_Kbn, XORDER_ASCEND, XTYPE_STRING, _
'                                                            ColMUKE_Code, XORDER_ASCEND, XTYPE_STRING
    
    
        SYUKA.QuickSort Min_Row, (SYUKA.UpperBound(1)), ColHIN_GAI, XORDER_ASCEND, XTYPE_STRING, _
                                                            ColMUKE_Code, XORDER_ASCEND, XTYPE_STRING
    
    
        
'        SYUKA.ReDim Min_Row, Row + 1, Min_Col, Max_Col             2013.03.26
'        SYUKA(Row + 1, ColDummy) = "--------------------------"    2013.03.26
        
        Command1.Enabled = True
    
    
    End If
    
    
    
    TDBGrid1.Style.Locked = True
    
    
    
Label2.Caption = Row
    
    
    Set TDBGrid1.Array = SYUKA
    
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    
    TDBGrid1.MoveFirst
    
'    Call Input_UnLock
    F1030211.MousePointer = vbDefault
    
'    Combo(pcmbMUKE_Code).SetFocus
    
    List_Disp_Proc = False

Label4(1).Caption = Format(Now, "HHMMSS")
    
End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1030211.MousePointer = vbHourglass

    Call Ctrl_Lock(F1030211)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1030211)


    F1030211.MousePointer = vbDefault

End Sub

Private Function Grid_Set_Proc(Row As Long) As Integer

Dim sts As Integer

    
    Grid_Set_Proc = True

    

    SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
                                                                
    SYUKA(Row, ColSEL) = ALL_Check                              '選択
                                                                
                                                                '注文区分
    SYUKA(Row, ColCyu_Kbn) = StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
    
    
    Select Case StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
        Case CYU_KBN_TUK    '月切
            SYUKA(Row, ColCyu_Kbn_Name) = CYU_KBN_1
        Case CYU_KBN_SPO    'スポット(緊急)
            SYUKA(Row, ColCyu_Kbn_Name) = CYU_KBN_2
        Case CYU_KBN_HJU    '補充
            SYUKA(Row, ColCyu_Kbn_Name) = CYU_KBN_3
        Case CYU_KBN_TOK    '特売(一斉出荷)
            SYUKA(Row, ColCyu_Kbn_Name) = CYU_KBN_4
        Case CYU_KBN_BOU    '貿易
            SYUKA(Row, ColCyu_Kbn_Name) = CYU_KBN_E
    End Select
                                                                    
                                                                    '出荷先ｺｰﾄﾞ
'   SYUKA(Row, ColMUKE_Code) = StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode)            '2013.03.26
    SYUKA(Row, ColMUKE_Code) = StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & StrConv(Y_SYUREC.SS_CODE, vbUnicode)        '2013.03.26
                                                                    '出荷先名称
    Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))
    Call UniCode_Conv(K0_MTS.SS_CODE, StrConv(Y_SYUREC.SS_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)

    Select Case sts
        Case BtNoErr
            SYUKA(Row, ColMUKE_Name) = StrConv(MTSREC.MUKE_NAME, vbUnicode)
        Case BtErrKeyNotFound
'            SYUKA(Row, ColMUKE_Name) = StrConv(MTSREC.MUKE_CODE, vbUnicode)                                        '2013.03.26
            SYUKA(Row, ColMUKE_Name) = StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)                                       '2013.03.26

        Case Else
            Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
            Exit Function
    End Select
                                                                    '伝票日付
    SYUKA(Row, ColDEN_DT) = Left(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 4) & "/" & Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 5, 2) & "/" & Right(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 2)
    SYUKA(Row, ColID_NO) = StrConv(Y_SYUREC.ID_NO, vbUnicode)       'ＩＤ№
    SYUKA(Row, ColDEN_NO) = StrConv(Y_SYUREC.DEN_NO, vbUnicode)     '伝票№
    SYUKA(Row, ColHIN_GAI) = StrConv(Y_SYUREC.ITEM_NO, vbUnicode)
    SYUKA(Row, ColHIN_GAI) = StrConv(Y_SYUREC.HIN_NO, vbUnicode)    '品番（外部）
                                                                    '出荷数（予定）
    SYUKA(Row, ColSURYO) = Format(CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)), "#0")
                                                                    '出荷数（実績）
    SYUKA(Row, ColJITU_SURYO) = Format(CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)), "#0")
                                                                    '印刷区分
    If IsDate(Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 1, 4) & "/" & _
        Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 5, 2) & "/" & _
        Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 7, 2)) Then
        SYUKA(Row, ColPrint) = "○"
    Else
        SYUKA(Row, ColPrint) = ""
    End If
    
    SYUKA(Row, ColHIN_NAI) = StrConv(Y_SYUREC.HIN_NAI, vbUnicode)   '品番（内部）
                                                                    '品目マスタ読込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            SYUKA(Row, ColHIN_Name) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        Case BtErrKeyNotFound
            SYUKA(Row, ColHIN_Name) = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    End Select
    
    '標準棚番   2014.01.07
    SYUKA(Row, ColST_TANABAN) = StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                StrConv(ITEMREC.ST_DAN, vbUnicode)
    '標準棚番   2014.01.07
    
    
    
    Grid_Set_Proc = False
End Function

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    
'------------------------------------------------------ 2013.03.26
    If SYUKA.Count(1) <= 0 Then
        Exit Sub
    End If
    
    
    
    If Sort_Tbl(ColIndex) = 0 Then
        Sort_Tbl(ColIndex) = 1
    Else
        If Sort_Tbl(ColIndex) = 1 Then
            Sort_Tbl(ColIndex) = 0
        End If
    
    End If
    
    If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then
                    
        SYUKA.QuickSort Min_Row, SYUKA.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
        
        Set TDBGrid1.Array = SYUKA
        
        TDBGrid1.ReBind
        TDBGrid1.Update
        TDBGrid1.MoveFirst


    End If


End Sub

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
            
            If Trim(Text(Index).Text) = "" Then
                Call UniCode_Conv(MTSREC.MUKE_CODE, "")
                Call UniCode_Conv(MTSREC.SS_CODE, "")
            Else
                
                Text(Index).Text = StrConv(Text(Index).Text, vbUpperCase)
                
                
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
            End If

            For i = 0 To Combo(pcmbMUKE_Code).ListCount - 1 '向け先
    
                If Right(Combo(pcmbMUKE_Code).List(i), 16) = StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode) Then
                    Combo(pcmbMUKE_Code).ListIndex = i
                    Exit For
                End If
            
    
            Next

    End Select

    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i

End Sub
Private Function Print_Proc(Mode As Integer) As Integer

Dim Lcnt            As Integer


Dim SAVE_Cyu_Kbn    As String * 1
Dim SAVE_MUKE_CODE  As String * 16
Dim PRI_HIN_GAI     As String * 13
Dim Betu_LOCATION   As String * 8

Dim com             As Integer
Dim sts             As Integer
Dim ans             As Integer
    

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim ZAIKO_QTY       As Long
Dim TEMP_QTY        As Long
Dim RetBuf          As String
    
Dim RePrint         As Boolean
    
Dim HTANABAN        As String * 8
    
    
Dim i               As Long         '2013.12.25
    
    Print_Proc = True

    Call Input_Lock
    
    
    '標準棚番順に並び変え   2014.01.07
    SYUKA.QuickSort Min_Row, SYUKA.UpperBound(1), ColST_TANABAN, Sort_Tbl(ColST_TANABAN), XTYPE_STRING

    Set TDBGrid1.Array = SYUKA

    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    '標準棚番順に並び変え   2014.01.07
    
    
    
Print_Cnt = 0
    
    
    
    Lcnt = 99
    
    Set Printer.Font = NormalFont
    
    Printer.Orientation = vbPRORLandscape
    Pdate = Date
    Ptime = Time




    Select Case Mode            '2013.12.25
        Case 0                  '2013.12.25


            com = BtOpGetGreaterEqual
            
            Do
                DoEvents
                                                    '出荷予定データ読み込み
                sts = Y_Syu_Get(RePrint, com)
                Select Case sts
                    Case BtNoErr
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Exit Function
                End Select
                                                    
                If Lcnt = 99 Then
                    SAVE_Cyu_Kbn = StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
                    SAVE_MUKE_CODE = StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & StrConv(Y_SYUREC.SS_CODE, vbUnicode)
                Else
                                                    '注文区分のブレーク
                    If SAVE_Cyu_Kbn <> StrConv(Y_SYUREC.CYU_KBN, vbUnicode) Then
                        Lcnt = LMAX + 1
                        SAVE_Cyu_Kbn = StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
                    End If
                                                    '向け先のブレーク
        '2008.11.17            If SAVE_MUKE_CODE <> StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & StrConv(Y_SYUREC.SS_CODE, vbUnicode) Then
        '2008.11.17                Lcnt = LMAX + 1
        '2008.11.17                SAVE_MUKE_CODE = StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & StrConv(Y_SYUREC.SS_CODE, vbUnicode)
        '2008.11.17            End If
                End If
        
                If Lcnt > LMAX Then                 'ヘッダーコントロール
                    If Head_Proc(SAVE_Cyu_Kbn, Lcnt) Then
                        Exit Function
                    End If
                    PRI_HIN_GAI = ""
                End If
                                                    
                '-----------------------------------------------------  '１行目
                If StrConv(Y_SYUREC.HIN_NO, vbUnicode) <> PRI_HIN_GAI Then
                    PRI_HIN_GAI = StrConv(Y_SYUREC.HIN_NO, vbUnicode)
                                                    '明細印刷
                                                    
                                                    
                    Printer.Print Tab(MGN_L - 5);
                    If RePrint Then
                        Printer.Print "再";
                    End If
                                                    
                    Printer.Print Tab(MGN_L);
                                                    
                    If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & StrConv(Y_SYUREC.SS_CODE, vbUnicode)) = "S8" Then
                    
        '                If S8_LOCATION_Proc("S8", HTANABAN) Then
        '                    Exit Function
        '                End If
                    
                                                        '標準棚番
        '                Printer.Print Mid(HTANABAN, 1, 2) & "-";
        '                Printer.Print Mid(HTANABAN, 3, 2) & "-";
        '                Printer.Print Mid(HTANABAN, 5, 2) & "-";
        '                Printer.Print Mid(HTANABAN, 7, 2);
                    
                                                        '標準棚番
                        Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2) & "-";
                        Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 3, 2) & "-";
                        Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 5, 2) & "-";
                        Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 7, 2);
                    
                    
                        HTANABAN = StrConv(Y_SYUREC.HTANABAN, vbUnicode)
                    
                    
                    
                    Else
                                                        '標準棚番
                        Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2) & "-";
                        Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 3, 2) & "-";
                        Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 5, 2) & "-";
                        Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 7, 2);
                    
                    
                        HTANABAN = StrConv(Y_SYUREC.HTANABAN, vbUnicode)
                    
                    End If
        
                    Printer.Print Tab(MGN_L + 13);                          '2008.11.17
                    Printer.Print StrConv(Y_SYUREC.MUKE_CODE, vbUnicode);   '2008.11.17
                    
        
        
        
        
                    Printer.Print Tab(MGN_L + 23);  '2008.11.17 13-->23
                                                    '品番(外)
                    Printer.Print Left(StrConv(Y_SYUREC.HIN_NO, vbUnicode), 13);
        
                    Printer.Print Tab(MGN_L + 37);
                                                    '標準棚　在庫数
                    If Len(Trim(HTANABAN)) = 0 Then
                        SUMI_QTY = 0
                        MI_QTY = 0
                    Else
                        If Zaiko_Syukei_Proc(SUMI_QTY, _
                                                MI_QTY, _
                                                Last_JGYOBU, _
                                                StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                                StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                                HTANABAN) Then
                            Exit Function
                        End If
                    End If
                               
                    ZAIKO_QTY = SUMI_QTY + MI_QTY
                    RetBuf = Format(ZAIKO_QTY, "#,##0")
                    
                    If Len(RetBuf) < 9 Then
                        RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                    End If
                    Printer.Print RetBuf;
                                                    
                    If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = "S8" Then
                        If Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2) = "S8" Then
                                                    '別置棚検索
                            If Tana_Kensaku(Betu_LOCATION) Then
                                Print_Proc = True
                                Exit Function
                            End If
                        
                        
                        Else
                                                    
                            If S8_LOCATION_Proc("S8", Betu_LOCATION) Then
                                Exit Function
                            Else
                                If Trim(Betu_LOCATION) = "" Then
                                    If Tana_Kensaku(Betu_LOCATION) Then
                                        Print_Proc = True
                                        Exit Function
                                    End If
                                End If
                            End If
                        
                        
                        
                        
                        End If
                    Else
                                                    '別置棚検索
                        If Tana_Kensaku(Betu_LOCATION) Then
                            Print_Proc = True
                            Exit Function
                        End If
                    
                    End If
                    
                    
                    SUMI_QTY = 0
                    MI_QTY = 0
                    
                    If Len(Trim(Betu_LOCATION)) = 0 Then
                    Else
                                                    '別置棚　在庫数
                        Printer.Print Tab(MGN_L + 48);
                        Printer.Print Left(Betu_LOCATION, 2) & "-" _
                                        & Mid(Betu_LOCATION, 3, 2) & "-" _
                                        & Mid(Betu_LOCATION, 5, 2) & "-" _
                                        & Right(Betu_LOCATION, 2);
                        
                        If Zaiko_Syukei_Proc(SUMI_QTY, _
                                                MI_QTY, _
                                                Last_JGYOBU, _
                                                StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                                StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                                Betu_LOCATION) Then
                            Exit Function
                        End If
                    End If
                    
                    Printer.Print Tab(MGN_L + 59);
                    ZAIKO_QTY = SUMI_QTY + MI_QTY
                    RetBuf = Format(ZAIKO_QTY, "#,##0")
                    If Len(RetBuf) < 9 Then
                        RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                    End If
                    Printer.Print RetBuf;
                                                    '商品化＆内職在庫数
                    Printer.Print Tab(MGN_L + 68);
                    If Zaiko_Syukei_Proc(SUMI_QTY, _
                                            MI_QTY, _
                                            Last_JGYOBU, _
                                            StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                            StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                            KASO_SYOHN_SOKO & "01" & "01" & "01") Then
                        Exit Function
                    End If
                    TEMP_QTY = SUMI_QTY + MI_QTY
                    If Zaiko_Syukei_Proc(SUMI_QTY, _
                                            MI_QTY, _
                                            Last_JGYOBU, _
                                            StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                            StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                            KASO_NAI_SOKO & "01" & "01" & "01") Then
                        Exit Function
                    End If
                    ZAIKO_QTY = TEMP_QTY + SUMI_QTY + MI_QTY
                    RetBuf = Format(ZAIKO_QTY, "#,##0")
                    If Len(RetBuf) < 9 Then
                        RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                    End If
                    Printer.Print RetBuf;
                                                    
                                                    '入荷倉庫在庫
                    Printer.Print Tab(MGN_L + 77);
                    If Zaiko_Syukei_Proc(SUMI_QTY, _
                                            MI_QTY, _
                                            Last_JGYOBU, _
                                            StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                            StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                            KASO_NYUKA_SOKO & "01" & "01" & "01") Then
                        Exit Function
                    End If
                                
                    ZAIKO_QTY = SUMI_QTY + MI_QTY
                    RetBuf = Format(ZAIKO_QTY, "#,##0")
                    If Len(RetBuf) < 9 Then
                        RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                    End If
                    Printer.Print RetBuf;
                End If
                
                '2003.06.03（注文区分）
        '        Printer.Print Tab(MGN_L + 76);
        '        Select Case StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
        '            Case CYU_KBN_SPO
        '                Printer.Print " 緊";
        '            Case CYU_KBN_HJU
        '                Printer.Print " 補";
        '            Case Else
        '                Printer.Print " 　";
        '        End Select
                '2003.06.03
                            
                                                    '伝票№
                Printer.Print Tab(MGN_L + 90);
'                Printer.Print Left(StrConv(Y_SYUREC.DEN_NO, vbUnicode), 6);
                '2019.06.05 上6桁から１０桁に変更。
                Printer.Print StrConv(Y_SYUREC.DEN_NO, vbUnicode);
                
                                    '2019.06.05 ↓+4
                Printer.Print Tab(MGN_L + 100 + 4);
                TEMP_QTY = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode) - CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)))
                RetBuf = Format(TEMP_QTY, "#,##0")
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Printer.Print RetBuf;
                                
                                  '2019.06.05 ↓+4
                Printer.Print Tab(MGN_L + 120 + 4);
                                                        '印刷フォント設定（Ｃｏｄｅ３９）
                Set Printer.Font = Code39Font
                                    'バーコード(*伝票ID*)
                Printer.Print "*" & StrConv(Y_SYUREC.ID_NO, vbUnicode) & "*";
                                                        '印刷フォント設定（通常）
                Set Printer.Font = NormalFont
                
                '-----------------------------------------------------  '２行目
                Printer.Print Tab(MGN_L + 90);
                Printer.Print StrConv(Y_SYUREC.ID_NO, vbUnicode);
        
                Printer.Print
                Printer.Print
                
                Lcnt = Lcnt + 3
        
                                                        '印刷日付設定更新
        '        If Right(Combo(pcmbPRINT_KBN).Text, 1) = Print_KBN_SIN Then
        '            Call UniCode_Conv(Y_SYUREC.PRINT_YMD, Format(Now, "YYYYMMDD"))
        '
        '            Do
        '
        '                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
        '                Select Case sts
        '                    Case BtNoErr
        '                        Exit Do
        '                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
        '
        '                        Beep
        '                        ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
        '                        If ans = vbCancel Then
        '                            Print_Proc = SYS_CANCEL
        '                            Exit Function
        '                        End If
        '                    Case Else
        '                        Call File_Error(sts, BtOpUpdate, "出荷予定")
        '                        Print_Proc = SYS_ERR
        '                        Exit Function
        '
        '                End Select
        '
        '
        '            Loop
        '        End If
                
         Print_Cnt = Print_Cnt + 1
                com = BtOpGetNext
                
            Loop



    '-------------------------------    2013.12.25
        Case 1
        
            SAVE_Cyu_Kbn = ""
            For i = 1 To SYUKA.Count(1)
                
                
                DoEvents
                                                    
                If SYUKA(i, ColSEL) Then
                                                    '出荷予定データ読み込み
                
                    Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU)
                    Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, SYUKA(i, ColID_NO))
                    
                    
                    sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            MsgBox "データ内容が変更されています。再表示して下さい"
                            Exit For
                        Case Else
                            Call File_Error(sts, BtOpUpdate, "出荷予定")
                            Print_Proc = SYS_ERR
                            Exit Function
                    End Select
            
                    If Lcnt > LMAX Then                 'ヘッダーコントロール
                        If Head_Proc(SAVE_Cyu_Kbn, Lcnt) Then
                            Exit Function
                        End If
                        PRI_HIN_GAI = ""
                    End If
                                                        
                    '-----------------------------------------------------  '１行目
                    PRI_HIN_GAI = ""
                    If StrConv(Y_SYUREC.HIN_NO, vbUnicode) <> PRI_HIN_GAI Then
                        PRI_HIN_GAI = StrConv(Y_SYUREC.HIN_NO, vbUnicode)
                                                        '明細印刷
                                                        
                                                        
                        Printer.Print Tab(MGN_L - 5);
                        
                        
                        If Trim(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode)) <> "" Then
                            RePrint = True
                        Else
                            RePrint = False
                        End If
                        
                        
                        If RePrint Then
                            Printer.Print "再";
                        End If
                                                        
                        Printer.Print Tab(MGN_L);
                                                        
                        If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & StrConv(Y_SYUREC.SS_CODE, vbUnicode)) = "S8" Then
                        
            '                If S8_LOCATION_Proc("S8", HTANABAN) Then
            '                    Exit Function
            '                End If
                        
                                                            '標準棚番
            '                Printer.Print Mid(HTANABAN, 1, 2) & "-";
            '                Printer.Print Mid(HTANABAN, 3, 2) & "-";
            '                Printer.Print Mid(HTANABAN, 5, 2) & "-";
            '                Printer.Print Mid(HTANABAN, 7, 2);
                        
                                                            '標準棚番
                            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2) & "-";
                            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 3, 2) & "-";
                            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 5, 2) & "-";
                            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 7, 2);
                        
                        
                            HTANABAN = StrConv(Y_SYUREC.HTANABAN, vbUnicode)
                        
                        
                        
                        Else
                                                            '標準棚番
                            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2) & "-";
                            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 3, 2) & "-";
                            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 5, 2) & "-";
                            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 7, 2);
                        
                        
                            HTANABAN = StrConv(Y_SYUREC.HTANABAN, vbUnicode)
                        
                        End If
            
                        Printer.Print Tab(MGN_L + 13);                          '2008.11.17
                        Printer.Print StrConv(Y_SYUREC.MUKE_CODE, vbUnicode);   '2008.11.17
                        
            
            
            
            
                        Printer.Print Tab(MGN_L + 23);  '2008.11.17 13-->23
                                                        '品番(外)
                        Printer.Print Left(StrConv(Y_SYUREC.HIN_NO, vbUnicode), 13);
            
                        Printer.Print Tab(MGN_L + 37);
                                                        '標準棚　在庫数
                        If Len(Trim(HTANABAN)) = 0 Then
                            SUMI_QTY = 0
                            MI_QTY = 0
                        Else
                            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                                    MI_QTY, _
                                                    Last_JGYOBU, _
                                                    StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                                    StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                                    HTANABAN) Then
                                Exit Function
                            End If
                        End If
                                   
                        ZAIKO_QTY = SUMI_QTY + MI_QTY
                        RetBuf = Format(ZAIKO_QTY, "#,##0")
                        
                        If Len(RetBuf) < 9 Then
                            RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                        End If
                        Printer.Print RetBuf;
                                                        
                        If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = "S8" Then
                            If Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2) = "S8" Then
                                                        '別置棚検索
                                If Tana_Kensaku(Betu_LOCATION) Then
                                    Print_Proc = True
                                    Exit Function
                                End If
                            
                            
                            Else
                                                        
                                If S8_LOCATION_Proc("S8", Betu_LOCATION) Then
                                    Exit Function
                                Else
                                    If Trim(Betu_LOCATION) = "" Then
                                        If Tana_Kensaku(Betu_LOCATION) Then
                                            Print_Proc = True
                                            Exit Function
                                        End If
                                    End If
                                End If
                            
                            
                            
                            
                            End If
                        Else
                                                        '別置棚検索
                            If Tana_Kensaku(Betu_LOCATION) Then
                                Print_Proc = True
                                Exit Function
                            End If
                        
                        End If
                        
                        
                        SUMI_QTY = 0
                        MI_QTY = 0
                        
                        If Len(Trim(Betu_LOCATION)) = 0 Then
                        Else
                                                        '別置棚　在庫数
                            Printer.Print Tab(MGN_L + 48);
                            Printer.Print Left(Betu_LOCATION, 2) & "-" _
                                            & Mid(Betu_LOCATION, 3, 2) & "-" _
                                            & Mid(Betu_LOCATION, 5, 2) & "-" _
                                            & Right(Betu_LOCATION, 2);
                            
                            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                                    MI_QTY, _
                                                    Last_JGYOBU, _
                                                    StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                                    StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                                    Betu_LOCATION) Then
                                Exit Function
                            End If
                        End If
                        
                        Printer.Print Tab(MGN_L + 59);
                        ZAIKO_QTY = SUMI_QTY + MI_QTY
                        RetBuf = Format(ZAIKO_QTY, "#,##0")
                        If Len(RetBuf) < 9 Then
                            RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                        End If
                        Printer.Print RetBuf;
                                                        '商品化＆内職在庫数
                        Printer.Print Tab(MGN_L + 68);
                        If Zaiko_Syukei_Proc(SUMI_QTY, _
                                                MI_QTY, _
                                                Last_JGYOBU, _
                                                StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                                StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                                KASO_SYOHN_SOKO & "01" & "01" & "01") Then
                            Exit Function
                        End If
                        TEMP_QTY = SUMI_QTY + MI_QTY
                        If Zaiko_Syukei_Proc(SUMI_QTY, _
                                                MI_QTY, _
                                                Last_JGYOBU, _
                                                StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                                StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                                KASO_NAI_SOKO & "01" & "01" & "01") Then
                            Exit Function
                        End If
                        ZAIKO_QTY = TEMP_QTY + SUMI_QTY + MI_QTY
                        RetBuf = Format(ZAIKO_QTY, "#,##0")
                        If Len(RetBuf) < 9 Then
                            RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                        End If
                        Printer.Print RetBuf;
                                                        
                                                        '入荷倉庫在庫
                        Printer.Print Tab(MGN_L + 77);
                        If Zaiko_Syukei_Proc(SUMI_QTY, _
                                                MI_QTY, _
                                                Last_JGYOBU, _
                                                StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                                StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                                KASO_NYUKA_SOKO & "01" & "01" & "01") Then
                            Exit Function
                        End If
                                    
                        ZAIKO_QTY = SUMI_QTY + MI_QTY
                        RetBuf = Format(ZAIKO_QTY, "#,##0")
                        If Len(RetBuf) < 9 Then
                            RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                        End If
                        Printer.Print RetBuf;
                    End If
                    
                    '2003.06.03（注文区分）削除
                    '    Printer.Print Tab(MGN_L + 76);
                    '    Select Case StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
                    '        Case CYU_KBN_SPO
                    '            Printer.Print " 緊";
                    '        Case CYU_KBN_HJU
                    '            Printer.Print " 補";
                    '        Case Else
                    '            Printer.Print " 　";
                    '    End Select
                    '2003.06.03
                                
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2013.12.25
                    If Trim(SAVE_Cyu_Kbn) = "" Then
                        Select Case StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
                            Case CYU_KBN_SPO
                                Printer.Print " 緊";
                            Case CYU_KBN_HJU
                                Printer.Print " 補";
                            Case Else
                                Printer.Print " 　";
                        End Select
                    End If
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2013.12.25
                                
                                
                                                        '伝票№
                    Printer.Print Tab(MGN_L + 90);
'                    Printer.Print Left(StrConv(Y_SYUREC.DEN_NO, vbUnicode), 6);
                    '2019.06.05 上6桁から１０桁に変更。
                    Printer.Print StrConv(Y_SYUREC.DEN_NO, vbUnicode);
                    
                                    '2019.06.05 ↓+4
                    Printer.Print Tab(MGN_L + 100 + 4);
                    TEMP_QTY = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode) - CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)))
                    RetBuf = Format(TEMP_QTY, "#,##0")
                    If Len(RetBuf) < 9 Then
                        RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                    End If
                    Printer.Print RetBuf;
                                      '2019.06.05 ↓+4
                    Printer.Print Tab(MGN_L + 120 + 4);
                                                            '印刷フォント設定（Ｃｏｄｅ３９）
                    Set Printer.Font = Code39Font
                                        'バーコード(*伝票ID*)
                    Printer.Print "*" & StrConv(Y_SYUREC.ID_NO, vbUnicode) & "*";
                                                            '印刷フォント設定（通常）
                    Set Printer.Font = NormalFont
                    
                    '-----------------------------------------------------  '２行目
                    Printer.Print Tab(MGN_L + 90);
                    Printer.Print StrConv(Y_SYUREC.ID_NO, vbUnicode);
            
            
                    Printer.Print
                    Printer.Print
                    
                    Lcnt = Lcnt + 3
            
                                                            '印刷日付設定更新
                    If Not RePrint Then
                        Call UniCode_Conv(Y_SYUREC.PRINT_YMD, Format(Now, "YYYYMMDD"))
            
                        Do
            
                            sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
            
                                    Beep
                                    ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                    If ans = vbCancel Then
                                        Print_Proc = SYS_CANCEL
                                        Exit Function
                                    End If
                                Case Else
                                    Call File_Error(sts, BtOpUpdate, "出荷予定")
                                    Print_Proc = SYS_ERR
                                    Exit Function
            
                            End Select
            
            
                        Loop
                    End If
                    
                    Print_Cnt = Print_Cnt + 1
                End If
            Next i
                
                
    End Select
    '-------------------------------    2013.12.25


    If Lcnt <> 99 Then
        Printer.EndDoc
    End If


Label3.Caption = Print_Cnt
    Call Input_UnLock

    Print_Proc = False

End Function
                                    
Private Function Head_Proc(CYU_KBN As String, Lcnt As Integer) As Integer
Dim i               As Integer
Dim sts             As Integer
Dim CYU_KBN_NAME    As String

    Head_Proc = True

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
    Printer.Print Tab(MGN_L);
    For i = 0 To UBound(JGYOBU_T)
        If Last_JGYOBU = JGYOBU_T(i).CODE Then
            Printer.Print "（" & RTrim(JGYOBU_T(i).NAME) & "）";
            Exit For
        End If
    Next i
    
    
    
    If Trim(CYU_KBN) <> "" Then                                     '2013.12.25
    
        Printer.Print Tab(MGN_L + 41);
        Select Case CYU_KBN
            Case CYU_KBN_TUK            '月切
                CYU_KBN_NAME = CYU_KBN_1
            Case CYU_KBN_SPO            'ｽﾎﾟｯﾄ
                CYU_KBN_NAME = CYU_KBN_2
            Case CYU_KBN_HJU            '補充
                CYU_KBN_NAME = CYU_KBN_3
            Case CYU_KBN_TOK            '特売り
                CYU_KBN_NAME = CYU_KBN_4
            Case CYU_KBN_BOU            '貿易
                CYU_KBN_NAME = CYU_KBN_E
        End Select
        
        
        
        Printer.Print "『" & CYU_KBN_NAME & "』出庫表";
    
    
    Else                                                            '2013.12.25
        Printer.Print "　　　　" & "出庫表";                        '2013.12.25
    End If                                                          '2013.12.25
    
    
    Printer.Print Tab(MGN_L + 91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
'2008.11.17    Printer.Print                                      '97.10.14

'2008.11.17    Printer.Print Tab(MGN_L);
'2008.11.17    Printer.Print "向け先：";
'2008.11.17    Printer.Print "[" & StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & "]" & "[" & StrConv(Y_SYUREC.SS_CODE, vbUnicode) & "]";
'2008.11.17    Printer.Print Tab(MGN_L + 30);
'2008.11.17    Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))
'2008.11.17    Call UniCode_Conv(K0_MTS.SS_CODE, StrConv(Y_SYUREC.SS_CODE, vbUnicode))
'2008.11.17    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
'2008.11.17    Select Case sts
'2008.11.17        Case BtNoErr
'2008.11.17            Printer.Print "[" & StrConv(MTSREC.MUKE_NAME, vbUnicode) & "]";
'2008.11.17            Printer.Print "[" & StrConv(MTSREC.SS_NAME, vbUnicode) & "]";
'2008.11.17        Case BtErrKeyNotFound
'2008.11.17        Case Else
'2008.11.17            Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
'2008.11.17            Exit Function
'2008.11.17    End Select
'2008.11.17
'2008.11.17    Set Printer.Font = Code39Font
'2008.11.17
'2008.11.17    If Len(Trim(StrConv(Y_SYUREC.SS_CODE, vbUnicode))) <> 0 Then
'2008.11.17        Printer.Print "*" & Trim(StrConv(Y_SYUREC.SS_CODE, vbUnicode)) & "*";
'2008.11.17    Else
'2008.11.17        Printer.Print "*" & Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) & "*";
'2008.11.17    End If
'2008.11.17    Set Printer.Font = NormalFont
    
    
    Printer.Print
    'Printer.Print                              '97.10.14
'    Printer.Print Tab(MGN_L + 90); "数量OK  ";
                                        '印刷フォント設定
'    Set Printer.Font = Code39Font
'    Printer.Print "*OK*"
'    Set Printer.Font = NormalFont
                                                '97.10.14 ここまで
    Printer.Print

    Printer.Print Tab(MGN_L);
    Printer.Print "標準棚番";
    Printer.Print Tab(MGN_L + 13);
    Printer.Print "向け先";
    Printer.Print Tab(MGN_L + 23);
    Printer.Print "品番（外部）";
    Printer.Print Tab(MGN_L + 36);
    Printer.Print "標準棚在庫";
    Printer.Print Tab(MGN_L + 48);
    Printer.Print "別置棚番";
    Printer.Print Tab(MGN_L + 60);
    Printer.Print "別置在庫";
    Printer.Print Tab(MGN_L + 69);
    Printer.Print "商品化室";
    Printer.Print Tab(MGN_L + 78);
    Printer.Print "入荷倉庫";
    Printer.Print Tab(MGN_L + 90);
    Printer.Print "伝票№";
                        '2019.06.05 ↓+4
    Printer.Print Tab(MGN_L + 103 + 4);
    Printer.Print "出荷数";
    Printer.Print

    Printer.Print

    Lcnt = 6 + MGN_U

    Head_Proc = False
End Function
Private Function Tana_Kensaku(Betu_LOCATION As String) As Integer

Dim sts As Integer

    Tana_Kensaku = True
    
    Betu_LOCATION = ""
    
    Call UniCode_Conv(K6_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K6_ZAIKO.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K6_ZAIKO.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
    Call UniCode_Conv(K6_ZAIKO.NYUKA_DT, "")
    Call UniCode_Conv(K6_ZAIKO.Soko_No, "")
    Call UniCode_Conv(K6_ZAIKO.Retu, "")
    Call UniCode_Conv(K6_ZAIKO.Ren, "")
    Call UniCode_Conv(K6_ZAIKO.Dan, "")
    
    Do
        sts = BTRV(BtOpGetGreater, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K6_ZAIKO, Len(K6_ZAIKO), 6)
        Select Case sts
                Case BtNoErr
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> StrConv(Y_SYUREC.NAIGAI, vbUnicode) Or _
                    Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> Trim(StrConv(Y_SYUREC.HIN_NO, vbUnicode)) Then
                    Exit Do
                End If
                If StrConv(ZAIKOREC.Soko_No, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2) Or _
                   StrConv(ZAIKOREC.Retu, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 3, 2) Or _
                   StrConv(ZAIKOREC.Ren, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 5, 2) Or _
                   StrConv(ZAIKOREC.Dan, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 7, 2) Then
                                                'システム倉庫の判定
                    Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ZAIKOREC.Soko_No, vbUnicode))
                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    Select Case sts
                        Case BtNoErr
                            If StrConv(SOKOREC.SOKO_BUN, vbUnicode) <> BUN_KASO Then
                                Betu_LOCATION = StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                                StrConv(ZAIKOREC.Dan, vbUnicode)
                                Exit Do
                        
                            End If
                        Case BtErrKeyNotFound
                                                '考えられないので読み飛ばし
                        Case Else
                            Call File_Error(sts, BtOpGetGreater, "倉庫マスタ")
                            Exit Function
                    End Select
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetGreater, "在庫データ")
                Exit Function
        End Select
            
            
    Loop
    
    Tana_Kensaku = False

End Function


Private Function Y_Syu_Get(RePrint As Boolean, com As Integer) As Integer

Dim sts         As Integer
Dim OP          As Integer
Dim ans         As Integer

Dim i           As Integer
Dim Skip_Flg    As Boolean

    
    
    Y_Syu_Get = False
    
    
    
'2008.11.17    If com = BtOpGetGreaterEqual Then
'2008.11.17                                        '最初のＫＥＹセット
'2008.11.17        Call UniCode_Conv(K5_Y_SYU.JGYOBU, Last_JGYOBU)
'2008.11.17        If Right(Combo(pcmbCyu_Kbn).Text, 1) <> " " Then
'2008.11.17            Call UniCode_Conv(K5_Y_SYU.KEY_CYU_KBN, Right(Combo(pcmbCyu_Kbn).Text, 1))
'2008.11.17        Else
'2008.11.17            Call UniCode_Conv(K5_Y_SYU.KEY_CYU_KBN, "")
'2008.11.17        End If
'2008.11.17        Call UniCode_Conv(K5_Y_SYU.KEY_MUKE_CODE, "")
'2008.11.17        Call UniCode_Conv(K5_Y_SYU.KEY_SS_CODE, "")
'2008.11.17        Call UniCode_Conv(K5_Y_SYU.HTANABAN, "")
'2008.11.17        Call UniCode_Conv(K5_Y_SYU.KEY_SYUKA_YMD, "")
'2008.11.17        Call UniCode_Conv(K5_Y_SYU.KEY_HIN_NO, "")
'2008.11.17    End If
    
    
    
    '2008.11.17 ↓
    If com = BtOpGetGreaterEqual Then
                                        '最初のＫＥＹセット
        Call UniCode_Conv(K6_Y_SYU.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K6_Y_SYU.KEY_CYU_KBN, Right(Combo(pcmbCyu_Kbn).Text, 1))
        Call UniCode_Conv(K6_Y_SYU.HTANABAN, "")
        Call UniCode_Conv(K6_Y_SYU.NAIGAI, "")
        Call UniCode_Conv(K6_Y_SYU.KEY_HIN_NO, "")
    End If
    '2008.11.17 ↑
    
    
    
    
    OP = com + BtSNoWait
    
    Do
        Do
            sts = BTRV(OP, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
            Select Case sts
                Case BtNoErr
                    '事業部のﾌﾞﾚｰｸ
                    If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                        
                        sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K5_Y_SYU, Len(K5_Y_SYU), 5)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "出荷予定ファイル")
                            Y_Syu_Get = sts
                            Exit Function
                        End If
                        Y_Syu_Get = BtErrEOF
                        Exit Function
                    End If
                    '指定が有れば注文区分をﾁｪｯｸ
                    If Right(Combo(pcmbCyu_Kbn).Text, 1) <> " " Then
                        If StrConv(Y_SYUREC.CYU_KBN, vbUnicode) <> Right(Combo(pcmbCyu_Kbn).Text, 1) Then
                            
                            sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K5_Y_SYU, Len(K5_Y_SYU), 5)
                            If sts Then
                                Call File_Error(sts, BtOpUnlock, "出荷予定ファイル")
                                Y_Syu_Get = sts
                                Exit Function
                            End If
                            
                            Y_Syu_Get = BtErrEOF
                            Exit Function
                        End If
                    End If
                    Exit Do
                Case BtErrEOF
                    Y_Syu_Get = sts
                    Exit Function
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Y_Syu_Get = BtErrEOF
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, OP, "出荷予定ファイル")
                    Y_Syu_Get = sts
                    Exit Function
            End Select
        
        Loop
                    
        Skip_Flg = False
                                '向け先 KEYﾌﾞﾚｰｸ
        If Trim(Text(ptxMUKE_CODE).Text) <> "" Then
            If Trim(Right(Combo(pcmbMUKE_Code).Text, 16)) <> "" Then
                If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) <> Trim(Left(Right(Combo(pcmbMUKE_Code).Text, 16), 8)) Or _
                    Trim(StrConv(Y_SYUREC.SS_CODE, vbUnicode)) <> Trim(Right(Combo(pcmbMUKE_Code).Text, 8)) Then
                    
                    
                    
                    sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K5_Y_SYU, Len(K5_Y_SYU), 5)
                    If sts Then
                        Call File_Error(sts, BtOpUnlock, "出荷予定ファイル")
                        Y_Syu_Get = sts
                        Exit Function
                    End If
                    Skip_Flg = True
                End If
            End If
        Else
            If NON_MUKE_FLG Then
                For i = 0 To UBound(NON_MUKE_CODE)
                    If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = Trim(NON_MUKE_CODE(i)) Then
                        Skip_Flg = True
                        Exit For
                    End If
                Next i
            End If
        End If
        
                                '処理完了済
        If CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) = CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) Then
            Skip_Flg = True
        End If
                                '印刷区分
        If Trim(Right(Combo(pcmbPRINT_KBN).Text, 1)) <> "" Then
            If Right(Combo(pcmbPRINT_KBN).Text, 1) = Print_KBN_SIN Then
                If IsDate(Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 1, 4) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 5, 2) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 7, 2)) Then
                    Skip_Flg = True
                End If
            Else
                If Not IsDate(Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 1, 4) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 5, 2) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 7, 2)) Then
                    Skip_Flg = True
                End If
            End If
        End If
        
                                '伝票日付範囲(開始)
        If StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) < (Text(ptxS_DEN_DT_YY).Text & Text(ptxS_DEN_DT_MM).Text & Text(ptxS_DEN_DT_DD).Text) Then
            Skip_Flg = True
        End If
                                '伝票日付範囲(終了)
        If Trim(Text(ptxE_DEN_DT_YY).Text & Text(ptxE_DEN_DT_MM).Text & Text(ptxE_DEN_DT_DD).Text) <> "" Then
            If StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) > (Text(ptxE_DEN_DT_YY).Text & Text(ptxE_DEN_DT_MM).Text & Text(ptxE_DEN_DT_DD).Text) Then
                Skip_Flg = True
            End If
        End If
                                '伝票番号
        If Trim(Text(ptxDEN_NO).Text) <> "" Then
            If Trim(StrConv(Y_SYUREC.DEN_NO, vbUnicode)) <> Trim(Text(ptxDEN_NO)) Then
                Skip_Flg = True
            End If
        Else
'''伝票№桁数指定、廃止
'''            If Len(Trim(StrConv(Y_SYUREC.DEN_NO, vbUnicode))) > 5 Then
'''                Skip_Flg = True
'''            End If
        End If
                                
        If Not Skip_Flg Then
                    
            Skip_Flg = True
                    
            For i = Min_Row To SYUKA.UpperBound(1)
        
                If StrConv(Y_SYUREC.ID_NO, vbUnicode) = SYUKA(i, ColID_NO) Then
                    If SYUKA(i, ColSEL) Then
                        Skip_Flg = False
                        Exit For
                    End If
                End If
        
            Next i
            
            If Not Skip_Flg Then
                If Not IsDate(Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 1, 4) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 5, 2) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 7, 2)) Then
                    RePrint = False
            
        
                    Call UniCode_Conv(Y_SYUREC.PRINT_YMD, Format(Now, "YYYYMMDD"))
                    
                    Do
                
                        sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Y_Syu_Get = BtErrEOF
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpUpdate, "出荷予定")
                                Y_Syu_Get = sts
                                Exit Function
                                
                        End Select
                    Loop
            
                
                Else
                    RePrint = True
                
                
                    sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K5_Y_SYU, Len(K5_Y_SYU), 5)
                    If sts Then
                        Call File_Error(sts, BtOpUnlock, "出荷予定ファイル")
                        Y_Syu_Get = sts
                        Exit Function
                    End If
                
                
                End If
            
                Y_Syu_Get = BtNoErr
                Exit Function
            End If
                    
        Else
            
            sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K5_Y_SYU, Len(K5_Y_SYU), 5)
            If sts Then
                Call File_Error(sts, BtOpUnlock, "出荷予定ファイル")
                Y_Syu_Get = sts
                Exit Function
            End If
                    
        End If
                    
                    
    
        OP = BtOpGetNext + BtSNoWait
    
    Loop
End Function


Private Function S8_LOCATION_Proc(Soko_No As String, _
                                        Betu_LOCATION As String) As Integer


Dim sts     As Integer


    S8_LOCATION_Proc = SYS_ERR


    Betu_LOCATION = ""


    Call UniCode_Conv(K4_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K4_ZAIKO.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K4_ZAIKO.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
    Call UniCode_Conv(K4_ZAIKO.Soko_No, Soko_No)
    Call UniCode_Conv(K4_ZAIKO.Retu, "")
    Call UniCode_Conv(K4_ZAIKO.Ren, "")
    Call UniCode_Conv(K4_ZAIKO.Dan, "")
    
    sts = BTRV(BtOpGetGreater, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K4_ZAIKO, Len(K4_ZAIKO), 4)
    Select Case sts
        Case BtNoErr
            If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> StrConv(Y_SYUREC.NAIGAI, vbUnicode) Or _
                Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> Trim(StrConv(Y_SYUREC.HIN_NO, vbUnicode)) Or _
                StrConv(ZAIKOREC.Soko_No, vbUnicode) <> Soko_No Then
            Else
                Betu_LOCATION = StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                StrConv(ZAIKOREC.Dan, vbUnicode)
            End If
        Case BtErrEOF
        Case Else
            Call File_Error(sts, BtOpGetGreater, "在庫データ")
            Exit Function
    End Select


    S8_LOCATION_Proc = False

End Function


Private Sub Text_LostFocus(Index As Integer)


    Text(Index).Text = StrConv(Text(Index).Text, vbUpperCase)

End Sub
