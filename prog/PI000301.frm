VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PI000301 
   Caption         =   "資材注文書発行"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16545
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
   ScaleHeight     =   10305
   ScaleWidth      =   16545
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox txtAVE_SYUKA 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Left            =   7200
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtAVE_SYUKA_cnt 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Left            =   7200
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.ListBox lotLIST 
      Height          =   780
      Left            =   5880
      Sorted          =   -1  'True
      TabIndex        =   88
      Top             =   3480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   45
      Left            =   4320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   44
      Left            =   14400
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   34
      Left            =   12480
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   24
      Left            =   10560
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   10
      Left            =   4305
      MaxLength       =   7
      TabIndex        =   11
      Top             =   2760
      Width           =   915
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   2280
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   83
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   11400
      TabIndex        =   82
      Text            =   "Text2"
      Top             =   9720
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   43
      Left            =   14400
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   42
      Left            =   14400
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   41
      Left            =   14400
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   40
      Left            =   14400
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   39
      Left            =   14400
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   38
      Left            =   14400
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   37
      Left            =   14400
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   36
      Left            =   14400
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   35
      Left            =   14400
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   33
      Left            =   12480
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   32
      Left            =   12480
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   31
      Left            =   12480
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   30
      Left            =   12480
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   29
      Left            =   12480
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   28
      Left            =   12480
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   27
      Left            =   12480
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   26
      Left            =   12480
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   25
      Left            =   12480
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   23
      Left            =   10560
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   22
      Left            =   10560
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   21
      Left            =   10560
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   20
      Left            =   10560
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   19
      Left            =   10560
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   18
      Left            =   10560
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   17
      Left            =   10560
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   16
      Left            =   10560
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   15
      Left            =   10560
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   14
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   13
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   12
      Left            =   4305
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   11
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   12
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   9
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   8
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   9
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   7
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2040
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   1
      Left            =   2280
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   7
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   4
      Left            =   4080
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   1575
      MaxLength       =   20
      TabIndex        =   3
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   2295
      Left            =   420
      TabIndex        =   46
      Top             =   4320
      Width           =   15915
      _ExtentX        =   28072
      _ExtentY        =   4048
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
      Columns(1).Caption=   "注文日"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "注文先名"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "資材品番"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "品名"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "数量"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "希望納期日"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "納入先"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1984"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=512"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2011"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1905"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=512"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3810"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3704"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=3493"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3387"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=512"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=4180"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=4075"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=512"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2778"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2672"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=2328"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2223"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=512"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=6085"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=5980"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=512"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
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
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=0,.fontsize=1200"
      _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=128"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=43,.parent=1,.bold=0,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(25)  =   ":id=43,.strikethrough=0,.charset=128"
      _StyleDefs(26)  =   ":id=43,.fontname=ＭＳ ゴシック"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=45,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=47,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=58,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(39)  =   ":id=58,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(40)  =   ":id=58,.fontname=ＭＳ ゴシック"
      _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=44"
      _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=45"
      _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=47"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=62,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(45)  =   ":id=62,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(46)  =   ":id=62,.fontname=ＭＳ ゴシック"
      _StyleDefs(47)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=44"
      _StyleDefs(48)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=45"
      _StyleDefs(49)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=47"
      _StyleDefs(50)  =   "Splits(0).Columns(2).Style:id=16,.parent=43"
      _StyleDefs(51)  =   "Splits(0).Columns(2).HeadingStyle:id=13,.parent=44"
      _StyleDefs(52)  =   "Splits(0).Columns(2).FooterStyle:id=14,.parent=45"
      _StyleDefs(53)  =   "Splits(0).Columns(2).EditorStyle:id=15,.parent=47"
      _StyleDefs(54)  =   "Splits(0).Columns(3).Style:id=28,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(55)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(56)  =   ":id=28,.fontname=ＭＳ ゴシック"
      _StyleDefs(57)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=44"
      _StyleDefs(58)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=45"
      _StyleDefs(59)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=47"
      _StyleDefs(60)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(61)  =   ":id=66,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(62)  =   ":id=66,.fontname=ＭＳ ゴシック"
      _StyleDefs(63)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44"
      _StyleDefs(64)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
      _StyleDefs(65)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
      _StyleDefs(66)  =   "Splits(0).Columns(5).Style:id=32,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(67)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(68)  =   ":id=32,.fontname=ＭＳ ゴシック"
      _StyleDefs(69)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=44"
      _StyleDefs(70)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=45"
      _StyleDefs(71)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=47"
      _StyleDefs(72)  =   "Splits(0).Columns(6).Style:id=70,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(73)  =   ":id=70,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(74)  =   ":id=70,.fontname=ＭＳ ゴシック"
      _StyleDefs(75)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=44"
      _StyleDefs(76)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=45"
      _StyleDefs(77)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=47"
      _StyleDefs(78)  =   "Splits(0).Columns(7).Style:id=74,.parent=43,.alignment=0"
      _StyleDefs(79)  =   "Splits(0).Columns(7).HeadingStyle:id=71,.parent=44"
      _StyleDefs(80)  =   "Splits(0).Columns(7).FooterStyle:id=72,.parent=45"
      _StyleDefs(81)  =   "Splits(0).Columns(7).EditorStyle:id=73,.parent=47"
      _StyleDefs(82)  =   "Named:id=33:Normal"
      _StyleDefs(83)  =   ":id=33,.parent=0"
      _StyleDefs(84)  =   "Named:id=34:Heading"
      _StyleDefs(85)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(86)  =   ":id=34,.wraptext=-1"
      _StyleDefs(87)  =   "Named:id=35:Footing"
      _StyleDefs(88)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(89)  =   "Named:id=36:Selected"
      _StyleDefs(90)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(91)  =   "Named:id=37:Caption"
      _StyleDefs(92)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(93)  =   "Named:id=38:HighlightRow"
      _StyleDefs(94)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(95)  =   "Named:id=39:EvenRow"
      _StyleDefs(96)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(97)  =   "Named:id=40:OddRow"
      _StyleDefs(98)  =   ":id=40,.parent=33"
      _StyleDefs(99)  =   "Named:id=41:RecordSelector"
      _StyleDefs(100) =   ":id=41,.parent=34"
      _StyleDefs(101) =   "Named:id=42:FilterBar"
      _StyleDefs(102) =   ":id=42,.parent=33"
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   6
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終 了"
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
      TabIndex        =   58
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "印 刷"
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
      TabIndex        =   55
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "消 去"
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
      TabIndex        =   54
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4920
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "注文数"
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
      Left            =   4080
      TabIndex        =   51
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取 消"
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
      TabIndex        =   50
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "更 新"
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
      TabIndex        =   47
      Top             =   9720
      Width           =   855
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid2 
      Height          =   2775
      Left            =   420
      TabIndex        =   81
      Top             =   6720
      Width           =   16035
      _ExtentX        =   28284
      _ExtentY        =   4895
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "注文日"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "注文№"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "注文先"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "資材品番"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "品名"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "注文数"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "注文残"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "在庫残"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "希望納期日"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "回答納期日"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "使用月"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2170"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2064"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=512"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1614"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1508"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3836"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3731"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=512"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2699"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2593"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=512"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=5318"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=5212"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=512"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=1773"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1667"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=1773"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1667"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=1773"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=1667"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=514"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=2328"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2223"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=512"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(9).Width=2302"
      Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2196"
      Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=516"
      Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(51)=   "Column(10).Width=1429"
      Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=1323"
      Splits(0)._ColumnProps(54)=   "Column(10)._ColStyle=516"
      Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
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
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=0,.fontsize=1200"
      _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=128"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=43,.parent=1,.bold=0,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(25)  =   ":id=43,.strikethrough=0,.charset=128"
      _StyleDefs(26)  =   ":id=43,.fontname=ＭＳ ゴシック"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=45,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=47,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=58,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(39)  =   ":id=58,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(40)  =   ":id=58,.fontname=ＭＳ ゴシック"
      _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=44"
      _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=45"
      _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=47"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=16,.parent=43"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=13,.parent=44"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=14,.parent=45"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=15,.parent=47"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=28,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(49)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(50)  =   ":id=28,.fontname=ＭＳ ゴシック"
      _StyleDefs(51)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=44"
      _StyleDefs(52)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=45"
      _StyleDefs(53)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=47"
      _StyleDefs(54)  =   "Splits(0).Columns(3).Style:id=66,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(55)  =   ":id=66,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(56)  =   ":id=66,.fontname=ＭＳ ゴシック"
      _StyleDefs(57)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=44"
      _StyleDefs(58)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=45"
      _StyleDefs(59)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=47"
      _StyleDefs(60)  =   "Splits(0).Columns(4).Style:id=32,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(61)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(62)  =   ":id=32,.fontname=ＭＳ ゴシック"
      _StyleDefs(63)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=44"
      _StyleDefs(64)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=45"
      _StyleDefs(65)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=47"
      _StyleDefs(66)  =   "Splits(0).Columns(5).Style:id=74,.parent=43,.alignment=1"
      _StyleDefs(67)  =   "Splits(0).Columns(5).HeadingStyle:id=71,.parent=44"
      _StyleDefs(68)  =   "Splits(0).Columns(5).FooterStyle:id=72,.parent=45"
      _StyleDefs(69)  =   "Splits(0).Columns(5).EditorStyle:id=73,.parent=47"
      _StyleDefs(70)  =   "Splits(0).Columns(6).Style:id=20,.parent=43,.alignment=1"
      _StyleDefs(71)  =   "Splits(0).Columns(6).HeadingStyle:id=17,.parent=44"
      _StyleDefs(72)  =   "Splits(0).Columns(6).FooterStyle:id=18,.parent=45"
      _StyleDefs(73)  =   "Splits(0).Columns(6).EditorStyle:id=19,.parent=47"
      _StyleDefs(74)  =   "Splits(0).Columns(7).Style:id=24,.parent=43,.alignment=1"
      _StyleDefs(75)  =   "Splits(0).Columns(7).HeadingStyle:id=21,.parent=44"
      _StyleDefs(76)  =   "Splits(0).Columns(7).FooterStyle:id=22,.parent=45"
      _StyleDefs(77)  =   "Splits(0).Columns(7).EditorStyle:id=23,.parent=47"
      _StyleDefs(78)  =   "Splits(0).Columns(8).Style:id=62,.parent=43,.alignment=0"
      _StyleDefs(79)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=44"
      _StyleDefs(80)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=45"
      _StyleDefs(81)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=47"
      _StyleDefs(82)  =   "Splits(0).Columns(9).Style:id=70,.parent=43"
      _StyleDefs(83)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=44"
      _StyleDefs(84)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=45"
      _StyleDefs(85)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=47"
      _StyleDefs(86)  =   "Splits(0).Columns(10).Style:id=78,.parent=43"
      _StyleDefs(87)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=44"
      _StyleDefs(88)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=45"
      _StyleDefs(89)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=47"
      _StyleDefs(90)  =   "Named:id=33:Normal"
      _StyleDefs(91)  =   ":id=33,.parent=0"
      _StyleDefs(92)  =   "Named:id=34:Heading"
      _StyleDefs(93)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(94)  =   ":id=34,.wraptext=-1"
      _StyleDefs(95)  =   "Named:id=35:Footing"
      _StyleDefs(96)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(97)  =   "Named:id=36:Selected"
      _StyleDefs(98)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(99)  =   "Named:id=37:Caption"
      _StyleDefs(100) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(101) =   "Named:id=38:HighlightRow"
      _StyleDefs(102) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(103) =   "Named:id=39:EvenRow"
      _StyleDefs(104) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(105) =   "Named:id=40:OddRow"
      _StyleDefs(106) =   ":id=40,.parent=33"
      _StyleDefs(107) =   "Named:id=41:RecordSelector"
      _StyleDefs(108) =   ":id=41,.parent=34"
      _StyleDefs(109) =   "Named:id=42:FilterBar"
      _StyleDefs(110) =   ":id=42,.parent=33"
   End
   Begin VB.Label lblTuki 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   95
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "月平均出荷数"
      Height          =   255
      Index           =   13
      Left            =   5640
      TabIndex        =   94
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblTuki 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   0
      Left            =   6840
      TabIndex        =   92
      Top             =   1920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "月平均出荷件数"
      Height          =   240
      Index           =   15
      Left            =   5400
      TabIndex        =   91
      Top             =   1560
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label lblSHIIRE_BIKOU 
      Height          =   255
      Left            =   4080
      TabIndex        =   89
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "受入済数"
      Height          =   255
      Index           =   12
      Left            =   3120
      TabIndex        =   86
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "設定日"
      Height          =   375
      Index           =   8
      Left            =   9240
      TabIndex        =   85
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "使用月"
      Height          =   255
      Index           =   13
      Left            =   3465
      TabIndex        =   84
      Top             =   2880
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "最適注文先選択"
      Height          =   375
      Index           =   7
      Left            =   9240
      TabIndex        =   80
      Top             =   240
      Width           =   7095
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "注文残"
      Height          =   375
      Index           =   10
      Left            =   9240
      TabIndex        =   79
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "納期予定日"
      Height          =   375
      Index           =   9
      Left            =   9240
      TabIndex        =   78
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "前回注文数"
      Height          =   375
      Index           =   6
      Left            =   9240
      TabIndex        =   77
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "前回注文日"
      Height          =   375
      Index           =   5
      Left            =   9240
      TabIndex        =   76
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "ﾘｰﾄﾞﾀｲﾑ"
      Height          =   375
      Index           =   4
      Left            =   9240
      TabIndex        =   75
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "ﾛｯﾄ数"
      Height          =   375
      Index           =   3
      Left            =   9240
      TabIndex        =   74
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "単価"
      Height          =   375
      Index           =   2
      Left            =   9240
      TabIndex        =   73
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BorderStyle     =   1  '実線
      Caption         =   "注文先名"
      Height          =   375
      Index           =   1
      Left            =   9240
      TabIndex        =   72
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  '実線
      Caption         =   "注文先ｺｰﾄﾞ"
      Height          =   375
      Index           =   0
      Left            =   9240
      TabIndex        =   71
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "発注ﾛｯﾄ"
      Height          =   255
      Index           =   11
      Left            =   480
      TabIndex        =   70
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "在庫残"
      Height          =   255
      Index           =   10
      Left            =   720
      TabIndex        =   69
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "金額"
      Height          =   255
      Index           =   9
      Left            =   3675
      TabIndex        =   68
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "単価"
      Height          =   255
      Index           =   8
      Left            =   960
      TabIndex        =   67
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "希望納期日"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   66
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "注文数"
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   65
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "注文№"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   64
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "納入先"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   63
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "注文先"
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   62
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "資材品番"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   61
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "注文日"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   60
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "担当者"
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   59
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "PI000301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WS_NO       As String * 10
    
'ラベル用添字
Private Const plblUSE_YM% = 13              '使用月 2007.12.05
    
'テキスト用添字
Private Const ptxTANTO_CODE% = 0            '担当者ｺｰﾄﾞ
Private Const ptxTANTO_NAME% = 1            '担当者名称
Private Const ptxORDER_DT% = 2              '注文日
Private Const ptxHIN_GAI% = 3               '品番
Private Const ptxHIN_NAME% = 4              '品名
Private Const ptxORDER_CODE% = 5            '注文先
Private Const ptxDELI_CODE% = 6             '納入先
Private Const ptxORDER_NO% = 7              '注文№
Private Const ptxORDER_QTY% = 8             '注文数

Private Const ptxUKEIRE_QTY% = 45           '受入数 2016.06.22


Private Const ptxY_NOUKI_DT% = 9            '納期予定日

Private Const ptxUSE_YM% = 10               '使用月 2007.01.08      ↓以下１ずつスライド


Private Const ptxTANKA% = 11                '単価
Private Const ptxKINGAKU% = 12              '金額

Private Const ptxZAIKO_QTY% = 13            '在庫残
Private Const ptxLOT% = 14                  '発注ﾛｯﾄ

Private Const ptxSHIIRE_CODE01% = 15        '仕入先ｺｰﾄﾞ 1
Private Const ptxSHIIRE_NAME01% = 16        '仕入先名称 1
Private Const ptxSHIIRE_TANKA01% = 17       '仕入単価 1
Private Const ptxSHIIRE_LOT01% = 18         '仕入ﾛｯﾄ 1
Private Const ptxSHIIRE_LT01% = 19          '仕入ﾘｰﾄﾞﾀｲﾑ 1
Private Const ptxZEN_ORDER_DT01% = 20       '前回注文日 1
Private Const ptxZEN_ORDER_QTY01% = 21      '前回注文数 1
Private Const ptxY_NOUKI_DT01% = 22         '納期予定日 1
Private Const ptxORDER_ZAN01% = 23          '注文残 1

Private Const ptxTANKA_DT01% = 24           '単価設定日 1


Private Const ptxSHIIRE_CODE02% = 25        '仕入先ｺｰﾄﾞ 2
Private Const ptxSHIIRE_NAME02% = 26        '仕入先名称 2
Private Const ptxSHIIRE_TANKA02% = 27       '仕入単価 2
Private Const ptxSHIIRE_LOT02% = 28         '仕入ﾛｯﾄ 2
Private Const ptxSHIIRE_LT02% = 29          '仕入ﾘｰﾄﾞﾀｲﾑ 2
Private Const ptxZEN_ORDER_DT02% = 30       '前回注文日 2
Private Const ptxZEN_ORDER_QTY02% = 31      '前回注文数 2
Private Const ptxY_NOUKI_DT02% = 32         '納期予定日 2
Private Const ptxORDER_ZAN02% = 33          '注文残 2

Private Const ptxTANKA_DT02% = 34           '単価設定日 2


Private Const ptxSHIIRE_CODE03% = 35        '仕入先ｺｰﾄﾞ 3
Private Const ptxSHIIRE_NAME03% = 36        '仕入先名称 3
Private Const ptxSHIIRE_TANKA03% = 37       '仕入単価 3
Private Const ptxSHIIRE_LOT03% = 38         '仕入ﾛｯﾄ 3
Private Const ptxSHIIRE_LT03% = 39          '仕入ﾘｰﾄﾞﾀｲﾑ 3
Private Const ptxZEN_ORDER_DT03% = 40       '前回注文日 3
Private Const ptxZEN_ORDER_QTY03% = 41      '前回注文数 3
Private Const ptxY_NOUKI_DT03% = 42         '納期予定日 3
Private Const ptxORDER_ZAN03% = 43          '注文残 3

Private Const ptxTANKA_DT03% = 44           '単価設定日 3

'コンボ用添字
Private Const pcmbORDER% = 0                '注文先
Private Const pcmbDELI% = 1                 '納入先


'Glid用環境

Private SHORDER  As New XArrayDB

Private Const Min_Row% = 1                  '最小行数
Private Const Min_Col% = 0                  '最小列数
Private Const Max_Col% = 7                  '最大列数


Private Const colORDER_NO% = 0              '注文№
Private Const colORDER_DT% = 1              '注文日
Private Const colORDER_NAME% = 2            '発注先名
Private Const colHIN_GAI% = 3               '品番
Private Const colHIN_NAME% = 4              '品名
Private Const colORDER_QTY% = 5             '注文数
Private Const colY_NOUKI_DT% = 6            '納期予定日
Private Const colDELI_NAME% = 7             '納入先



Private Sort_Tbl(colORDER_NO To colDELI_NAME) _
                As Integer                  'ｿｰﾄの制御 0:昇順 1:降順
Private Tbl_Set_F   As Boolean
                                            
                                            
                                            
                                            
                                            
                                            
'---------------    注文残用    2007.07.27


Private Z_SHORDER  As New XArrayDB



Private Const Z_Min_Row% = 1                '最小行数
Private Const Z_Min_Col% = 0                '最小列数
Private Const Z_Max_Col% = 10               '最大列数   8-->10 2007.12.05

Private Const colZ_ORDER_DT% = 0            '注文日時
Private Const colZ_ORDER_NO% = 1            'CODE
Private Const colZ_ORDER_NAME% = 2          '注文先名
Private Const colZ_HIN_GAI% = 3             '資材品番
Private Const colZ_HIN_NAME% = 4            '品名
Private Const colZ_ORDER_QTY% = 5           '注文数
Private Const colZ_ZAN_QTY% = 6             '注文残
Private Const colZ_ZAIKO_QTY% = 7           '在庫数
Private Const colZ_Y_NOUKI_DT% = 8          '予定納期

Private Const colZ_ANS_NOUKI_DT% = 9        '回答納期日 2008.01.10
Private Const colZ_USE_YM% = 10             '使用月 2008.01.10
                                            
                                            
Private Z_Sort_Tbl(colZ_ORDER_DT To colZ_Y_NOUKI_DT) _
                As Integer                  'ｿｰﾄの制御 0:昇順 1:降順
Private Z_Tbl_Set_F   As Boolean
                                            
                                            
Private svHinban    As String               '表示制御用品番
                                            
                                            
'---------------    注文残用    2007.07.27
Private NOUNYU      As String * 5


'---------------    予定納期省略可否  0:必須入力　1:省略可      2007.09.06
Private YOTEI_NOUKI As Integer



'---------------    入力モード   0：通常 1:注文数   2007.11.12
Private Input_Mode  As Integer

'---------------    大阪ＰＣモード  True:大阪PC　False:以外     2008.01.10
Private OSAKA_MODE  As Boolean


'---------------    注文書　再発行有無(=1:再発行有り　以外：なし    2013.02.14
Private REPRINT_FLG As Boolean

'---------------    キャンセル時のログ                  2016.04.25
Private PI00030_LOG         As String

Private LIST_MAX            As Long                     '2017.11.21


Private SHIIRE_SELECT       As Integer                  '2017.11.22


'Private Const Last_Update_Day$ = "[PI00030] 2018.04.20 09:00"
'Private Const Last_Update_Day$ = "[PI00030] 2018.11.24 18:00"
'Private Const Last_Update_Day$ = "[PI00030] 2019.11.01 17:15 注文数6桁エラー対応"  'lot5桁から6桁に変更 注文数6桁ENTERでエラーの為
Private Const Last_Update_Day$ = "[PI00030] 2019.11.07 14:00 ActiveReportライセンス対応"

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PI000301.MousePointer = vbHourglass

    Call Ctrl_Lock(PI000301)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PI000301)


    PI000301.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer

Dim Sumi_QTY    As Long
Dim Mi_QTY      As Long
    
    
Dim svTanka     As String   '2009.08.03
    
    
Dim SHIIRE_I    As Integer  '2017.11.21
    
Dim SHIIRE_LOT          As String   '2017.11.21
    
Dim SHIIRE_LOT01        As String   '2017.11.21
Dim SHIIRE_LOT02        As String   '2017.11.21
Dim SHIIRE_LOT03        As String   '2017.11.21
    
    
Dim SHIIRE_LOT_T(0 To 2) As String   '2017.12.07
Dim SHIIRE_LOT_J(0 To 2) As String   '2017.12.07
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        
        Case ptxTANTO_CODE      '担当者
        
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTANTO_CODE).Text)

            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    Text1(ptxTANTO_NAME).Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    Text1(ptxTANTO_NAME).Text = ""
            
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Exit Function
            
            End Select
        
        Case ptxORDER_DT        '注文日
            
            If Not IsDate(Text1(ptxORDER_DT).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxORDER_DT).Text = Format(CDate(Text1(ptxORDER_DT).Text), "YYYY/MM/DD")
            End If
        
        
        Case ptxHIN_GAI         '品番
    
            Text1(Mode).Text = StrConv(RTrim(Text1(Mode).Text), vbUpperCase)
                
                    
            Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
    
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                                    
                                                    '月平均出荷数集計データより月平均出荷数獲得 2018.04.19
                    Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, SHIZAI)
                    Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, Text1(ptxHIN_GAI).Text)
                    sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
                    
                    Select Case sts
                        Case BtNoErr
                        
                            txtAVE_SYUKA_cnt.Text = Format(Val(StrConv(AVE_SYUKAREC.TOTAL_AVE_CNT, vbUnicode)), "#,##0.0")
                            txtAVE_SYUKA.Text = Format(Val(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode)), "#,##0.0")
                        
                        
                        
                        
                        
                        Case BtErrKeyNotFound
                            txtAVE_SYUKA_cnt.Text = ""
                            txtAVE_SYUKA.Text = ""
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "月平均出荷数集計データ")
                            Unload Me
                    End Select
                
                
                
                
                Case BtErrKeyNotFound
                    
                    
                    Call UniCode_Conv(ITEMREC.HIN_NAME, "")     '2018.04.19
                    Call UniCode_Conv(ITEMREC.SHIIRE_BIKOU, "") '2018.04.19
                    
                    
                    Text1(ptxHIN_NAME).Text = ""
                    Text1(ptxZAIKO_QTY).Text = ""
            
            
            
                    txtAVE_SYUKA.Text = ""
                    txtAVE_SYUKA_cnt.Text = ""
            
            
                    For i = 0 To 2
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).CODE, "")
                    Next i
                    
                    
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            
            End Select
            
            Text1(ptxHIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
            
            
            lblSHIIRE_BIKOU.Caption = StrConv(ITEMREC.SHIIRE_BIKOU, vbUnicode)  '2018.04.19
            
            
            
            
            
            
            
            If Zaiko_Syukei_Proc(Sumi_QTY, Mi_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                Exit Function
            
            End If
            Text1(ptxZAIKO_QTY).Text = Format(Sumi_QTY + Mi_QTY, "#0")  '2007.10.30
        
        
''''''''''''''''    2011.09.28
            For i = ptxSHIIRE_CODE01 To ptxTANKA_DT03
            
                Text1(i).Text = ""
            
            Next i
''''''''''''''''    2011.09.28
        
            
            
            j = ptxSHIIRE_CODE01
            For i = 0 To 2
            
                If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(i).CODE, vbUnicode)) = "" Then
                Else
                
                    Text1(j).Text = StrConv(ITEMREC.G_SHIIRE_TBL(i).CODE, vbUnicode)
                    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(j).Text)
                
                    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                    Select Case sts
                        Case BtNoErr
                        
                                                    
                        
                        Case BtErrKeyNotFound
                            
                            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
                    
                    
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
                            Exit Function
                    
                    End Select
                    '仕入先名
                    Text1(j + 1).Text = StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode)
                    '単価
                    If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA, vbUnicode)) Then
                        Text1(j + 2).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA, vbUnicode)), "#0.00")
                    Else
                        Text1(j + 2).Text = ""
                    End If
                    'ﾛｯﾄ数
                    If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(i).LOT, vbUnicode)) Then
                        Text1(j + 3).Text = Format(CLng(StrConv(ITEMREC.G_SHIIRE_TBL(i).LOT, vbUnicode)), "#0")
                    Else
                        Text1(j + 3).Text = ""
                    End If
                    'ﾘｰﾄﾞﾀｲﾑ
                    If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(i).LEAD_TIME, vbUnicode)) Then
                        Text1(j + 4).Text = Format(CLng(StrConv(ITEMREC.G_SHIIRE_TBL(i).LEAD_TIME, vbUnicode)), "#0")
                    Else
                        Text1(j + 4).Text = ""
                    End If
                    '前回注文日
                    If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, vbUnicode)) = "" Then
                        Text1(j + 5).Text = ""
                    Else
                        Text1(j + 5).Text = Mid(StrConv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, vbUnicode), 1, 4) & "/" & _
                                            Mid(StrConv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, vbUnicode), 5, 2) & "/" & _
                                            Mid(StrConv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, vbUnicode), 7, 2)
                   End If
                    '前回注文数
                    If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, vbUnicode)) Then
                        Text1(j + 6).Text = Format(CLng(StrConv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, vbUnicode)), "#0")
                    Else
                        Text1(j + 6).Text = ""
                    End If
                    
                    
                    Text1(j + 7).Text = ""
                    Text1(j + 8).Text = ""
                    
                    Call UniCode_Conv(K1_P_SHORDER.JGYOBU, SHIZAI)      '事業部
                    Call UniCode_Conv(K1_P_SHORDER.NAIGAI, NAIGAI_NAI)  '国内
                                                                        '品番
                    Call UniCode_Conv(K1_P_SHORDER.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                    Call UniCode_Conv(K1_P_SHORDER.ORDER_DT, "zzzzzzzz")
                    Call UniCode_Conv(K1_P_SHORDER.ORDER_NO, "zzzzz")
                                                                                                                    
                    com = BtOpGetLessEqual
                    Do
                        DoEvents
                        sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K1_P_SHORDER, Len(K1_P_SHORDER), 1)
                        Select Case sts
                            Case BtNoErr
                            
                                If StrConv(P_SHORDER_REC.JGYOBU, vbUnicode) <> SHIZAI Or _
                                    StrConv(P_SHORDER_REC.NAIGAI, vbUnicode) <> NAIGAI_NAI Or _
                                    StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode) <> StrConv(ITEMREC.HIN_GAI, vbUnicode) Then
                            
                                    Exit Do
                                End If
                                'If 仕入先／単価追加    2007.09.06
                                If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) <> P_KAN_ON And _
                                    StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) <> P_CANCEL_ON And _
                                    (Trim(StrConv(ITEMREC.G_SHIIRE_TBL(i).CODE, vbUnicode)) = Trim(StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode)) And _
                                    Trim(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA, vbUnicode)) = Trim(StrConv(P_SHORDER_REC.TANKA, vbUnicode))) Then
                                    '予定納期
                                    If Trim(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode)) <> "" Then    '2007.09.06
                                        Text1(j + 7).Text = Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 1, 4) & "/" & _
                                                            Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                                            Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 7, 2)
                                    Else
                                        Text1(j + 7).Text = ""
                                    End If
                                    '注文残
                                    Text1(j + 8).Text = Format(CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) - CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)), "#0")
                                                            
                                    Exit Do
                                End If
                            Case BtErrEOF
                                
                                Exit Do
                            Case Else
                                
                                Call File_Error(sts, BtOpGetLessEqual, "資材注文ﾃﾞｰﾀ")
                                Exit Function
                        
                        End Select
                    
                    
                    
                    
                        com = BtOpGetLess
                    
                    
                    Loop
                
                
                                    '前回注文日
                    If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, vbUnicode)) = "" Then
                        Text1(j + 9).Text = ""
                    Else
                        Text1(j + 9).Text = Mid(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, vbUnicode), 1, 4) & "/" & _
                                            Mid(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, vbUnicode), 5, 2) & "/" & _
                                            Mid(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, vbUnicode), 7, 2)
                   End If

                
                
                    j = j + 10
                
                End If
            
            Next i
        
        
        
            If Text2.Text <> Text1(ptxHIN_GAI).Text Then
                
'                '↓ 2007.08.01
'                If Trim(StrConv(ITEMREC.LAST_CODE, vbUnicode)) = "" Then
'
'                    Call Text1_DblClick(ptxSHIIRE_CODE01)   '2007.11.05
'                Else
'                    j = ptxSHIIRE_CODE01
'                    For i = 0 To 2
'
'                        If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(i).CODE, vbUnicode)) = "" Then
'                            i = 3
'                            Exit For
'                        End If
'
'                        If Trim(StrConv(ITEMREC.LAST_CODE, vbUnicode)) = _
'                            Trim(StrConv(ITEMREC.G_SHIIRE_TBL(i).CODE, vbUnicode)) And _
'                            StrConv(ITEMREC.LAST_TANKA, vbUnicode) = _
'                            StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA, vbUnicode) Then
'
'                            Call Text1_DblClick(j)
'
'                            If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA, vbUnicode)) Then
'                                Text1(j + 2).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA, vbUnicode)), "#0.00")
'                            Else
'                                Text1(j + 2).Text = ""
'                            End If
'
'                            Exit For
'
'                        End If
'
'                        j = j + 9
'
'
'                    Next i
'
'                    If i > 2 Then
'
'                        '注文先ｺｰﾄﾞ
'                        Text1(ptxORDER_CODE).Text = StrConv(ITEMREC.LAST_CODE, vbUnicode)
'                        '注文先名
'                        For i = 0 To Combo1(pcmbORDER).ListCount - 1
'
'                            If Trim(Text1(ptxORDER_CODE).Text) = Trim(Right(Combo1(pcmbORDER).List(i), 5)) Then
'                                Combo1(pcmbORDER).ListIndex = i
'                                Exit For
'                            End If
'
'                        Next i
'                        '単価
'                        If IsNumeric(StrConv(ITEMREC.LAST_TANKA, vbUnicode)) Then
'                            Text1(ptxTANKA).Text = Format(CDbl(StrConv(ITEMREC.LAST_TANKA, vbUnicode)), "#0.00")
'                        Else
'                            Text1(ptxTANKA).Text = ""
'                        End If
'                        '金額
'                        If IsNumeric(Text1(ptxORDER_QTY).Text) And IsNumeric(Text1(ptxTANKA).Text) Then
'                            Text1(ptxKINGAKU).Text = Format(CLng(CDbl(Text1(ptxORDER_QTY).Text) * _
'                                                                CDbl(Text1(ptxTANKA).Text)), "#,##0")
'                        End If
'                        'ﾛｯﾄ数
'                        Text1(ptxLOT).Text = ""
'                        '予定納期
'                        Text1(ptxY_NOUKI_DT).Text = ""
'                    End If
'
'                End If
'                '↑ 2007.08.01
                
                
                svTanka = ""
                j = -1
                For i = 0 To 2
                    If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA, vbUnicode)) = "" Then
                    Else
                        If svTanka < StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, vbUnicode) Then
                            j = i
                            svTanka = StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, vbUnicode)
                        End If
                    End If
                Next i
                
                If j = -1 Then
                    Combo1(pcmbORDER).ListIndex = -1
                    Text1(ptxORDER_CODE).Text = ""
                    Text1(ptxTANKA).Text = ""
                
                Else
                    
                    Select Case j
                        Case 0
                            Call Text1_DblClick(15)
                        Case 1
                            Call Text1_DblClick(25)
                        Case 2
                            Call Text1_DblClick(35)

                    End Select
                End If
                
                
                
                
                Text2.Text = Text1(ptxHIN_GAI).Text
            Else
                If Trim(Text1(ptxSHIIRE_CODE01).Text) <> "" And Trim(Text1(ptxORDER_CODE).Text) = "" Then
'                    '↓ 2007.08.01
'                    If Trim(StrConv(ITEMREC.LAST_CODE, vbUnicode)) = "" Then
'                        Call Text1_DblClick(ptxSHIIRE_CODE01)
'                    Else
'                        j = ptxSHIIRE_CODE01
'                        For i = 0 To 2
'
'                            If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(i).CODE, vbUnicode)) = "" Then
'                                i = 3
'                                Exit For
'                            End If
'
'                            If Trim(StrConv(ITEMREC.LAST_CODE, vbUnicode)) = _
'                                Trim(StrConv(ITEMREC.G_SHIIRE_TBL(i).CODE, vbUnicode)) And _
'                                StrConv(ITEMREC.LAST_TANKA, vbUnicode) = _
'                                StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA, vbUnicode) Then
'
'                                Call Text1_DblClick(j)
'
'                                If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA, vbUnicode)) Then
'                                    Text1(j + 2).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA, vbUnicode)), "#0.00")
'                                Else
'                                    Text1(j + 2).Text = ""
'                                End If
'
'                                Exit For
'
'                            End If
'
'                            j = j + 9
'
'
'                        Next i
'
'                        If i > 2 Then
'
'                            '注文先ｺｰﾄﾞ
'                            Text1(ptxORDER_CODE).Text = StrConv(ITEMREC.LAST_CODE, vbUnicode)
'                            '注文先名
'                            For i = 0 To Combo1(pcmbORDER).ListCount - 1
'
'                                If Trim(Text1(ptxORDER_CODE).Text) = Trim(Right(Combo1(pcmbORDER).List(i), 5)) Then
'                                    Combo1(pcmbORDER).ListIndex = i
'                                    Exit For
'                                End If
'
'                            Next i
'                            '単価
'                            If IsNumeric(StrConv(ITEMREC.LAST_TANKA, vbUnicode)) Then
'                                Text1(ptxTANKA).Text = Format(CDbl(StrConv(ITEMREC.LAST_TANKA, vbUnicode)), "#0.00")
'                            Else
'                                Text1(ptxTANKA).Text = ""
'                            End If
'                            '金額
'                            If IsNumeric(Text1(ptxORDER_QTY).Text) And IsNumeric(Text1(ptxTANKA).Text) Then
'                                Text1(ptxKINGAKU).Text = Format(CLng(CDbl(Text1(ptxORDER_QTY).Text) * _
'                                                                    CDbl(Text1(ptxTANKA).Text)), "#,##0")
'                            End If
'                            'ﾛｯﾄ数
'                            Text1(ptxLOT).Text = ""
'                            '予定納期
'                            Text1(ptxY_NOUKI_DT).Text = ""
'                        End If
'
'                    End If
                    '↑ 2007.08.01
            
            
                    svTanka = ""
                    j = -1
                    For i = 0 To 2
                        If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA, vbUnicode)) = "" Then
                        Else
                            If svTanka < StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, vbUnicode) Then
                                j = i
                                svTanka = StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, vbUnicode)
                            End If
                        End If
                    Next i
                    
                    If j = -1 Then
                        Combo1(pcmbORDER).ListIndex = -1
                        Text1(ptxORDER_CODE).Text = ""
                        Text1(ptxTANKA).Text = ""
                    
                    Else
                        Select Case j
                            Case 0
                                Call Text1_DblClick(15)
                            Case 1
                                Call Text1_DblClick(24)
                            Case 2
                                Call Text1_DblClick(33)
    
                        End Select
    
                    
                    End If
            
            
                    Text2.Text = Text1(ptxHIN_GAI).Text
            
                End If
            End If
        
            Text1(ptxHIN_GAI).SetFocus
                    
        
        Case ptxORDER_CODE   '注文先
            
    
            Text1(Mode).Text = StrConv(RTrim(Text1(Mode).Text), vbUpperCase)        '2017.11.21
            
            
            Combo1(pcmbORDER).ListIndex = -1
            For i = 0 To Combo1(pcmbORDER).ListCount - 1
                If Text1(ptxORDER_CODE).Text = Trim(Right(Combo1(pcmbORDER).List(i), 5)) Then
                    Combo1(pcmbORDER).ListIndex = i
                    Exit For
               End If
           
            Next i
    
            If i > Combo1(pcmbORDER).ListCount - 1 Then
                MsgBox "入力した項目はエラーです。(注文先未登録)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        
                    
        '品目マスタ　登録済の取引先のみOK   2007.11.05
            Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
    
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                    For i = 0 To 2
                        If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(i).CODE, vbUnicode)) = Text1(ptxORDER_CODE).Text Then
                            Exit For
                        End If
                    Next i
                
                    If i > 2 Then
                        MsgBox "入力した項目はエラーです。(注文先未登録)"
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
                
                
                Case BtErrKeyNotFound
                    
                    Text1(ptxHIN_NAME).Text = ""
                    Text1(ptxZAIKO_QTY).Text = ""
            
            
                    For i = 0 To 2
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).CODE, "")
                    Next i
                    
                    
                    MsgBox "入力した項目はエラーです。(品目マスタ未登録)"
                    Text1(Mode).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            
            End Select
        
        
        
        
            For i = ptxSHIIRE_CODE01 To ptxSHIIRE_CODE03 Step 9
                If Text1(ptxORDER_CODE).Text = Text1(i).Text Then
                                    
                    If IsDate(Text1(ptxORDER_DT).Text) And IsNumeric(Text1(i + 4).Text) Then
                    
                        Text1(ptxY_NOUKI_DT).Text = Format(DateAdd("d", CDbl(Text1(i + 4).Text), Text1(ptxORDER_DT).Text), "YYYY/MM/DD")
                    End If
                    '単価
                    Text1(ptxTANKA).Text = Trim(Text1(i + 2).Text)
                    '金額
                    If IsNumeric(Text1(ptxTANKA).Text) And IsNumeric(Text1(ptxORDER_QTY).Text) Then
                        '2009.11.02
'                        Text1(ptxKINGAKU).Text = Format(CLng(CDbl(Text1(ptxTANKA).Text) * CLng(Text1(ptxORDER_QTY).Text)), "#0")
                    
                    
                    
                        Select Case StrConv(P_KANRIREC.SHI_MARUME, vbUnicode)
                            Case "0"    '切捨て
                                Text1(ptxKINGAKU).Text = Format(ToRoundDown(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                                CLng(Text1(ptxORDER_QTY).Text)), 0), "#,##0")
                            
                
                            Case "5"    '四捨五入
                            
                                Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                                CLng(Text1(ptxORDER_QTY).Text)), 0), "#,##0")
                            
                            
                            
                            
                            Case "9"    '切り上げ
                        
                        
                                Text1(ptxKINGAKU).Text = Format(ToRoundUp(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                                CLng(Text1(ptxORDER_QTY).Text)), 0), "#,##0")
                    
                        
                        
                            Case Else    '四捨五入
                            
                                Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                                CLng(Text1(ptxORDER_QTY).Text)), 0), "#,##0")
                        
                        
                        End Select
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    End If
                    Exit For
                End If
            Next i
        
        
        
        
        
        Case ptxDELI_CODE   '納入先
            
            
            Text1(Mode).Text = StrConv(RTrim(Text1(Mode).Text), vbUpperCase)        '2017.11.21
            
            If Trim(Text1(ptxDELI_CODE).Text) = "" Then
                Combo1(pcmbDELI).ListIndex = -1
            Else
            
               Combo1(pcmbDELI).ListIndex = -1
               For i = 0 To Combo1(pcmbDELI).ListCount - 1
                   If Trim(Text1(ptxDELI_CODE).Text) = Trim(Right(Combo1(pcmbDELI).List(i), 5)) Then
                       Combo1(pcmbDELI).ListIndex = i
                       Exit For
                   End If
               
               Next i
        
               If i > Combo1(pcmbDELI).ListCount - 1 Then
                   MsgBox "入力した項目はエラーです。"
                   Text1(Mode).SetFocus
                   Exit Function
               End If
            End If
        
        
        Case ptxORDER_NO        '注文№ 2007.11.13
                
            If Input_Mode = 1 Then
            
                '資材注文データのチェック
                sts = P_SHORDER_Read_Proc(1)
                Select Case sts
                    Case False, BtNoErr
                                
                        If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
                            MsgBox "キャンセルされています。"
                            Text1(Mode).SetFocus
                            Exit Function
                        End If
                    
                        If CLng(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)) <> 0 Then
                            MsgBox "受入実績があります。"
                            Text1(Mode).SetFocus
                            Exit Function
                        End If
                        
                        
                        If StrConv(P_SHORDER_REC.PRINT_F, vbUnicode) = P_PRINT_OFF Then
                            MsgBox "注文書未発行です。"
                            Text1(Mode).SetFocus
                            Exit Function
                        End If
                    
                    Case BtErrKeyNotFound
                        MsgBox "注文№未登録です。"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "資材注文データ")
                        Exit Function
                End Select
            End If
        
        
        Case ptxORDER_QTY       '注文数
    
            If Not IsNumeric(Text1(ptxORDER_QTY).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(Mode).SetFocus
                Exit Function
            Else
                If CLng(Text1(ptxORDER_QTY).Text) = 0 Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
                
                Text1(ptxORDER_QTY).Text = Format(CLng(Text1(ptxORDER_QTY).Text), "#0")
            
            
            '>>>>>>>>>>>>>>>>>>>>>  ロット数による再選択    2017.11.21
                
                If SHIIRE_SELECT <> 1 Then
                
                    SHIIRE_I = -1
                    
                    
                    If Trim(Text1(ptxSHIIRE_LOT01).Text) = "" Then
                        'SHIIRE_LOT01 = "99999"
                        SHIIRE_LOT01 = "999999"
                    Else
                        SHIIRE_LOT01 = Text1(ptxSHIIRE_LOT01).Text
                    End If
                    If Trim(Text1(ptxSHIIRE_LOT02).Text) = "" Then
                        'SHIIRE_LOT02 = "99999"
                        SHIIRE_LOT02 = "999999"
                    Else
                        SHIIRE_LOT02 = Text1(ptxSHIIRE_LOT02).Text
                    End If
                    
                    If Trim(Text1(ptxSHIIRE_LOT03).Text) = "" Then
                        'SHIIRE_LOT03 = "99999"
                        SHIIRE_LOT03 = "999999"
                    Else
                        SHIIRE_LOT03 = Text1(ptxSHIIRE_LOT03).Text
                    End If
                    
                    
                    
                    If Val(SHIIRE_LOT01) < Val(SHIIRE_LOT02) And Val(SHIIRE_LOT01) < Val(SHIIRE_LOT03) Then
                        SHIIRE_I = ptxSHIIRE_LOT01 - 3
                    End If
                    
                    If Val(SHIIRE_LOT02) < Val(SHIIRE_LOT01) And Val(SHIIRE_LOT02) < Val(SHIIRE_LOT03) Then
                        SHIIRE_I = ptxSHIIRE_LOT02 - 3
                    End If
                    
                    If Val(SHIIRE_LOT03) < Val(SHIIRE_LOT01) And Val(SHIIRE_LOT03) < Val(SHIIRE_LOT02) Then
                        SHIIRE_I = ptxSHIIRE_LOT03 - 3
                    End If
                    
                    
                    
                                        
'>>>>>>>>>>     並べ替え　2017.12.07
                    lotLIST.Clear
                    For i = ptxSHIIRE_LOT01 To ptxSHIIRE_LOT03 Step 10
                    
                    
                    
                        
                        If Val(Text1(i).Text) = 0 Then
                            'SHIIRE_LOT = "99999"
                            SHIIRE_LOT = "999999"
                        Else
                            SHIIRE_LOT = Text1(i).Text
                        End If
                    
                    
                        lotLIST.AddItem Format(Val(SHIIRE_LOT), "00000") & i
                    
                    
                    
                    Next i
                    




'>>>>>>>>>>     並べ替え    2017.12.07
                    
                    
                    
                    
                    For i = 0 To 2
                        If Val(Mid(lotLIST.List(i), 1, 6)) > Val(Text1(ptxORDER_QTY).Text) Then '2019/11/1 lot5桁から6桁に変更 注文数6桁ENTERでエラーの為
                            Exit For
                        End If
                    
                    
                    
                        SHIIRE_I = Val(Mid(lotLIST.List(i), 6, 2)) - 3  '2019/11/1 lot5桁から6桁に変更 注文数6桁ENTERでエラーの為
                    Next i
                
                                
                    'If SHIIRE_I <> -1 Then          '2017.11.29
                    If SHIIRE_I > 0 Then         '2017.11.29
                        Call SHIIRE_Disp_Proc(SHIIRE_I)
                    End If
                        
            End If
        '>>>>>>>>>>>>>>>>>>>>>  ロット数による再選択    2017.11.21
            
            
            
            
            
                If IsNumeric(Text1(ptxTANKA).Text) Then
                    
                    Select Case StrConv(P_KANRIREC.SHI_MARUME, vbUnicode)
                        Case "0"    '切捨て
                            Text1(ptxKINGAKU).Text = Format(ToRoundDown(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                            CLng(Text1(ptxORDER_QTY).Text)), 0), "#,##0")
                        
            
                        Case "5"    '四捨五入
                        
                            Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                            CLng(Text1(ptxORDER_QTY).Text)), 0), "#,##0")
                        
                        
                        
                        
                        Case "9"    '切り上げ
                    
                    
                            Text1(ptxKINGAKU).Text = Format(ToRoundUp(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                            CLng(Text1(ptxORDER_QTY).Text)), 0), "#,##0")
                
                    
                    
                        Case Else    '四捨五入
                        
                            Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                            CLng(Text1(ptxORDER_QTY).Text)), 0), "#,##0")
                    
                    
                    End Select
                    
                    
                Else
                    Text1(ptxKINGAKU).Text = "0"
                End If
            End If
    
    
    
    
        
    
        Case ptxY_NOUKI_DT      '納期予定日
        
            If YOTEI_NOUKI Then '2007.09.06
                If Trim(Text1(ptxY_NOUKI_DT).Text) = "" Then
                Else
            
                    If Not IsDate(Text1(ptxY_NOUKI_DT).Text) Then
                        MsgBox "入力した項目はエラーです。"
                        Text1(Mode).SetFocus
                        Exit Function
                    Else
                        Text1(ptxY_NOUKI_DT).Text = Format(CDate(Text1(ptxY_NOUKI_DT).Text), "YYYY/MM/DD")
                    End If
                End If
            Else
                If Not IsDate(Text1(ptxY_NOUKI_DT).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxY_NOUKI_DT).Text = Format(CDate(Text1(ptxY_NOUKI_DT).Text), "YYYY/MM/DD")
                End If
            End If
    
            If OSAKA_MODE Then      '2008.01.10
                If Trim(Text1(ptxUSE_YM).Text) = "" Then
                    Text1(ptxUSE_YM).Text = Left(Text1(ptxY_NOUKI_DT).Text, 7)
                End If
            End If
    
    
        Case ptxUSE_YM      '使用月 2008.01.10
        
            If OSAKA_MODE Then
                If Trim(Text1(ptxUSE_YM).Text) = "" Then
                Else
            
                    If Not IsDate(Text1(ptxUSE_YM).Text & "/01") Then
                        MsgBox "入力した項目はエラーです。"
                        Text1(Mode).SetFocus
                        Exit Function
                    Else
                        Text1(ptxUSE_YM).Text = Left(Format(CDate(Text1(ptxUSE_YM).Text & "/01"), "YYYY/MM/DD"), 7)
                    End If
                End If
            Else
            End If
    
    
    
    
        Case ptxTANKA           '単価
    
            If Not IsNumeric(Text1(ptxTANKA).Text) Then
                MsgBox "入力した項目はエラーです。"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxTANKA).Text = Format(CDbl(Text1(ptxTANKA).Text), "#0.00")
            
                If IsNumeric(Text1(ptxORDER_QTY).Text) Then
                    Select Case StrConv(P_KANRIREC.SHI_MARUME, vbUnicode)
                        Case "0"    '切捨て
                            Text1(ptxKINGAKU).Text = Format(ToRoundDown(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                            CLng(Text1(ptxORDER_QTY).Text)), 0), "#,##0")
                        
            
                        Case "5"    '四捨五入
                        
                            Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                            CLng(Text1(ptxORDER_QTY).Text)), 0), "#,##0")
                        
                        
                        
                        
                        Case "9"    '切り上げ
                    
                    
                            Text1(ptxKINGAKU).Text = Format(ToRoundUp(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                            CLng(Text1(ptxORDER_QTY).Text)), 0), "#,##0")
                
                    
                    
                        Case Else    '四捨五入
                        
                            Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                            CLng(Text1(ptxORDER_QTY).Text)), 0), "#,##0")
                    
                    
                    End Select
                
                End If
            End If
    
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function
Private Function Item_Disp_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   画面表示
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer

Dim Sumi_QTY    As Long
Dim Mi_QTY      As Long

    Item_Disp_Proc = True
    
        
    Text1(ptxTANTO_CODE).Text = StrConv(P_SHORDER_REC.TANTO_CODE, vbUnicode)       '担当者ｺｰﾄﾞ／名称
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTANTO_CODE).Text)

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxTANTO_NAME).Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text1(ptxTANTO_NAME).Text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Exit Function
    
    End Select
                                                                                    '注文日
    Text1(ptxORDER_DT).Text = Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 7, 2)
        
        
    Text1(ptxHIN_GAI).Text = Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))       '品番／品名／在庫残
        
    Text2.Text = Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))                    '2007.09.06
        
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                         
            Call UniCode_Conv(ITEMREC.SHIIRE_BIKOU, "")     '2018.04.19
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    
    End Select
        
        
        
        
        
        
    Text1(ptxHIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    
    lblSHIIRE_BIKOU.Caption = StrConv(ITEMREC.SHIIRE_BIKOU, vbUnicode)      '2018.04.19
    
                                    '月平均出荷数集計データより月平均出荷数獲得 2018.04.19
    Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, Text1(ptxHIN_GAI).Text)
    sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    
    Select Case sts
        Case BtNoErr
        
            txtAVE_SYUKA_cnt.Text = Format(Val(StrConv(AVE_SYUKAREC.TOTAL_AVE_CNT, vbUnicode)), "#,##0.0")
            txtAVE_SYUKA.Text = Format(Val(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode)), "#,##0.0")
                        
                        
                        
                        
                        
        Case BtErrKeyNotFound
        
        
        
        
        Case BtErrKeyNotFound
            txtAVE_SYUKA_cnt.Text = ""
            txtAVE_SYUKA.Text = ""
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "月平均出荷数集計データ")
            Unload Me
    End Select
    
    
    
    
    If Zaiko_Syukei_Proc(Sumi_QTY, Mi_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                            StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
        Exit Function
    
    End If
        
    Text1(ptxZAIKO_QTY).Text = Format(Sumi_QTY + Mi_QTY, "#0")
                                                                                    '注文先
    Text1(ptxORDER_CODE).Text = Trim(StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode))
    Combo1(pcmbORDER).ListIndex = -1
    For i = 0 To Combo1(pcmbORDER).ListCount - 1
    
        If Text1(ptxORDER_CODE).Text = Trim(Right(Combo1(pcmbORDER).List(i), 5)) Then
            Combo1(pcmbORDER).ListIndex = i
            Exit For
        End If
    
    Next i
                                                                                    
                                                                                    
                                                                                    
                                                                                    
                                                                                    '納入先
    Text1(ptxDELI_CODE).Text = Trim(StrConv(P_SHORDER_REC.DELI_CODE, vbUnicode))
    Combo1(pcmbDELI).ListIndex = -1
    For i = 0 To Combo1(pcmbDELI).ListCount - 1
    
        If Text1(ptxDELI_CODE).Text = Trim(Right(Combo1(pcmbDELI).List(i), 5)) Then
            Combo1(pcmbDELI).ListIndex = i
            Exit For
        End If
    
    Next i
                                                                                    '注文№
    Text1(ptxORDER_NO).Text = Trim(StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode))
                                                                                    '注文数
    If Mode = 0 Then    '2007.11.12
        Text1(ptxORDER_QTY).Text = Format(CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)), "#0")
    
                                                                                        
        
        If IsNumeric(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)) Then                                     '2016.06.22
            Text1(ptxUKEIRE_QTY).Text = Format(CLng(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)), "#0")    '2016.06.22
        Else                                                                                                '2016.06.22
            Text1(ptxUKEIRE_QTY).Text = 0                                                                   '2016.06.22
        End If                                                                                              '2016.06.22

                                                                                        
                                                                                        
                                                                                        
                                                                                        '単価
        Text1(ptxTANKA).Text = Format(CDbl(StrConv(P_SHORDER_REC.TANKA, vbUnicode)), "#0.00")
                                                                                        '金額
        
                
        
        
        
        
        
        
        '2009.11.02
'        Text1(ptxKINGAKU).Text = Format(CDbl(StrConv(P_SHORDER_REC.TANKA, vbUnicode)) * _
'                                        CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)), "#,##0")
        
        
        Select Case StrConv(P_KANRIREC.SHI_MARUME, vbUnicode)
            Case "0"    '切捨て
                Text1(ptxKINGAKU).Text = Format(ToRoundDown(CCur(CCur(StrConv(P_SHORDER_REC.TANKA, vbUnicode)) * _
                                                CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))), 0), "#,##0")
            

            Case "5"    '四捨五入
            
                Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(StrConv(P_SHORDER_REC.TANKA, vbUnicode)) * _
                                                CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))), 0), "#,##0")
            
            
            
            
            Case "9"    '切り上げ
        
        
                Text1(ptxKINGAKU).Text = Format(ToRoundUp(CCur(CCur(StrConv(P_SHORDER_REC.TANKA, vbUnicode)) * _
                                                CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))), 0), "#,##0")
    
    
    
            Case Else   '四捨五入
            
                Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(StrConv(P_SHORDER_REC.TANKA, vbUnicode)) * _
                                                CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))), 0), "#,##0")
    
    
        End Select
    
    
    
    
    
    
    
    Else
        If Trim(Text1(ptxORDER_QTY).Text) = "" Then
            Text1(ptxORDER_QTY).Text = Format(CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)), "#0")
        
                                                                                        '単価
            Text1(ptxTANKA).Text = Format(CDbl(StrConv(P_SHORDER_REC.TANKA, vbUnicode)), "#0.00")
                                                                                            '金額
            '2009.11.02
'            Text1(ptxKINGAKU).Text = Format(CDbl(StrConv(P_SHORDER_REC.TANKA, vbUnicode)) * _
'                                            CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)), "#,##0")
        
        
        Select Case StrConv(P_KANRIREC.SHI_MARUME, vbUnicode)
            Case "0"    '切捨て
                Text1(ptxKINGAKU).Text = Format(ToRoundDown(CCur(CCur(StrConv(P_SHORDER_REC.TANKA, vbUnicode)) * _
                                                CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))), 0), "#,##0")
            

            Case "5"    '四捨五入
            
                Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(StrConv(P_SHORDER_REC.TANKA, vbUnicode)) * _
                                                CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))), 0), "#,##0")
            
            
            
            
            Case "9"    '切り上げ
        
        
                Text1(ptxKINGAKU).Text = Format(ToRoundUp(CCur(CCur(StrConv(P_SHORDER_REC.TANKA, vbUnicode)) * _
                                                CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))), 0), "#,##0")
    
        
        
            Case Else   '四捨五入
            
                Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(StrConv(P_SHORDER_REC.TANKA, vbUnicode)) * _
                                                CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))), 0), "#,##0")
        
        End Select
        
        
        
        
        
        
        
        
        
        
        
        
        End If
    End If
                                                                                    '納期予定日
    
    If Trim(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode)) <> "" Then                '2007.09.06
        Text1(ptxY_NOUKI_DT).Text = Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 7, 2)
    Else
        Text1(ptxY_NOUKI_DT).Text = ""
    End If
                                                                                        
                                                                                        
                                                                                        
                                                                                    '使用月 2008.01.10
    If Trim(StrConv(P_SHORDER_REC.USE_YM, vbUnicode)) <> "" Then
        Text1(ptxUSE_YM).Text = Mid(StrConv(P_SHORDER_REC.USE_YM, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SHORDER_REC.USE_YM, vbUnicode), 5, 2)
    Else
        Text1(ptxUSE_YM).Text = ""
    End If
                                                                                        
                                                                                        '単価
'    Text1(ptxTANKA).Text = Format(CDbl(StrConv(P_SHORDER_REC.TANKA, vbUnicode)), "#0.00")
                                                                                    '金額
'    Text1(ptxKINGAKU).Text = Format(CDbl(StrConv(P_SHORDER_REC.TANKA, vbUnicode)) * _
'                                    CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)), "#,##0")
                                                                                    '発注ﾛｯﾄ
    Text1(ptxLOT).Text = Format(CLng(StrConv(P_SHORDER_REC.LOT, vbUnicode)), "#0")
                                                                                            
                                                            
                                                            
                                                            
    '優先仕入先表示
    j = ptxSHIIRE_CODE01
    For i = 0 To 2
            
        If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(i).CODE, vbUnicode)) = "" Then
        Else
                
            Text1(j).Text = StrConv(ITEMREC.G_SHIIRE_TBL(i).CODE, vbUnicode)
            Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(j).Text)
                
            sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
            Select Case sts
                Case BtNoErr
                
                                            
                
                Case BtErrKeyNotFound
                    
                    Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
            
            
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
                    Exit Function
            
            End Select
            '仕入先名
            Text1(j + 1).Text = StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode)
            '単価
            If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA, vbUnicode)) Then
                Text1(j + 2).Text = Format(CDbl(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA, vbUnicode)), "#0.00")
            Else
                Text1(j + 2).Text = ""
            End If
            'ﾛｯﾄ数
            If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(i).LOT, vbUnicode)) Then
                Text1(j + 3).Text = Format(CLng(StrConv(ITEMREC.G_SHIIRE_TBL(i).LOT, vbUnicode)), "#0")
            Else
                Text1(j + 3).Text = ""
            End If
            'ﾘｰﾄﾞﾀｲﾑ
            If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(i).LEAD_TIME, vbUnicode)) Then
                Text1(j + 4).Text = Format(CLng(StrConv(ITEMREC.G_SHIIRE_TBL(i).LEAD_TIME, vbUnicode)), "#0")
            Else
                Text1(j + 4).Text = ""
            End If
            '前回注文日
            If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, vbUnicode)) = "" Then
                Text1(j + 5).Text = ""
            Else
                Text1(j + 5).Text = Mid(StrConv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, vbUnicode), 7, 2)
            End If
            '前回注文数
            If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, vbUnicode)) Then
                Text1(j + 6).Text = Format(CLng(StrConv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, vbUnicode)), "#0")
            Else
                Text1(j + 6).Text = ""
            End If
                    
                    
            Text1(j + 7).Text = ""
            Text1(j + 8).Text = ""
            
            Call UniCode_Conv(K1_P_SHORDER.JGYOBU, SHIZAI)      '事業部
            Call UniCode_Conv(K1_P_SHORDER.NAIGAI, NAIGAI_NAI)  '国内
                                                                '品番
            Call UniCode_Conv(K1_P_SHORDER.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
            Call UniCode_Conv(K1_P_SHORDER.ORDER_DT, "zzzzzzzz")
            Call UniCode_Conv(K1_P_SHORDER.ORDER_NO, "zzzzz")
                                                                                                            
            com = BtOpGetLessEqual
            Do
                DoEvents
                sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K1_P_SHORDER, Len(K1_P_SHORDER), 1)
                Select Case sts
                    Case BtNoErr
                    
                        If StrConv(P_SHORDER_REC.JGYOBU, vbUnicode) <> SHIZAI Or _
                            StrConv(P_SHORDER_REC.NAIGAI, vbUnicode) <> NAIGAI_NAI Or _
                            StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode) <> StrConv(ITEMREC.HIN_GAI, vbUnicode) Then
                    
                            Exit Do
                        End If
                    
                                'If 仕入先／単価追加    2007.09.06
                        If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) <> P_KAN_ON And _
                            StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) <> P_CANCEL_ON And _
                            (Trim(StrConv(ITEMREC.G_SHIIRE_TBL(i).CODE, vbUnicode)) = Trim(StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode)) And _
                            Trim(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA, vbUnicode)) = Trim(StrConv(P_SHORDER_REC.TANKA, vbUnicode))) Then
                            
                            
                            '予定納期
                            If Trim(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode)) <> "" Then    '2007.09.06
                                Text1(j + 7).Text = Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 1, 4) & "/" & _
                                                    Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                                    Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 7, 2)
                            Else
                                Text1(j + 7).Text = ""
                            End If
                            '注文残
                            Text1(j + 8).Text = Format(CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) - CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)), "#0")
                                                    
                            Exit Do
                        End If
                    Case BtErrEOF
                        
                        Exit Do
                    Case Else
                        
                        Call File_Error(sts, BtOpGetLessEqual, "資材注文ﾃﾞｰﾀ")
                        Exit Function
                
                End Select
            
            
            
                com = BtOpGetLess
            
            Loop
        
            '単価設定日
            If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, vbUnicode)) = "" Then
                Text1(j + 9).Text = ""
            Else
                Text1(j + 9).Text = Mid(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, vbUnicode), 7, 2)
            End If
        
        
            j = j + 10
        
        End If
    
    Next i
                                                            
                                                            
                                                            
    
    Item_Disp_Proc = False

End Function

Private Function Cancel_Proc() As Integer
'----------------------------------------------------------------------------
'                  資材注文ﾃﾞｰﾀｷｬﾝｾﾙ更新
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim ans         As Integer
Dim com         As Integer

Dim SEQNO       As Integer



Dim i           As Integer


    Cancel_Proc = True
                                        
                                        
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If

    
    
    
    '---------------------------------------------------    '資材注文ﾃﾞｰﾀ削除
    
    Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, Text1(ptxORDER_NO).Text)
    
    Do
    
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
            
        Select Case sts
            Case BtNoErr
                com = BtOpDelete
                
                Exit Do
            
            Case BtErrKeyNotFound
                com = 0
            
                            
                Exit Do
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。< P_SHORDER.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            
            
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "資材注文ﾃﾞｰﾀ")
                GoTo Abort_Tran
        End Select

    Loop
    
    
    If com = BtOpDelete Then
    
        Do
            
            DoEvents
            
            sts = BTRV(BtOpDelete, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_SHORDER.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        If com = BtOpUpdate Then
                            sts = BTRV(BtOpUnlock, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
                            If sts Then
                                Call File_Error(sts, BtOpUnlock, "資材注文ﾃﾞｰﾀ")
                            End If
                        End If
                        GoTo Abort_Tran
                    End If
                Case Else
                    Call File_Error(sts, BtOpDelete, "資材注文ﾃﾞｰﾀ")
                    GoTo Abort_Tran
            End Select
        
        Loop
    
    End If
End_Tran:
                                        
    '>>>>>>>>>>>>>>>>   LOG　2016.04.25
    If com = BtOpDelete Then
        If PI00030_LOG <> "" Then
            Call LOG_OUT(PI00030_LOG, "<CANCEL> ORDER_No." & StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode) & " 注文日:" & StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode) & " 品番:" & StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode) & " 注文先:" & StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode) & " 注文数:" & Val(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)))
        End If
    End If
    '>>>>>>>>>>>>>>>>   LOG　2016.04.25
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    Cancel_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    
    Call Input_UnLock

End Function
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                  資材注文ﾃﾞｰﾀ更新
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer

'Dim ORDERNO         As Integer
Dim ORDERNO         As Long             '2017.10.13



Dim i               As Integer
Dim j               As Integer

Dim Min_Order_DT    As String * 8
Dim Save_I          As Integer



    Update_Proc = True
                                        
    Call Input_Lock
                                        
    DoEvents
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If

    If Text1(ptxORDER_NO).Text = "" Then
                                        
                                            
        Do                                              '2013.10.08
            DoEvents                                    '2013.10.08
                                            
                                            
                                            
                                            
                                                '管理ファイルより資材注文番号の獲得
            Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
            
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound
                    
                        If P_KANRI_MAKE_Proc() Then
                            GoTo Abort_Tran
                        End If
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE         'これは無い
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Update_Proc = True
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "管理マスタ")
                        GoTo Abort_Tran
                
                End Select
            
            
            Loop
        
            '指図票№＋１
        
        
            If CLng(StrConv(P_KANRIREC.ORDER_NO, vbUnicode)) = 99999 Then
                Call UniCode_Conv(P_KANRIREC.ORDER_NO, "00001")
            Else
                
                Call UniCode_Conv(P_KANRIREC.ORDER_NO, Format(CLng(StrConv(P_KANRIREC.ORDER_NO, vbUnicode)) + 1, "00000"))
            End If
        
        
            Do
                
                DoEvents
                
                sts = BTRV(BtOpUpdate, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            sts = BTRV(BtOpUnlock, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                            If sts Then
                                Call File_Error(sts, BtOpUnlock, "管理マスタ")
                            End If
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "管理マスタ")
                        GoTo Abort_Tran
                End Select
            Loop
    
            ORDERNO = CLng(StrConv(P_KANRIREC.ORDER_NO, vbUnicode))
    
    
                    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< 2013.10.08 注文ﾃﾞｰﾀのﾁｪｯｸ
            Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, StrConv(P_KANRIREC.ORDER_NO, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Exit Do
                Case Else
                    Call File_Error(sts, BtOpUpdate, "管理マスタ")
                    GoTo Abort_Tran
            End Select
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< 2013.10.08 注文ﾃﾞｰﾀのﾁｪｯｸ
        Loop                                            '2013.10.08
    
    Else
        
        ORDERNO = CLng(Text1(ptxORDER_NO).Text)
    
    End If

    
    
    
    '---------------------------------------------------    '資材注文データ更新
    Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, Text1(ptxORDER_NO).Text)
    
    Do
    
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
            
        Select Case sts
            Case BtNoErr
            
                com = BtOpUpdate
                Exit Do
            
            Case BtErrKeyNotFound
            
                com = BtOpInsert
                Exit Do
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_SHORDER.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            
            
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "資材注文ﾃﾞｰﾀ")
                GoTo Abort_Tran
        End Select

    Loop
    
    If com = BtOpInsert Then
        Call UniCode_Conv(P_SHORDER_REC.ORDER_NO, Format(ORDERNO, "00000"))     '注文№
    
        Call UniCode_Conv(P_SHORDER_REC.KAN_F, P_KAN_OFF)                       '完了ﾌﾗｸﾞ
        Call UniCode_Conv(P_SHORDER_REC.KAN_DT, "")                             '完了日
        Call UniCode_Conv(P_SHORDER_REC.BUNNOU_CNT, "00")                       '受入回数
        Call UniCode_Conv(P_SHORDER_REC.UKEIRE_QTY, "00000000.00")              '受入数
    
    
        Call UniCode_Conv(P_SHORDER_REC.ANS_NOUKI_DT, "")                       '回答納期日 2016.01.26
    
    
        Call UniCode_Conv(P_SHORDER_REC.CANCEL_F, P_CANCEL_OFF)                 'ｷｬﾝｾﾙﾌﾗｸﾞ
        Call UniCode_Conv(P_SHORDER_REC.CANCEL_DATETIME, "")                    'ｷｬﾝｾﾙ日時
    
        Call UniCode_Conv(P_SHORDER_REC.PRINT_F, P_PRINT_OFF)                   '印刷ﾌﾗｸﾞ
    
        Call UniCode_Conv(P_SHORDER_REC.WS_NO, WS_NO)                           '入力端末
    
        Call UniCode_Conv(P_SHORDER_REC.FILLER, "")
    
    End If
    
    
    If REPRINT_FLG And com = BtOpUpdate Then                                    '2013.02.14
        Call UniCode_Conv(P_SHORDER_REC.PRINT_F, P_PRINT_OFF)                   '印刷ﾌﾗｸﾞ
    End If                                                                      '2013.02.14
    
    Call UniCode_Conv(P_SHORDER_REC.WS_NO, WS_NO)                           '入力端末
        
        
                                                                                '注文日
    Call UniCode_Conv(P_SHORDER_REC.ORDER_DT, Format(Text1(ptxORDER_DT).Text, "YYYYMMDD"))
    
    
    
    Call UniCode_Conv(P_SHORDER_REC.TANTO_CODE, Text1(ptxTANTO_CODE).Text)      '担当者
    Call UniCode_Conv(P_SHORDER_REC.JGYOBU, SHIZAI)                             '事業部（＝資材）
    Call UniCode_Conv(P_SHORDER_REC.NAIGAI, NAIGAI_NAI)                         '国内外
    Call UniCode_Conv(P_SHORDER_REC.HIN_GAI, Text1(ptxHIN_GAI).Text)            '品番
    Call UniCode_Conv(P_SHORDER_REC.ORDER_CODE, Text1(ptxORDER_CODE).Text)      '注文先ｺｰﾄﾞ
    Call UniCode_Conv(P_SHORDER_REC.DELI_CODE, Text1(ptxDELI_CODE).Text)        '納入先ｺｰﾄﾞ
    Call UniCode_Conv(P_SHORDER_REC.ORDER_QTY, Format(CDbl(Text1(ptxORDER_QTY).Text), _
                                                                "00000000.00")) '注文数
    
    
    
    
    
    If Trim(Text1(ptxY_NOUKI_DT).Text) <> "" Then                               '2007.09.06
        Call UniCode_Conv(P_SHORDER_REC.Y_NOUKI_DT, Format(CDate(Text1(ptxY_NOUKI_DT).Text), _
                                                                    "YYYYMMDD"))    '予定納期
    Else
        Call UniCode_Conv(P_SHORDER_REC.Y_NOUKI_DT, "")
    End If


'    Call UniCode_Conv(P_SHORDER_REC.ANS_NOUKI_DT, "")                           '回答納期日    2016.01.26
    
    If Trim(Text1(ptxUSE_YM).Text) <> "" Then                               '2007.12.10
        Call UniCode_Conv(P_SHORDER_REC.USE_YM, Format(CDate(Text1(ptxUSE_YM).Text & "/01"), _
                                                                    "YYYYMMDD"))    '使用月
    Else
        Call UniCode_Conv(P_SHORDER_REC.USE_YM, "")
    End If

    
    Call UniCode_Conv(P_SHORDER_REC.TANKA, Format(CDbl(Text1(ptxTANKA).Text), _
                                                                "00000000.00")) '単価
    
    If IsNumeric(Text1(ptxLOT).Text) Then
        Call UniCode_Conv(P_SHORDER_REC.LOT, Format(CDbl(Text1(ptxLOT).Text), _
                                                                    "00000000"))    'ﾛｯﾄ
    Else
        Call UniCode_Conv(P_SHORDER_REC.LOT, "00000001")
    End If


    '品目ﾏｽﾀ読込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            MsgBox "品目マスタが他端末で変更されました。更新処理を中止します。"
            GoTo Abort_Tran
        Case Else
            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目ﾏｽﾀ")
            GoTo Abort_Tran
    End Select
    '仕入区分
    Call UniCode_Conv(P_SHORDER_REC.G_SHIIRE_KBN, StrConv(ITEMREC.G_SHIIRE_KBN, vbUnicode))
    '収支単位
    Call UniCode_Conv(P_SHORDER_REC.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))


    '受払先ﾏｽﾀ読込み
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxORDER_CODE).Text)
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            MsgBox "受払先マスタが他端末で変更されました。更新処理を中止します。"
            GoTo Abort_Tran
        Case Else
            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "受払先ﾏｽﾀ")
            GoTo Abort_Tran
    End Select

                                                                                '取引先区分
    Call UniCode_Conv(P_SHORDER_REC.TORI_KBN, StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode))

    Call UniCode_Conv(P_SHORDER_REC.FILLER, "")
                                                                                '更新日時
    Call UniCode_Conv(P_SHORDER_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    Do
        
        DoEvents
        
        sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_SHORDER.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    If com = BtOpUpdate Then
                        sts = BTRV(BtOpUnlock, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "資材注文ﾃﾞｰﾀ")
                        End If
                    End If
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, com, "資材注文ﾃﾞｰﾀ")
                GoTo Abort_Tran
        End Select
    
    Loop
    
    
    '---------------------------------------------------    '品目マスタ更新
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
    
    Do
    
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            
        Select Case sts
            Case BtNoErr
            
                Exit Do
            
            Case BtErrKeyNotFound
            
                MsgBox "品目マスタが削除されています。更新を中止します。"
                GoTo Abort_Tran
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            
            
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                GoTo Abort_Tran
        End Select

    Loop
    
    For i = 0 To 2
    
        If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(i).CODE, vbUnicode)) = Trim(StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode)) And _
            StrConv(ITEMREC.G_SHIIRE_TBL(i).TANKA, vbUnicode) = StrConv(P_SHORDER_REC.TANKA, vbUnicode) Then '2007.09.06
            Exit For
        End If
    Next i
    
    
    If i <= 2 Then
        '前回注文日
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode))
        '前回注文数
        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))
    End If

    '最新仕入先コード   2007.05.28
    Call UniCode_Conv(ITEMREC.LAST_CODE, Text1(ptxORDER_CODE).Text)
    '最新仕入単価       2007.05.28
    Call UniCode_Conv(ITEMREC.LAST_TANKA, Format(CDbl(Text1(ptxTANKA).Text), _
                                                                "00000000.00"))


    Do
        
        DoEvents
        
        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    If sts Then
                        Call File_Error(sts, BtOpUnlock, "品目ﾏｽﾀ")
                    End If
                End If
                GoTo Abort_Tran
            Case Else
                Call File_Error(sts, BtOpUpdate, "品目ﾏｽﾀ")
                GoTo Abort_Tran
        End Select
    
    Loop



End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
    Call Input_UnLock
    
    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    
    Call Input_UnLock

End Function


Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    Select Case Index
        Case pcmbORDER          '注文先
            Text1(ptxORDER_CODE).Text = Trim(Right(Combo1(pcmbORDER).Text, 5))
        Case pcmbDELI           '納入先
            Text1(ptxDELI_CODE).Text = Trim(Right(Combo1(pcmbDELI).Text, 5))
    End Select
    
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    
    Select Case Index
        Case pcmbORDER          '注文先
            Text1(ptxORDER_CODE).Text = Trim(Right(Combo1(pcmbORDER).Text, 5))
        Case pcmbDELI           '納入先
            Text1(ptxDELI_CODE).Text = Trim(Right(Combo1(pcmbDELI).Text, 5))
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans         As Integer
Dim i           As Integer

Dim rpt         As New PI00030F1
Dim f           As New PI000302

Dim sts         As Integer

    Select Case Index
        Case P_CMD_Upd        '更新
            
            
            Select Case Input_Mode
                Case 0
                    For i = ptxTANTO_CODE To ptxORDER_ZAN03
                    
                        If i <> ptxORDER_QTY Then   '2017.11.21
                            If Error_Check_Proc(i) Then     'エラーチェック
                                Exit Sub
                            End If
                        End If                      '2017.11.21
                    
                    Next i
                Case 1
                    For i = ptxORDER_NO To ptxORDER_QTY
                    
                        If Error_Check_Proc(i) Then     'エラーチェック
                            Exit Sub
                        End If
                    
                    Next i
            End Select
            
            
            '>>>>>>>>>>>>>>>    注文書発行済みのチェック    2016.04.25
            If Input_Mode <> 1 Then
                
                
                If Trim(Text1(ptxORDER_NO).Text) <> "" Then
                
                                    
                    Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, Text1(ptxORDER_NO))
                    sts = BTRV(BtOpGetEqual, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
                        
                    Select Case sts
                        Case BtNoErr
                            If StrConv(P_SHORDER_REC.PRINT_F, vbUnicode) <> P_PRINT_OFF Then
                                MsgBox "注文書発行済みの為、注文数以外の変更は出来ません。注文数を変更する時は、「注文数」を押下して下さい。"
                                Exit Sub
                            End If
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "注文データ")
                            Unload Me
                    
                    End Select
                
                
                End If
                
                
            End If
            '>>>>>>>>>>>>>>>    注文書発行済みのチェック    2016.04.25
            
            
            
            
            Beep
            ans = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
                
                If List_Disp_Proc() Then
                    Unload Me
                End If
                
                If Init_Proc() Then
                    Unload Me
                End If
            
            
                If Input_Mode = 1 Then  '2007.11.12
                
                    Text1(ptxTANTO_CODE).Text = ""
                    Text1(ptxTANTO_NAME).Text = ""
                    Text1(ptxORDER_DT).Text = ""
                End If
            
                Set Z_SHORDER = Nothing
                Set TDBGrid2.Array = Z_SHORDER
                        
                        
                TDBGrid2.ReBind
                TDBGrid2.Update
                TDBGrid2.MoveFirst
                
                Z_Tbl_Set_F = False
            
            
            
            
            
            End If
            
            '========================================================= 2007/03/19 =====
''            Text1(ptxTANTO_CODE).SetFocus
            
            If Input_Mode = 0 Then  '2007.11.12
                Text1(ptxHIN_GAI).SetFocus
            Else
                Text1(ptxORDER_NO).SetFocus
            End If
            '==========================================================================
        
        Case P_CMD_DEL                      '削除
        
            If Input_Mode = 1 Then  '2007.11.12
                MsgBox "注文数変更モードです。処理を切り替えてください。"
                Exit Sub
            End If
            
            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2016.03.08
'            If CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)) <> 0 Then        '2016.06.22
            
            
Debug.Print CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode))
            If Val(Text1(ptxUKEIRE_QTY).Text) <> 0 Then                             '2016.06.22
                MsgBox "受入実績が有るので取り消し出来ません"
                Exit Sub
            End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2016.03.08
            
            
            
            
            
            Beep
            ans = MsgBox("取り消しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                If Cancel_Proc() Then
                    Unload Me
                End If
                
                If Init_Proc() Then
                    Unload Me
                End If
            
                If List_Disp_Proc() Then
                    Unload Me
                End If
            
            
                Set Z_SHORDER = Nothing
                Set TDBGrid2.Array = Z_SHORDER
                        
                        
                TDBGrid2.ReBind
                TDBGrid2.Update
                TDBGrid2.MoveFirst
                
                Z_Tbl_Set_F = False
            
            
            
            End If
            
            Text1(ptxTANTO_CODE).SetFocus
    
        Case P_CMD_DSP                      '検索/表示
            
            Select Case Input_Mode
                Case 0
                    Input_Mode = 1
                    Command1(4).Caption = "通常"
                    If Init_Proc() Then
                        Unload Me
                    End If
                    
                    
                    Text1(ptxTANTO_CODE).Text = ""
                    Text1(ptxTANTO_NAME).Text = ""
                    Text1(ptxORDER_DT).Text = ""
                    
                    
                    
                    Set Z_SHORDER = Nothing
                    Set TDBGrid2.Array = Z_SHORDER
                            
                            
                    TDBGrid2.ReBind
                    TDBGrid2.Update
                    TDBGrid2.MoveFirst
                    
                    Z_Tbl_Set_F = False
                    
                    
                    Text1(ptxORDER_NO).SetFocus
                Case 1
                    Command1(4).Caption = "注文数"
                    Input_Mode = 0
                    If Init_Proc() Then
                        Unload Me
                    End If
                    
                    
                    
                    
                    
                    Text1(ptxHIN_GAI).SetFocus
            
            End Select
        
        
        
        
        Case 7                              '消去　2016.04.25
        
            If Init_Proc() Then
                Unload Me
            End If
        
        
            If Input_Mode = 1 Then  '2007.11.12
            
                Text1(ptxTANTO_CODE).Text = ""
                Text1(ptxTANTO_NAME).Text = ""
                Text1(ptxORDER_DT).Text = ""
            End If
        
            Set Z_SHORDER = Nothing
            Set TDBGrid2.Array = Z_SHORDER
                    
                    
            TDBGrid2.ReBind
            TDBGrid2.Update
            TDBGrid2.MoveFirst
            
            Z_Tbl_Set_F = False
        
        
            Text1(ptxTANTO_CODE).SetFocus
        
        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        
        Case P_CMD_PRT                      '印刷
 
            If Input_Mode = 1 Then  '2007.11.12
                MsgBox "注文数変更モードです。処理を切り替えてください。"
                Exit Sub
            End If
            
            
            If Not Tbl_Set_F Then
                MsgBox "印刷対象がありません。"
            Else
            
                
                If Print_Proc() Then
                    Unload Me
                End If
                
                
                
                Set SHORDER = Nothing
                Set TDBGrid1.Array = SHORDER
                        
                        
                TDBGrid1.ReBind
                TDBGrid1.Update
                TDBGrid1.MoveFirst
                
                Tbl_Set_F = False
            
            
                Set Z_SHORDER = Nothing
                Set TDBGrid2.Array = Z_SHORDER
                        
                        
                TDBGrid2.ReBind
                TDBGrid2.Update
                TDBGrid2.MoveFirst
                
                Z_Tbl_Set_F = False
            
            End If
            
'            Text1(ptxTANTO_CODE).SetFocus
            
            
        Case P_CMD_End                      '終了
    
                        
            If Tbl_Set_F Then
                ans = MsgBox("印刷処理が行われていません。入力処理に戻りますか？", vbYesNo + vbQuestion, "確認入力")
                If ans = vbYes Then
                Else
                    Unload Me
                End If
            Else
                Unload Me
            End If
    End Select

End Sub


Private Sub Form_DblClick()
'    PrintForm          2017.11.17
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command1(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

Dim c       As String * 128
Dim sts     As Integer
Dim i       As Integer

Dim TUKI    As Integer          '2018.04.19


Dim sBuffer As String

'    If App.PrevInstance Then                       '2017.11.21
'        Beep                                       '2017.11.21
'        MsgBox "同一プログラム実行中です。"        '2017.11.21
'        End                                        '2017.11.21
'    End If                                         '2017.11.21

    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        WS_NO = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        WS_NO = "???"
    End If


    PI000301.Caption = PI000301.Caption & Last_Update_Day   '2017.11.21

                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                
                                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> P_SYS--> PI0030 2016.04.25
                                
                                '納入先取り込み
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "NOUNYU", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "NOUNYU", App.EXEName, c) Then
    Else
        NOUNYU = RTrim(c)
    End If
                                
                                
                                '備考１取り込み '2007.07.20
'    If GetIni(App.EXEName, "BIKOU_1", "P_SYS", c) Then
    If GetIni(App.EXEName, "BIKOU_1", App.EXEName, c) Then
        pubBikou_1 = ""
    Else
        pubBikou_1 = Trim(c)
    End If
                                '備考２取り込み '2007.07.20
'    If GetIni(App.EXEName, "BIKOU_2", "P_SYS", c) Then
    If GetIni(App.EXEName, "BIKOU_2", App.EXEName, c) Then
        pubBikou_2 = ""
    Else
        pubBikou_2 = Trim(c)
    End If
                                '備考３取り込み '2007.07.20
'    If GetIni(App.EXEName, "BIKOU_3", "P_SYS", c) Then
    If GetIni(App.EXEName, "BIKOU_3", App.EXEName, c) Then
        pubBikou_3 = ""
    Else
        pubBikou_3 = Trim(c)
    End If
                                
                                
                                '予定納期の省略可否 '2007.09.06
'    If GetIni(App.EXEName, "YOTEI_NOUKI", "P_SYS", c) Then
    If GetIni(App.EXEName, "YOTEI_NOUKI", App.EXEName, c) Then
        YOTEI_NOUKI = False
    Else
        
        If Not IsNumeric(Trim(c)) Then
            YOTEI_NOUKI = False
        Else
                
            If Trim(c) = "1" Then
                YOTEI_NOUKI = True
            Else
                YOTEI_NOUKI = False
            End If
        End If
    End If
                                
                                
                                '使用月の入力有無 '2008.01.10
'    If GetIni(App.EXEName, "OSAKA_MODE", "P_SYS", c) Then
    If GetIni(App.EXEName, "OSAKA_MODE", App.EXEName, c) Then
        OSAKA_MODE = False
    Else
        
        If Not IsNumeric(Trim(c)) Then
            OSAKA_MODE = False
        Else
                
            If Trim(c) = "1" Then
                OSAKA_MODE = True
            Else
                OSAKA_MODE = False
            End If
        End If
    End If
                                
                                
                                'キャンセル時のログ 2016.04.25
    If GetIni(App.EXEName, "PI00030_LOG", App.EXEName, c) Then
        PI00030_LOG = ""
    Else
        PI00030_LOG = Trim(c)
    End If
                                
                                
    Label1(plblUSE_YM).Visible = OSAKA_MODE
    Text1(ptxUSE_YM).Visible = OSAKA_MODE
    Text1(ptxUSE_YM).TabStop = OSAKA_MODE
                                    
    TDBGrid2.Columns(colZ_USE_YM).Visible = OSAKA_MODE
    TDBGrid2.Columns(colZ_ANS_NOUKI_DT).Visible = OSAKA_MODE
                                '使用月の入力有無 '2008.01.10
                                
                                
                                
                                '注文書　再発行　有無   2013.02.15
    If GetIni(App.EXEName, "REPRINT_FLG", App.EXEName, c) Then
'    If GetIni(App.EXEName, "REPRINT_FLG", "P_SYS", c) Then
        REPRINT_FLG = False
    Else
        If Trim(c) = "1" Then
            REPRINT_FLG = True
        Else
            REPRINT_FLG = False
        End If
    End If
                                
                                
                                
                                '最大表示行   2017.11.21
                                
    If GetIni(App.EXEName, "LIST_MAX", App.EXEName, c) Then
        LIST_MAX = 0
    Else
        LIST_MAX = Val(c)
    End If
                                
                                '仕入選択   2017.11.22
                                
    If GetIni(App.EXEName, "SHIIRE_SELECT", App.EXEName, c) Then
        SHIIRE_SELECT = 0
    Else
        SHIIRE_SELECT = Val(c)
    End If
                                
                                
                                '月平均出荷数/件数算出処理   2018.04.19
                                
    If GetIni(App.EXEName, "TUKI", App.EXEName, c) Then
        TUKI = 3
    Else
        TUKI = Val(Trim(c))
    End If
    lblTuki(0).Caption = "(" & Format(TUKI, "#0") & "ｶ月)"
    lblTuki(1).Caption = "(" & Format(TUKI, "#0") & "ｶ月)"
                                
                                
                                
                                
                                
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '管理マスタＯＰＥＮ
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '受払先マスタＯＰＥＮ
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '資材注文ﾃﾞｰﾀＯＰＥＮ
    If P_SHORDER_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '資材注文ﾃﾞｰﾀＯＰＥＮ(別ﾎﾟｲﾝﾀｰ)
    If wP_SHORDER_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '在庫ﾃﾞｰﾀＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '月平均出荷数ＯＰＥＮ   '2018.04.19
    If AVE_SYUKA_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    
    Load PI000302
    
    
    
    '管理マスタの読み込み
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)

    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            If P_KANRI_MAKE_Proc() Then
                Unload Me
            End If
        Case Else
            Call File_Error(sts, BtOpGetEqual, "管理マスタ")
            Unload Me
    End Select
        
    
    
    
    '注文先
    If Ukeharai_Set_Proc(pcmbORDER) Then
        Unload Me
    End If
    '納入先
    If Ukeharai_Set_Proc(pcmbDELI) Then
        Unload Me
    End If
    
    
    Input_Mode = 0      '2007.11.12
    
    '画面初期設定
    If Init_Proc() Then
        Unload Me
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
                                            
                                            
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
    
    
                                            '管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "管理マスタ")
        End If
    End If
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
                                            '受払先マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "受払先マスタ")
        End If
    End If
                                            '資材注文ﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材注文ﾃﾞｰﾀ")
        End If
    End If
                                            '資材注文ﾃﾞｰﾀＣＬＯＳＥ（別ﾎﾟｲﾝﾀｰ）
    sts = BTRV(BtOpClose, wP_SHORDER_POS, wP_SHORDER_REC, Len(wP_SHORDER_REC), K2_wP_SHORDER, Len(K2_wP_SHORDER), 2)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材注文ﾃﾞｰﾀ")
        End If
    End If
                                            '在庫ﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫ﾃﾞｰﾀ")
        End If
    End If
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PI000301 = Nothing
    Set PI000302 = Nothing

    End
End Sub





Private Sub TDBGrid1_DblClick()
Dim sts As Integer
    
    
    If IsNull(TDBGrid1.Bookmark) Then       '2016.04.25
        Exit Sub                            '2016.04.25
    End If                                  '2016.04.15
        
    Text1(ptxORDER_NO).Text = SHORDER(TDBGrid1.Bookmark, colORDER_NO)
    '資材注文データのチェック
    sts = P_SHORDER_Read_Proc()
    Select Case sts
        Case False, BtNoErr
                    
            If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = P_KAN_ON Then
                MsgBox "他端末で書き換えられています。"
                TDBGrid1.SetFocus
                Exit Sub
            End If
        
        Case BtErrKeyNotFound
            MsgBox "他端末で書き換えられています。"
            TDBGrid1.SetFocus
            Exit Sub
        Case Else
            Exit Sub
    End Select
    
        
    

End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)


    If IsNull(TDBGrid1.Bookmark) Then       '2016.04.25
        Exit Sub                            '2016.04.25
    End If                                  '2016.04.15


    If Sort_Tbl(ColIndex) = 0 Then
        Sort_Tbl(ColIndex) = 1
    Else
        If Sort_Tbl(ColIndex) = 1 Then
            Sort_Tbl(ColIndex) = 0
        End If
    
    End If
    
    If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then
                    
        SHORDER.QuickSort Min_Row, SHORDER.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
        
        Set TDBGrid1.Array = SHORDER
        
        TDBGrid1.ReBind
        TDBGrid1.Update
        TDBGrid1.MoveFirst


    End If



End Sub



Private Sub TDBGrid2_DblClick()
Dim sts As Integer
    
    If IsNull(TDBGrid2.Bookmark) Then       '2016.04.25
        Exit Sub                            '2016.04.25
    End If                                  '2016.04.15
    
    
    
    Text1(ptxORDER_NO).Text = Z_SHORDER(TDBGrid2.Bookmark, colZ_ORDER_NO)
    '資材注文データのチェック
    sts = P_SHORDER_Read_Proc()
    Select Case sts
        Case False, BtNoErr
                    
            If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = P_KAN_ON Then
                MsgBox "他端末で書き換えられています。"
                TDBGrid1.SetFocus
                Exit Sub
            End If
        
        Case BtErrKeyNotFound
            MsgBox "他端末で書き換えられています。"
            TDBGrid1.SetFocus
            Exit Sub
        Case Else
            Exit Sub
    End Select
    
        
    

End Sub

Private Sub TDBGrid2_HeadClick(ByVal ColIndex As Integer)
    
    
    If IsNull(TDBGrid2.Bookmark) Then       '2016.04.25
        Exit Sub                            '2016.04.25
    End If                                  '2016.04.15
    
    
    
    If Z_Sort_Tbl(ColIndex) = 0 Then
        Z_Sort_Tbl(ColIndex) = 1
    Else
        If Z_Sort_Tbl(ColIndex) = 1 Then
            Z_Sort_Tbl(ColIndex) = 0
        End If
    
    End If
    
    If Z_Sort_Tbl(ColIndex) = 0 Or Z_Sort_Tbl(ColIndex) = 1 Then
                    
        Z_SHORDER.QuickSort Z_Min_Row, Z_SHORDER.UpperBound(1), ColIndex, Z_Sort_Tbl(ColIndex), XTYPE_STRING
        
        Set TDBGrid2.Array = Z_SHORDER
        
        TDBGrid2.ReBind
        TDBGrid2.Update
        TDBGrid2.MoveFirst


    End If

End Sub

Private Sub Text1_DblClick(Index As Integer)

    Select Case Index
        Case ptxSHIIRE_CODE01 To ptxTANKA_DT01

        
            If Trim(Text1(ptxSHIIRE_CODE01).Text) = "" Then
            Else
            
                Call SHIIRE_Disp_Proc(ptxSHIIRE_CODE01)
            
            End If
        
        Case ptxSHIIRE_CODE02 To ptxTANKA_DT02
        
            If Trim(Text1(ptxSHIIRE_CODE01).Text) = "" Then
            Else
                Call SHIIRE_Disp_Proc(ptxSHIIRE_CODE02)
            End If
        
        Case ptxSHIIRE_CODE03 To ptxTANKA_DT03
            
            If Trim(Text1(ptxSHIIRE_CODE01).Text) = "" Then
            Else
                Call SHIIRE_Disp_Proc(ptxSHIIRE_CODE03)
            End If
    
    End Select
    
    '========================================================= 2007/03/19 =====
''    Text1(ptxORDER_CODE).SetFocus
    Text1(ptxORDER_QTY).SetFocus
    '==========================================================================


End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

    '2007.07.27
    If Index = ptxHIN_GAI Then
        svHinban = Text1(Index).Text
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    If Index = ptxHIN_GAI Then
        Text1(Index).Text = StrConv(RTrim(Text1(Index).Text), vbUpperCase)
    End If
        
        
    If Error_Check_Proc(Index) Then     'エラーチェック
        Exit Sub
    End If
        
        
        
    Call Tab_Ctrl(Shift)        '移動
End Sub
Private Function Init_Proc() As Integer
'----------------------------------------------------------------------------
'                   入力画面の初期設定
'----------------------------------------------------------------------------
Dim i           As Integer
Dim sts         As Integer


Dim TANTO_CODE  As String
Dim TANTO_NAME  As String


    Init_Proc = True
    
    TANTO_CODE = Text1(ptxTANTO_CODE).Text
    TANTO_NAME = Text1(ptxTANTO_NAME).Text
    
    
    '2012.09.21▼
'    For i = ptxTANTO_CODE To ptxORDER_ZAN03
'        Text1(i).Text = ""
'    Next i
    
    For i = ptxTANTO_CODE To ptxTANKA_DT03
        Text1(i).Text = ""
    Next i
    
    '2012.09.21▲
    
    
    Text1(ptxUKEIRE_QTY).Text = ""  '2016.06.22
    
    
    Text1(ptxTANTO_CODE).Text = TANTO_CODE
    Text1(ptxTANTO_NAME).Text = TANTO_NAME
    
    Text1(ptxDELI_CODE).Text = NOUNYU
    
    
    '注文日＝当日
    Text1(ptxORDER_DT).Text = Format(Now, "YYYY/MM/DD")



    lblSHIIRE_BIKOU.Caption = ""        '2018.04.19
    txtAVE_SYUKA.Text = ""              '2018.04.19
    txtAVE_SYUKA_cnt.Text = ""         '2018.04.19



    For i = pcmbORDER To pcmbDELI
        
        Combo1(i).ListIndex = -1
    Next i


    If List_Disp_Proc() Then
        Exit Function
    End If

    '入力可否を切り替える   2007.11.12
    Select Case Input_Mode
        Case 0  '通常入力
        
        
            For i = ptxTANTO_CODE To ptxTANKA
            
                If i = ptxTANTO_NAME Or i = ptxHIN_NAME Then
                Else
                                
                    If i = ptxORDER_NO Then
                        Text1(i).Locked = True
                        Text1(i).BackColor = &H8000000F
                        Text1(i).TabStop = False
                    Else
                        Text1(i).Locked = False
                        Text1(i).BackColor = &H80000005
                        Text1(i).TabStop = True
                    End If
                End If
            Next i
        
            For i = pcmbORDER To pcmbDELI
                Combo1(i).Locked = False
                Combo1(i).BackColor = &H80000005
            Next i
                    
        
            For i = ptxSHIIRE_CODE01 To ptxORDER_ZAN03
                Text1(i).Enabled = True
            Next i
            
            
            
        
        

        Case 1  '注文数変更
                
            For i = ptxTANTO_CODE To ptxTANKA
                
                If i = ptxTANTO_NAME Or i = ptxHIN_NAME Then
                Else
                
                    If i = ptxORDER_NO Or i = ptxORDER_QTY Then
                        Text1(i).Locked = False
                        Text1(i).BackColor = &H80000005
                        Text1(i).TabStop = True
                    Else
                        Text1(i).Locked = True
                        Text1(i).BackColor = &H8000000F
                        Text1(i).TabStop = False
                    End If
                End If
            Next i
        
            For i = pcmbORDER To pcmbDELI
                Combo1(i).Locked = True
                Combo1(i).BackColor = &H8000000F
            Next i
        
            For i = ptxSHIIRE_CODE01 To ptxORDER_ZAN03
                Text1(i).Enabled = False
            Next i
        
            
        
    End Select


    'ｿｰﾄ情報の初期化
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0             'ﾃﾞﾌｫﾙﾄ昇順
    Next i

    Sort_Tbl(colHIN_NAME) = 9       'ｿｰﾄ除外

    Init_Proc = False

End Function
Private Function Ukeharai_Set_Proc(Index As Integer) As Integer
'----------------------------------------------------------------------------
'                   受払先マスタをコンボにセットする。
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer




Dim i           As Integer
    
    Ukeharai_Set_Proc = True
    
    Combo1(Index).Clear
    
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, "")

    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "受払先マスタ")
                Exit Function
        
        End Select

        
        
        Combo1(Index).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode) & "            " & _
                                StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        
        com = BtOpGetNext
    
    Loop

    Ukeharai_Set_Proc = False
    



End Function
Private Function P_SHORDER_Read_Proc(Optional Mode As Integer = 0) As Integer
'----------------------------------------------------------------------------
'                   資材注文データの読み込み
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
    
    
    P_SHORDER_Read_Proc = True
    
    
    '資材注文ﾃﾞｰﾀ
    Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, Text1(ptxORDER_NO))
    sts = BTRV(BtOpGetEqual, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
        
    Select Case sts
        Case BtNoErr
        
Debug.Print StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode)
        
        
        Case Else
            P_SHORDER_Read_Proc = sts
            Exit Function
    
    End Select
    
    
    
    
    
    If Item_Disp_Proc(Mode) Then
        Exit Function
    End If
    
    
'    If Mode = 1 Then    '注文数変更の為の再読み込み
    
    
        Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, Text1(ptxORDER_NO))
        sts = BTRV(BtOpGetEqual, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case Else
                P_SHORDER_Read_Proc = sts
                Exit Function
        
        End Select
    
    
    
    
'    End If
    
    P_SHORDER_Read_Proc = False
        
    

End Function
Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'           資材注文ﾃﾞｰﾀの表示
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim com     As Integer

Dim Row     As Long

    List_Disp_Proc = True
    PI000301.MousePointer = vbHourglass
    
    PI000301.Enabled = False                '2016.04.25
    
    Set SHORDER = Nothing
    Tbl_Set_F = False
    
    
    Call UniCode_Conv(K2_P_SHORDER.WS_NO, WS_NO)
    Call UniCode_Conv(K2_P_SHORDER.PRINT_F, P_PRINT_OFF)
    Call UniCode_Conv(K2_P_SHORDER.ORDER_CODE, "")
    Call UniCode_Conv(K2_P_SHORDER.ORDER_NO, "")
    
    com = BtOpGetGreaterEqual
    
    Row = Min_Row - 1
       
    Do
    
        DoEvents
    
        sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K2_P_SHORDER, Len(K2_P_SHORDER), 2)
            
        Select Case sts
            Case BtNoErr
                If Trim(StrConv(P_SHORDER_REC.WS_NO, vbUnicode)) <> Trim(WS_NO) Or _
                    StrConv(P_SHORDER_REC.PRINT_F, vbUnicode) <> P_PRINT_OFF Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "資材注文ﾃﾞｰﾀ")
                Exit Function
        End Select
    
        
        
        
        Row = Row + 1
        
        
        If Grid_Set_Proc(Row) Then
            Exit Function
        End If
        Tbl_Set_F = True
        
        com = BtOpGetNext
    
    Loop
    
    Set TDBGrid1.Array = SHORDER
            
    If Row <> (Min_Row - 1) Then
        SHORDER.QuickSort Min_Row, SHORDER.UpperBound(1), colORDER_NO, XORDER_ASCEND, XTYPE_STRING
    End If
            
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    
    
    PI000301.Enabled = True                 '2016.04.25
    
    PI000301.MousePointer = vbDefault
    List_Disp_Proc = False
    


End Function

Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           資材注文ﾃﾞｰﾀの内容をｸﾞﾘｯﾄﾞにｾｯﾄする
'----------------------------------------------------------------------------
Dim sts As Integer

    Grid_Set_Proc = True
    
    SHORDER.ReDim Min_Row, Row, Min_Col, Max_Col


    '注文№
    SHORDER(Row, colORDER_NO) = StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode)
    '注文日
    SHORDER(Row, colORDER_DT) = Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 7, 2)
    '注文先
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
            Exit Function
    End Select
    '注文先
    SHORDER(Row, colORDER_NAME) = StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode) & " " & _
                                StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode)
    '品番
    SHORDER(Row, colHIN_GAI) = StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)
    '品名
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    End Select
    '品名
    SHORDER(Row, colHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    '注文数
    SHORDER(Row, colORDER_QTY) = Format(CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)), "#0")
    '納期予定日
    If Trim(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode)) <> "" Then '2007.09.06
        SHORDER(Row, colY_NOUKI_DT) = Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 7, 2)
    Else
        SHORDER(Row, colY_NOUKI_DT) = ""
    End If
        
    '納入先
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHORDER_REC.DELI_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
            Exit Function
    End Select
    '注文先
    SHORDER(Row, colDELI_NAME) = StrConv(P_SHORDER_REC.DELI_CODE, vbUnicode) & " " & _
                                StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode)
        

    Grid_Set_Proc = False

End Function
Private Function Print_Proc() As Integer
'----------------------------------------------------------------------------
'           印刷処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
Dim Save_Order_Code As String * 5
                
Dim rpt         As New PI00030F1
Dim f           As New PI000302

                
    Call UniCode_Conv(K2_wP_SHORDER.WS_NO, WS_NO)
    Call UniCode_Conv(K2_wP_SHORDER.PRINT_F, P_PRINT_OFF)
    Call UniCode_Conv(K2_wP_SHORDER.ORDER_CODE, "")
    Call UniCode_Conv(K2_wP_SHORDER.ORDER_NO, "")
                
    com = BtOpGetGreaterEqual
                
    Save_Order_Code = ""

                
    Do
        DoEvents
        
        sts = BTRV(com, wP_SHORDER_POS, wP_SHORDER_REC, Len(wP_SHORDER_REC), K2_wP_SHORDER, Len(K2_wP_SHORDER), 2)
            
        Select Case sts
            Case BtNoErr
            
                If StrConv(wP_SHORDER_REC.WS_NO, vbUnicode) <> WS_NO Or _
                    StrConv(wP_SHORDER_REC.PRINT_F, vbUnicode) <> P_PRINT_OFF Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
                Exit Function
        End Select
    
        If Trim(Save_Order_Code) = "" Then
    
            Set rpt = New PI00030F1
        
            'レポートを印刷します。（true：印刷ダイアログあり false：なし）
            rpt.PrintReport False
        
            Set rpt = Nothing
    
    
    
'            f.RunReport rpt
'            f.Show
    
            Save_Order_Code = StrConv(wP_SHORDER_REC.ORDER_CODE, vbUnicode)
    
    
        End If
    
        If Save_Order_Code <> StrConv(wP_SHORDER_REC.ORDER_CODE, vbUnicode) Then
    
            Set rpt = New PI00030F1
        
            'レポートを印刷します。（true：印刷ダイアログあり false：なし）
            rpt.PrintReport False
        
            Set rpt = Nothing


'            f.RunReport rpt
'            f.Show
    
            Save_Order_Code = StrConv(wP_SHORDER_REC.ORDER_CODE, vbUnicode)
    
    
        End If
    
        com = BtOpGetNext
    
    Loop
                



End Function

Private Sub SHIIRE_Disp_Proc(Index As Integer)
'----------------------------------------------------------------------------
'           最適仕入先からの表示
'----------------------------------------------------------------------------
Dim i   As Integer
    
    
    '注文先ｺｰﾄﾞ
    Text1(ptxORDER_CODE).Text = Trim(Text1(Index).Text)
    '注文先名
    For i = 0 To Combo1(pcmbORDER).ListCount - 1
    
        If Text1(ptxORDER_CODE).Text = Trim(Right(Combo1(pcmbORDER).List(i), 5)) Then
            Combo1(pcmbORDER).ListIndex = i
            Exit For
        End If
    
    Next i
    '単価
    Text1(ptxTANKA).Text = Text1(Index + 2).Text
    '金額
    If IsNumeric(Text1(ptxORDER_QTY).Text) And IsNumeric(Text1(ptxTANKA).Text) Then
        Text1(ptxKINGAKU).Text = Format(CLng(CDbl(Text1(ptxORDER_QTY).Text) * _
                                            CDbl(Text1(ptxTANKA).Text)), "#,##0")
    End If
    'ﾛｯﾄ数
    Text1(ptxLOT).Text = Text1(Index + 3).Text

    '予定納期
    If IsDate(Text1(ptxORDER_DT).Text) And IsNumeric(Text1(Index + 4).Text) Then
    
        Text1(ptxY_NOUKI_DT).Text = Format(DateAdd("d", CDbl(Text1(Index + 4).Text), Text1(ptxORDER_DT).Text), "YYYY/MM/DD")
    Else
        Text1(ptxY_NOUKI_DT).Text = ""          '2007.09.06
    End If


    '使用月 2008.01.10

    If OSAKA_MODE Then
        Text1(ptxUSE_YM).Text = Left(Text1(ptxY_NOUKI_DT).Text, 7)
    Else
        Text1(ptxUSE_YM).Text = ""
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)

Dim sts As Integer

        '2007.07.27
    Select Case Index
        Case ptxHIN_GAI
            
            
            Text1(Index).Text = StrConv(RTrim(Text1(Index).Text), vbUpperCase)
            
'            If svHinban <> Text1(Index).Text Then      2016.01.18
                If Z_List_Disp_Proc() Then
                    Unload Me
                End If
'            End If                                     2016.01.18
        Case ptxORDER_CODE  '2017.11.21
    
            Text1(Index).Text = StrConv(RTrim(Text1(Index).Text), vbUpperCase)
    
        Case ptxDELI_CODE  '2017.11.21
    
            Text1(Index).Text = StrConv(RTrim(Text1(Index).Text), vbUpperCase)
    
    
        Case ptxORDER_NO    '2007.11.12
''            If Input_Mode = 1 Then
''
''
''                '資材注文データのチェック
''                sts = P_SHORDER_Read_Proc(1)
''                Select Case sts
''                    Case False, BtNoErr
''
''                        If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = P_PRINT_ON Then
''                            MsgBox "注文書未発行です。"
''                            Text1(Index).SetFocus
''                            Exit Sub
''                        End If
''
''                        If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
''                            MsgBox "キャンセルされています。"
''                            Text1(Index).SetFocus
''                            Exit Sub
''                        End If
''
''                        If CLng(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)) <> 0 Then
''                            MsgBox "受入実績があります。"
''                            Text1(Index).SetFocus
''                            Exit Sub
''                        End If
''                    Case BtErrKeyNotFound
''                        MsgBox "注文№未登録です。"
''                        Text1(Index).SetFocus
''                        Exit Sub
''                    Case Else
''                        Exit Sub
''                End Select
''
''            End If
    End Select



End Sub
Private Function Z_List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'           資材注文残の表示    2007.07.27
'----------------------------------------------------------------------------
Dim sts                 As Integer
Dim com                 As Integer

Dim Row               As Long

Dim Skip_Flg            As Boolean

Dim i                   As Integer


    Z_List_Disp_Proc = True
    
    PI000301.MousePointer = vbHourglass
'    PI000301.Enabled = False                '2017.10.13
    
    'ｿｰﾄ情報の初期化
    For i = 0 To UBound(Z_Sort_Tbl)
        Z_Sort_Tbl(i) = 0           'ﾃﾞﾌｫﾙﾄ昇順
    Next i

    Z_Sort_Tbl(colZ_HIN_NAME) = 9   'ｿｰﾄ除外
    
    
    
    
    Set Z_SHORDER = Nothing
    
    Row = Z_Min_Row - 1
       
    
    
    Call UniCode_Conv(K1_P_SHORDER.JGYOBU, SHIZAI)                  '事業部
    Call UniCode_Conv(K1_P_SHORDER.NAIGAI, NAIGAI_NAI)              '国内外
    Call UniCode_Conv(K1_P_SHORDER.HIN_GAI, Text1(ptxHIN_GAI).Text) '品番(外部)
    Call UniCode_Conv(K1_P_SHORDER.ORDER_DT, "")                    '注文日
    Call UniCode_Conv(K1_P_SHORDER.ORDER_NO, "")                    '注文№
    
    
    com = BtOpGetGreaterEqual
    
       
       
    Do
    
        DoEvents
    
        sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K1_P_SHORDER, Len(K1_P_SHORDER), 1)
            
        Select Case sts
            Case BtNoErr
                
                If StrConv(P_SHORDER_REC.JGYOBU, vbUnicode) <> SHIZAI Or _
                    StrConv(P_SHORDER_REC.NAIGAI, vbUnicode) <> NAIGAI_NAI Or _
                    Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)) <> Trim(Text1(ptxHIN_GAI).Text) Then
                    Exit Do
                End If
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "資材注文ﾃﾞｰﾀ")
                Exit Function
        End Select
    
    
        Skip_Flg = False
    
        If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = P_KAN_ON Then
            Skip_Flg = True
        End If
        
        If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
            Skip_Flg = True
        End If
        
        
        
        If Not Skip_Flg Then
    
            Row = Row + 1
            
            If Row > LIST_MAX Then              '2017.11.21
                Exit Do                         '2017.11.21
            End If                              '2017.11.21
            
            
            
            If Z_Grid_Set_Proc(Row) Then
                Exit Function
            End If
        
        
        
        End If
        
        com = BtOpGetNext
    
    Loop
    
    
    
    Set TDBGrid2.Array = Z_SHORDER
    TDBGrid2.ReBind
    TDBGrid2.Update
    TDBGrid2.MoveFirst
    
'    PI000301.Enabled = True                '2017.10.13
    PI000301.MousePointer = vbDefault
    Z_List_Disp_Proc = False
    


End Function


Private Function Z_Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           資材注文残の内容をｸﾞﾘｯﾄﾞにｾｯﾄする   2007.07.27
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim i           As Integer

Dim Mi_QTY      As Long
Dim Sumi_QTY    As Long

    Z_Grid_Set_Proc = True
    
    Z_SHORDER.ReDim Z_Min_Row, Row, Z_Min_Col, Z_Max_Col
    
    Z_SHORDER(Row, colZ_ORDER_DT) = Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 7, 2)
    
    
    '注文№
    Z_SHORDER(Row, colZ_ORDER_NO) = Trim(StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode))
    '注文名
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
            Exit Function
    End Select
    Z_SHORDER(Row, colZ_ORDER_NAME) = Trim(StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode)) & " " & StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode)
    '資材品番
    Z_SHORDER(Row, colZ_HIN_GAI) = StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)
    '品名
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    End Select
    Z_SHORDER(Row, colZ_HIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    '手配数
    Z_SHORDER(Row, colZ_ORDER_QTY) = Format(CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)), "#,##0")
    '注文残
    Z_SHORDER(Row, colZ_ZAN_QTY) = Format(CDbl(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) - CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)), "#,##0")
    '現在庫
    If Zaiko_Syukei_Proc(Sumi_QTY, Mi_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                            StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
        Exit Function
    End If
    Z_SHORDER(Row, colZ_ZAIKO_QTY) = Format(Mi_QTY + Sumi_QTY, "#,##0")
    
    '納期予定日
    If Trim(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode)) <> "" Then    '2007.09.06
        Z_SHORDER(Row, colZ_Y_NOUKI_DT) = Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 7, 2)
    Else
        Z_SHORDER(Row, colZ_Y_NOUKI_DT) = ""
    End If
    
    
    
    '使用月 2007.12.05
    If Trim(StrConv(P_SHORDER_REC.USE_YM, vbUnicode)) <> "" Then
        Z_SHORDER(Row, colZ_USE_YM) = Mid(StrConv(P_SHORDER_REC.USE_YM, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(P_SHORDER_REC.USE_YM, vbUnicode), 5, 2)
    Else
        Z_SHORDER(Row, colZ_USE_YM) = ""
    End If
    
    '回答納期日 2007.12.05
    If Trim(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode)) <> "" Then
        Z_SHORDER(Row, colZ_ANS_NOUKI_DT) = Mid(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 7, 2)
    Else
        Z_SHORDER(Row, colZ_ANS_NOUKI_DT) = ""
    End If
    
    
    
    Z_Grid_Set_Proc = False

End Function

' ------------------------------------------------------------------------
'       指定した精度の数値に切り上げします。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に切り上げられた数値。
' ------------------------------------------------------------------------
Private Function ToRoundUp(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    
        


    dCoef = (10 ^ iDigits)



    
    
    
    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
        Case Is > 0
            ToRoundUp = (Int(dValue * dCoef) + 1) / dCoef
        Case Is < 0
            ToRoundUp = (Fix(dValue * dCoef) - 1) / dCoef
        Case Else
            ToRoundUp = dValue
    End Select


'    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
'        Case Is > 0
'            ToRoundUp = (Int(dValue * dCoef + 0.9)) / dCoef
'        Case Is < 0
'            ToRoundUp = (Fix(dValue * dCoef - 0.9)) / dCoef
'        Case Else
'            ToRoundUp = dValue
'    End Select



End Function

' ------------------------------------------------------------------------
'       指定した精度の数値に切り捨てします。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に切り捨てられた数値。
' ------------------------------------------------------------------------
Public Function ToRoundDown(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
        Case Is > 0
            ToRoundDown = Int(dValue * dCoef) / dCoef
        Case Is < 0
            ToRoundDown = Fix(dValue * dCoef) / dCoef
        Case Else
            ToRoundDown = dValue
    End Select
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

