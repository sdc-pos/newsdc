VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PI000411 
   Caption         =   "資材仕入処理"
   ClientHeight    =   11085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15510
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
   ScaleHeight     =   11085
   ScaleWidth      =   15510
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox txtP_NYUKA_QTY 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.TextBox txtP_NYUKA_DT 
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton Command2 
      Caption         =   "対象収支"
      Height          =   375
      Left            =   11400
      TabIndex        =   65
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   24
      Left            =   8600
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   23
      Left            =   8085
      MaxLength       =   3
      TabIndex        =   9
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   11760
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1080
      Width           =   960
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   11
      Left            =   4410
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   22
      Left            =   10230
      MaxLength       =   10
      TabIndex        =   27
      Top             =   4200
      Width           =   1590
   End
   Begin VB.CheckBox Check1 
      Caption         =   "POS在庫計上"
      Height          =   375
      Index           =   0
      Left            =   12960
      TabIndex        =   25
      Top             =   3840
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   8600
      Sorted          =   -1  'True
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   5
      Top             =   1080
      Width           =   2115
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   8085
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1080
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   13
      Left            =   4800
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   21
      Left            =   10230
      MaxLength       =   10
      TabIndex        =   26
      Top             =   3720
      Width           =   1590
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   20
      Left            =   13590
      MaxLength       =   8
      TabIndex        =   24
      Top             =   3360
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      Index           =   19
      Left            =   10230
      MaxLength       =   8
      TabIndex        =   23
      Top             =   3360
      Width           =   1590
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   18
      Left            =   13590
      MaxLength       =   7
      TabIndex        =   22
      Top             =   2760
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   17
      Left            =   10230
      MaxLength       =   10
      TabIndex        =   21
      Top             =   2760
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   15
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   16
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   14
      Left            =   1575
      MaxLength       =   11
      TabIndex        =   18
      Top             =   3720
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   12
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   10
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   9
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2400
      Width           =   750
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   1
      Left            =   2310
      Locked          =   -1  'True
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2850
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2040
      Width           =   750
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   2
      Left            =   2310
      Locked          =   -1  'True
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2850
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2115
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1560
      Width           =   750
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   3
      Left            =   4095
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2745
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   1575
      MaxLength       =   5
      TabIndex        =   0
      Top             =   240
      Width           =   750
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   4935
      Left            =   105
      TabIndex        =   28
      Top             =   5160
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   8705
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "注文日時"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "注文№"
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
      Columns(9).Caption=   "回答納期"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "収支"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2408"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2302"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=512"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1614"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1508"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=512"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3016"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2910"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2699"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2593"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=512"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=3334"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=3228"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=512"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2064"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1958"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=2064"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1958"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=514"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=2064"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=1958"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=514"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=2328"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2223"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=512"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(9).Width=2328"
      Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2223"
      Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=516"
      Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(51)=   "Column(10).Width=1217"
      Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=1111"
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
      _StyleDefs(72)  =   "Splits(0).Columns(6).Style:id=24,.parent=43,.alignment=1"
      _StyleDefs(73)  =   "Splits(0).Columns(6).HeadingStyle:id=21,.parent=44"
      _StyleDefs(74)  =   "Splits(0).Columns(6).FooterStyle:id=22,.parent=45"
      _StyleDefs(75)  =   "Splits(0).Columns(6).EditorStyle:id=23,.parent=47"
      _StyleDefs(76)  =   "Splits(0).Columns(7).Style:id=20,.parent=43,.alignment=1"
      _StyleDefs(77)  =   "Splits(0).Columns(7).HeadingStyle:id=17,.parent=44"
      _StyleDefs(78)  =   "Splits(0).Columns(7).FooterStyle:id=18,.parent=45"
      _StyleDefs(79)  =   "Splits(0).Columns(7).EditorStyle:id=19,.parent=47"
      _StyleDefs(80)  =   "Splits(0).Columns(8).Style:id=70,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(81)  =   ":id=70,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(82)  =   ":id=70,.fontname=ＭＳ ゴシック"
      _StyleDefs(83)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=44"
      _StyleDefs(84)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=45"
      _StyleDefs(85)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=47"
      _StyleDefs(86)  =   "Splits(0).Columns(9).Style:id=74,.parent=43"
      _StyleDefs(87)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=44"
      _StyleDefs(88)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=45"
      _StyleDefs(89)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=47"
      _StyleDefs(90)  =   "Splits(0).Columns(10).Style:id=78,.parent=43"
      _StyleDefs(91)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=44"
      _StyleDefs(92)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=45"
      _StyleDefs(93)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=47"
      _StyleDefs(94)  =   "Named:id=33:Normal"
      _StyleDefs(95)  =   ":id=33,.parent=0"
      _StyleDefs(96)  =   "Named:id=34:Heading"
      _StyleDefs(97)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(98)  =   ":id=34,.wraptext=-1"
      _StyleDefs(99)  =   "Named:id=35:Footing"
      _StyleDefs(100) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(101) =   "Named:id=36:Selected"
      _StyleDefs(102) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(103) =   "Named:id=37:Caption"
      _StyleDefs(104) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(105) =   "Named:id=38:HighlightRow"
      _StyleDefs(106) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(107) =   "Named:id=39:EvenRow"
      _StyleDefs(108) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(109) =   "Named:id=40:OddRow"
      _StyleDefs(110) =   ":id=40,.parent=33"
      _StyleDefs(111) =   "Named:id=41:RecordSelector"
      _StyleDefs(112) =   ":id=41,.parent=34"
      _StyleDefs(113) =   "Named:id=42:FilterBar"
      _StyleDefs(114) =   ":id=42,.parent=33"
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
      TabIndex        =   40
      Top             =   10320
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
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   10320
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
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   10320
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
      Index           =   8
      Left            =   7920
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   10320
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
      Index           =   7
      Left            =   6600
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   10320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "納期変更"
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
      Index           =   6
      Left            =   5760
      TabIndex        =   35
      Top             =   10320
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
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   10320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "最 新"
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
      TabIndex        =   33
      Top             =   10320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ｷｬﾝｾﾙ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   32
      Top             =   10320
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
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   10320
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
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   10320
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
      TabIndex        =   29
      Top             =   10320
      Width           =   855
   End
   Begin VB.Label lblP_NYUKA_QTY 
      Alignment       =   1  '右揃え
      Caption         =   "前借数"
      Height          =   255
      Left            =   6120
      TabIndex        =   68
      Top             =   3960
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label LBLP_NYUKA_DT 
      Alignment       =   1  '右揃え
      Caption         =   "前借日付"
      Height          =   255
      Left            =   3600
      TabIndex        =   67
      Top             =   3960
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   " 収支単位"
      Height          =   240
      Index           =   1
      Left            =   6930
      TabIndex        =   64
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "使用月"
      Height          =   255
      Index           =   21
      Left            =   10815
      TabIndex        =   63
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "回答納期日"
      Height          =   255
      Index           =   20
      Left            =   3045
      TabIndex        =   62
      Top             =   3000
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "消費税額"
      Height          =   255
      Index           =   19
      Left            =   9075
      TabIndex        =   61
      Top             =   4320
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "仕入区分"
      Height          =   255
      Index           =   18
      Left            =   6930
      TabIndex        =   60
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "）"
      Height          =   255
      Index           =   17
      Left            =   5880
      TabIndex        =   59
      Top             =   3480
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "（うち納品済数"
      Height          =   255
      Index           =   16
      Left            =   2730
      TabIndex        =   57
      Top             =   3480
      Width           =   1905
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "金額"
      Height          =   255
      Index           =   15
      Left            =   9495
      TabIndex        =   56
      Top             =   3840
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "注文残"
      Height          =   255
      Index           =   14
      Left            =   12750
      TabIndex        =   55
      Top             =   3480
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "今回受入数量"
      Height          =   255
      Index           =   13
      Left            =   8655
      TabIndex        =   54
      Top             =   3480
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "処理年月"
      Height          =   255
      Index           =   12
      Left            =   12330
      TabIndex        =   53
      Top             =   2880
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "受入日"
      Height          =   255
      Index           =   9
      Left            =   9285
      TabIndex        =   52
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "在庫残"
      Height          =   255
      Index           =   10
      Left            =   735
      TabIndex        =   51
      Top             =   4320
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "発注ﾛｯﾄ"
      Height          =   255
      Index           =   11
      Left            =   525
      TabIndex        =   50
      Top             =   4680
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "単価"
      Height          =   255
      Index           =   8
      Left            =   945
      TabIndex        =   49
      Top             =   3840
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "注文数"
      Height          =   255
      Index           =   7
      Left            =   630
      TabIndex        =   48
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "希望納期日"
      Height          =   255
      Index           =   6
      Left            =   210
      TabIndex        =   47
      Top             =   3000
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "納入先"
      Height          =   255
      Index           =   5
      Left            =   525
      TabIndex        =   46
      Top             =   2520
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "注文先"
      Height          =   255
      Index           =   4
      Left            =   525
      TabIndex        =   45
      Top             =   2160
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "担当者"
      Height          =   255
      Index           =   2
      Left            =   525
      TabIndex        =   44
      Top             =   1680
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "資材品番"
      Height          =   255
      Index           =   1
      Left            =   525
      TabIndex        =   43
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "注文日"
      Height          =   255
      Index           =   0
      Left            =   630
      TabIndex        =   42
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      Caption         =   "注文№"
      Height          =   255
      Index           =   3
      Left            =   630
      TabIndex        =   41
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "PI000411"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private NOUKI_MODE  As Boolean
Private Input_Mode  As Boolean

Private WS_NO       As String * 10
    
Private KASO_NYUKA  As String * 2           '入荷倉庫
Private POS_UMU     As Boolean              'POSｼｽﾃﾑの有無
    
Private MEMO_TEXT   As String               '履歴メモ
   
   
Private LIST_MAX    As Long                 '2016.01.14
   
'ラベル用添字
Private Const plblANS_NOUKI_DT% = 20        '回答納期日 2008.01.10
Private Const plblUSE_YM% = 21              '使用月     2008.01.10
   
   
'テキスト用添字

Private Const ptxORDER_NO% = 0              '注文№
Private Const ptxORDER_DT% = 1              '注文日
Private Const ptxHIN_GAI% = 2               '品番
Private Const ptxHIN_NAME% = 3              '品名
Private Const ptxG_SHIIRE_KBN% = 4          '仕入区分

Private Const ptxUSE_YM% = 5                '使用月 2008.01.10

Private Const ptxTANTO_CODE% = 6            '担当者ｺｰﾄﾞ
Private Const ptxTANTO_NAME% = 7            '担当者名称
Private Const ptxORDER_CODE% = 8            '注文先
Private Const ptxDELI_CODE% = 9             '納入先
Private Const ptxY_NOUKI_DT% = 10           '納期予定日


Private Const ptxANS_NOUKI_DT% = 11         '回答納期日 2008.01.10


Private Const ptxORDER_QTY% = 12            '注文数
Private Const ptxUKEIRE_QTY% = 13           '受入済数
Private Const ptxTANKA% = 14                '単価
Private Const ptxZAIKO_QTY% = 15            '在庫残
Private Const ptxLOT% = 16                  '発注ﾛｯﾄ
Private Const ptxUKEIRE_DT% = 17            '受入日


Private Const ptxKEIJYO_YM% = 18            '計上年月
Private Const ptxKONKAI_UKEIRE_QTY% = 19    '今回納品数量
Private Const ptxZAN_QTY% = 20              '注文残
Private Const ptxKINGAKU% = 21              '金額
Private Const ptxZEI_KIN% = 22              '消費税

Private Const ptxG_SYUSHI% = 23             '収支ｺｰﾄﾞ   2012.12.28
Private Const ptxSYUSHI_NM% = 24            '収支名     2012.12.28

'コンボ用添字
Private Const pcmbG_SHIIRE_KBN% = 0         '仕入区分
Private Const pcmbORDER% = 1                '注文先
Private Const pcmbDELI% = 2                 '納入先


'コマンド特殊機能
Private Const cmdNOUKI% = 6                 '取り消し

'ﾁｪｯｸﾎﾞｯｸｽ用添字
Private Const chkZAIKO_F% = 0

'Glid用環境
Private SHORDER  As New XArrayDB

Private Const Min_Row% = 1                  '最小行数
Private Const Min_Col% = 0                  '最小列数
Private Const Max_Col% = 10                 '最大列数   9--> 10 2016.01.19


Private Const colORDER_DT% = 0              '注文日
Private Const colORDER_NO% = 1              '注文№
Private Const colORDER_NAME% = 2            '発注先名
Private Const colHIN_GAI% = 3               '品番
Private Const colHIN_NAME% = 4              '品名
Private Const colORDER_QTY% = 5             '注文数
Private Const colZAN_QTY% = 6               '注文残
Private Const colZAIKO_QTY% = 7             '在庫残
Private Const colY_NOUKI_DT% = 8            '納期予定日
Private Const colANS_NOUKI_DT% = 9          '回答予定日
Private Const colG_SYUSHI% = 10             '収支   2016.01.19



Private Sort_Tbl(colORDER_DT To colG_SYUSHI) _
                As Integer                  'ｿｰﾄの制御 0:昇順 1:降順
Private Tbl_Set_F   As Boolean

Private Save_UKEIRE_QTY     As Long             '受入数のセーブ
                                            
Private wkUKEIRE_QTY        As String
Private wkTANKA             As String

'---------------    大阪ＰＣモード  True:大阪PC　False:以外     2008.01.10
Private OSAKA_MODE  As Boolean


Private G_SYUSHI_TBL    As Variant          ' 対象収支      2008.10.09


Private P_NYUKA_DSP     As Integer          '前借表示   2016.09.08


Private UKEIRE_DT       As Integer          '上下限設定 受入日　2017.04.25
Private KEIJYO_YM       As Integer          '上下限設定 計上月　2017.04.25

'Private Const LAST_UPDATE_DAY$ = " [PI00041] 2018.01.31 10:30"
'Private Const LAST_UPDATE_DAY$ = " [PI00041] 2018.04.09 09:00"
Private Const LAST_UPDATE_DAY$ = " [PI00041] 2019.10.02 12:00" '消費税10%+1円になるバグ修正

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PI000411.MousePointer = vbHourglass


    Call Ctrl_Lock(PI000411)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PI000411)


    PI000411.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer

Dim SUMI_QTY    As Long
Dim MI_QTY      As Long
    
Dim wkDATE      As String
Dim ckDATE      As String               '2018.01.31
    
Dim ZEI         As Long
Dim wkKINGAKU   As Long
       
Dim SYUSHI_ON   As Boolean              '2008.10.09
    
    
    Error_Check_Proc = True
    
    Select Case Mode
        
        Case ptxORDER_NO    '注文№
        
            If Not NOUKI_MODE Then
            
                If Trim(Text1(ptxORDER_NO).Text) = "" Then
                    '注文なし入力
                
                    Call Input_Area_Proc(1)
                
                Else
            
                    Call Input_Area_Proc(0)
            
            
                    If IsNumeric(Text1(ptxORDER_NO).Text) Then
                        Text1(ptxORDER_NO).Text = Format(CLng(Text1(ptxORDER_NO).Text), "00000")
                    End If
                    '資材注文ﾃﾞｰﾀのﾁｪｯｸ
                    
                    If Text1(ptxORDER_NO).Text = StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode) Then
                        sts = BtNoErr
                    Else
                        sts = P_SHORDER_Read_Proc()
                    End If
                    Select Case sts
                        Case False, BtNoErr
                                    
                            If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = P_KAN_ON Then
                                MsgBox "仕入処理済みです。"
                                Text1(Mode).SetFocus
                                Exit Function
                            End If
                        
                            If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
                                MsgBox "キャンセル処理済みです。"
                                Text1(Mode).SetFocus
                                Exit Function
                            End If
                            '2007.09.06 予定納期未設定は受入不可
'                            If Trim(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode)) = "" Then
'                                MsgBox "予定納期が設定されていません。"
'                                Text1(Mode).SetFocus
'                                Exit Function
'                            End If
                        
                        
                        
                            SYUSHI_ON = False               '2008.10.09
                            If GLB_SYUSHI_F = "" Then       '2008.10.09
                                SYUSHI_ON = True
                            Else
                                SYUSHI_ON = False
                                
                                For i = 0 To UBound(G_SYUSHI_TBL)

' 2012.12.28 Upd >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'                                    If Trim(StrConv(P_SHORDER_REC.G_SYUSHI, vbUnicode)) = G_SYUSHI_TBL(i) Then
                                    If Trim(Text1(ptxG_SYUSHI).Text) = G_SYUSHI_TBL(i) Then
                                        SYUSHI_ON = True
                                        Exit For
                                    End If
' 2012.12.28 Upd <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


                                Next i
                            End If



                            If Not SYUSHI_ON Then           '2008.10.09
'                                MsgBox "収支違いです。"                        '2016.04.28
'                                MsgBox "収支対象外です。[対象収支]の押下で入力可能収支を確認しく下さい。(PI00041.INI)"
                                MsgBox "PI00041.iniに収支の登録がありません。" & Chr(13) & Chr(10) & "入力可能収支を確認して下さい。"
                                Text1(Mode).SetFocus
                                Exit Function
                            End If


                        Case BtErrKeyNotFound
                            MsgBox "入力した項目はエラーです。"
                            Text1(Mode).SetFocus
                            Exit Function
                        Case Else
                            Exit Function
                    End Select

                End If
            End If
        
        Case ptxHIN_GAI     '品番外
            If Not NOUKI_MODE Then
            
                Text1(i).Text = StrConv(RTrim(Text1(i).Text), vbUpperCase)      '2013.10.08
            
            
                sts = Hin_Item_Disp_Proc()
                Select Case sts
                    Case False
                    Case BtErrKeyNotFound
                        MsgBox "入力した項目はエラーです。"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Exit Function
                End Select
            
                
                
                If Check1(chkZAIKO_F).Value = vbChecked Then                    '在庫計上有無によりエラーチェックの有無判定 2011.04.13

                
                    If Not POS_UMU Then      '2006.04.26 ＰＯＳなしなら
                    
                    
                    
                    
                    
                        If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" And _
                            Trim(StrConv(ITEMREC.ST_RETU, vbUnicode)) = "" And _
                            Trim(StrConv(ITEMREC.ST_REN, vbUnicode)) = "" And _
                            Trim(StrConv(ITEMREC.ST_DAN, vbUnicode)) = "" Then
                    
                            MsgBox "標準棚番が設定されていません。"
                            Text1(Mode).SetFocus
                            Exit Function
                    
                        End If
                    
                    End If
                End If
            
            
            
            End If
        
        Case ptxG_SHIIRE_KBN    '仕入区分
            If Not NOUKI_MODE Then
        
                Combo1(pcmbG_SHIIRE_KBN).ListIndex = -1
                For i = 0 To Combo1(pcmbG_SHIIRE_KBN).ListCount - 1
                
                    If Text1(ptxG_SHIIRE_KBN).Text = Trim(Left(Right(Combo1(pcmbG_SHIIRE_KBN).List(i), 3), 2)) Then
                        Combo1(pcmbG_SHIIRE_KBN).ListIndex = i
                        Exit For
                    End If
                
                Next i
        
                If i = -1 Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            
            End If
        
        Case ptxUSE_YM  '使用月   2007.12.05
        
            If NOUKI_MODE Then
            
                
                If OSAKA_MODE Then
                
                
                    If Trim(Text1(ptxUSE_YM).Text) = "" Then
                    Else
                        If Not IsDate(Text1(ptxUSE_YM).Text & "/01") Then
                            MsgBox "入力した項目はエラーです。"
                            Text1(Mode).SetFocus
                            Exit Function
                        Else
                            Text1(ptxUSE_YM).Text = Format(CDate(Text1(ptxUSE_YM).Text & "/01"), "YYYY/MM/DD")
                        End If
            
                    End If
            
                End If
            
            End If
        
        
        Case ptxTANTO_CODE  '担当者
            If Not NOUKI_MODE Then
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
            
            
            
            End If
        
        
        Case ptxORDER_CODE      '注文先
            If Not NOUKI_MODE Then
        
        
        
                Text1(ptxORDER_CODE).Text = StrConv(Text1(ptxORDER_CODE).Text, vbUpperCase)         '2017.04.25
        
        
        
                Combo1(pcmbORDER).ListIndex = -1
                For i = 0 To Combo1(pcmbORDER).ListCount - 1
                
                    If Text1(ptxORDER_CODE).Text = Trim(Right(Combo1(pcmbORDER).List(i), 5)) Then
                        Combo1(pcmbORDER).ListIndex = i
                        Exit For
                    End If
                
                Next i
        
                If i = -1 Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            
            End If
        
        
        Case ptxY_NOUKI_DT  '納期予定日
        
            If NOUKI_MODE Then
            
                If Not IsDate(Text1(ptxY_NOUKI_DT).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxY_NOUKI_DT).Text = Format(CDate(Text1(ptxY_NOUKI_DT).Text), "YYYY/MM/DD")
                End If
            
            End If
        
        
        Case ptxANS_NOUKI_DT  '回答納期日   2008.01.10
        
            If NOUKI_MODE Then
            
                
                If OSAKA_MODE Then
                
                
                    If Trim(Text1(ptxANS_NOUKI_DT).Text) = "" Then
                    Else
                        If Not IsDate(Text1(ptxANS_NOUKI_DT).Text) Then
                            MsgBox "入力した項目はエラーです。"
                            Text1(Mode).SetFocus
                            Exit Function
                        Else
                            Text1(ptxANS_NOUKI_DT).Text = Format(CDate(Text1(ptxANS_NOUKI_DT).Text), "YYYY/MM/DD")
                        End If
            
                    End If
            
                End If
            
            End If
        
        
        
        Case ptxUKEIRE_DT   '受入日
            
            If Not NOUKI_MODE Then
            
                If Not IsDate(Text1(ptxUKEIRE_DT).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    
                    
                    Text1(ptxUKEIRE_DT).Text = Format(CDate(Text1(ptxUKEIRE_DT).Text), "YYYY/MM/DD")
                    
                    If Input_Mode Then
                    
                        '注文日
                        Text1(ptxORDER_DT).Text = Format(CDate(Text1(ptxUKEIRE_DT).Text), "YYYY/MM/DD")
                        '納期予定日
                        Text1(ptxY_NOUKI_DT).Text = Format(CDate(Text1(ptxUKEIRE_DT).Text), "YYYY/MM/DD")
                    
                    End If
                
                
                End If
        
        
'>>>>>>>>>>>>>>>>>>>>>  上下限範囲ﾁｪｯｸ 2017.04.25
                If DateAdd("m", UKEIRE_DT * -1, Format(Now, "YYYY/MM/DD")) <= Text1(ptxUKEIRE_DT).Text And _
                    DateAdd("m", UKEIRE_DT, Format(Now, "YYYY/MM/DD")) >= Text1(ptxUKEIRE_DT).Text Then
                Else
                    MsgBox "受入日付が日付範囲を超えています。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If



'>>>>>>>>>>>>>>>>>>>>>  上下限範囲ﾁｪｯｸ 2017.04.25
        
        
        
        
            End If
        
        
        
        
        
        
        
        
        
        
        Case ptxKEIJYO_YM       '処理年月
            
            If Not NOUKI_MODE Then
            
                wkDATE = Text1(ptxKEIJYO_YM).Text & "/01"
                
                If Not IsDate(wkDATE) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    
                    
                    
                    
                    
                    
                    
                    wkDATE = Format(CDate(Text1(ptxKEIJYO_YM).Text), "YYYY/MM/DD")
                    
                    Text1(ptxKEIJYO_YM).Text = Mid(wkDATE, 1, 7)
                
                
                
'>>>>>>>>>>>>>>>>>>>>>  上下限範囲ﾁｪｯｸ 2017.04.25
                
                
                
               If Format(DateAdd("m", KEIJYO_YM * -1, Format(Now, "YYYY/MM/DD")), "YYYY/MM/DD") > Text1(ptxKEIJYO_YM).Text & Right(Format(Now, "YYYY/MM/DD"), 3) Then
                    
                    MsgBox "処理年月が日付範囲を超えています。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If

                
'>>>>>>>>>>>>>  2018.01.31
                ckDATE = (Text1(ptxKEIJYO_YM).Text & Right(Format(Now, "YYYY/MM/DD"), 3))
                Do
                
                    If IsDate(ckDATE) Then
                        Exit Do
                    End If
                    ckDATE = Left(ckDATE, 8) & Val(Right(ckDATE, 2)) - 1
                Loop
'>>>>>>>>>>>>>  2018.01.31

'                If Format(DateAdd("m", KEIJYO_YM, Format(Now, "YYYY/MM/DD")), "YYYY/MM/DD") < (Text1(ptxKEIJYO_YM).Text & Right(Format(Now, "YYYY/MM/DD"), 3)) Then    '2018.01.31
                If Format(DateAdd("m", KEIJYO_YM, Format(Now, "YYYY/MM/DD")), "YYYY/MM/DD") < ckDATE Then                                                               '2018.01.31
                    MsgBox "処理年月が日付範囲を超えています。"
                    Text1(Mode).SetFocus
                    Exit Function
                End If



'>>>>>>>>>>>>>>>>>>>>>  上下限範囲ﾁｪｯｸ 2017.04.25
                
                
                
                
                
                
                End If
            End If
        
        Case ptxKONKAI_UKEIRE_QTY   '受入数
    
            If Not NOUKI_MODE Then
            
                If Not IsNumeric(Text1(ptxKONKAI_UKEIRE_QTY).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxKONKAI_UKEIRE_QTY).Text = Format(CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text), "#0")
                    
                                        
                    If Input_Mode Then
                        '注文数
                        Text1(ptxORDER_QTY).Text = Format(CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text), "#0")
                        '受入済数
                        Text1(ptxUKEIRE_QTY).Text = "0"
                        Call UniCode_Conv(P_SHORDER_REC.UKEIRE_QTY, "00000000.00")  '2007.08.01
                    End If
                    
                    
''                    If CLng(Text1(ptxORDER_QTY).Text) - CLng(Text1(ptxUKEIRE_QTY).Text) < CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text) Then
''                        MsgBox "入力した項目はエラーです。"
''                       Text1(Mode).SetFocus
''                        Exit Function
''                    End If
                    
                    
                    
                    If CLng(CLng(Text1(ptxORDER_QTY).Text) - CLng(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)) - CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)) < 0 Then
                        Text1(ptxZAN_QTY).Text = "0"
                    Else
                        
                        If Save_UKEIRE_QTY = CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text) Then
                        Else
                        
                            Text1(ptxZAN_QTY).Text = Format(CLng(Text1(ptxORDER_QTY).Text) - CLng(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)) - CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text), "#0")
                            Save_UKEIRE_QTY = CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)
                        End If
                    End If
                    
                    If Trim(Text1(ptxKINGAKU).Text) = "" Then
                        '2009.11.02
'                        Text1(ptxKINGAKU).Text = Format(CDbl(Text1(ptxTANKA).Text) * CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text), "#,##0")
                    
                        
                        '2015.09.18
                        If Not IsNumeric(Text1(ptxTANKA).Text) Then
                            Text1(ptxTANKA).Text = "0"
                        End If
                        '2015.09.18
                        
                        
                        
                        Select Case StrConv(P_KANRIREC.SHI_MARUME, vbUnicode)
                            Case "0"    '切捨て
                                Text1(ptxKINGAKU).Text = Format(ToRoundDown(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                                CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")
                            
                
                            Case "5"    '四捨五入
                            
                                Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                                CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")
                            
                            
                            
                            
                            Case "9"    '切り上げ
                        
                        
                                Text1(ptxKINGAKU).Text = Format(ToRoundUp(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                                CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")
                    
                        
                        
                            Case Else    '四捨五入
                            
                                Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                                CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")
                        
                        
                        End Select
                    
                    
                    
                    
                    
                        '2007.11.01 他ｾﾝﾀｰ、時給は消費税計算しない
                        Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxORDER_CODE).Text)
                        sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                MsgBox "受払先マスタ未登録です。"
                                Text1(ptxORDER_CODE).SetFocus
                                Exit Function
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
                                Exit Function
                        End Select
                                           
                                           
                                           
                        If StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) <> P_TORI_ANOTHER And StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) <> P_TORI_JIKYU Then
                                           
                                           
                                           
                                           
                    
                            If CLng(Text1(ptxKINGAKU).Text) >= 0 Then
                                If Format(Text1(ptxUKEIRE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                                    ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                                            CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
                                Else
                                    ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                            CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))
                                End If
                            Else
                                
                                wkKINGAKU = CLng(Text1(ptxKINGAKU).Text) * -1
                                
                                If Format(Text1(ptxUKEIRE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                                    ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                                            CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
                                Else
                                    ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                            CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))
                                End If
                                ZEI = ZEI * -1
                            End If
    
                            Text1(ptxZEI_KIN).Text = Format(ZEI, "#,##0")
                        Else
                            Text1(ptxZEI_KIN).Text = "0"
                        End If
                    
                    End If
                
                    If CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text) <= 0 Then
                                        
                        Check1(chkZAIKO_F).Value = vbUnchecked  '2007.08.02
                    End If
                End If
    
    
    
            End If
    
        Case ptxZAN_QTY         '注文残
    
            If Not NOUKI_MODE Then
    
    
                If Not IsNumeric(Text1(ptxZAN_QTY).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxZAN_QTY).Text = Format(CLng(Text1(ptxZAN_QTY).Text), "#0")
                    '注文残
                    If (CLng(Text1(ptxORDER_QTY).Text) - CLng(Text1(ptxUKEIRE_QTY).Text) - CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)) = CLng(Text1(ptxZAN_QTY).Text) Or _
                        CLng(Text1(ptxZAN_QTY).Text) = 0 Then
                    Else
                        MsgBox "入力した項目はエラーです。"
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
                
                        
                End If
            
            End If
    
    
        Case ptxTANKA       '単価
    
            If Not NOUKI_MODE Then
    
    
                If Not IsNumeric(Text1(ptxTANKA).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    '単価
                    Text1(ptxTANKA).Text = Format(CDbl(Text1(ptxTANKA).Text), "#0.00")
                    '金額計算
                    If IsNumeric(Text1(ptxORDER_QTY).Text) Then
                        
                        If Trim(Text1(ptxKINGAKU).Text) = "" Then
                            '2009.11.02
'                            Text1(ptxKINGAKU).Text = Format(CLng(CDbl(Text1(ptxTANKA).Text) * CDbl(Text1(ptxKONKAI_UKEIRE_QTY).Text)), "#,##0")
                        
                            Select Case StrConv(P_KANRIREC.SHI_MARUME, vbUnicode)
                                Case "0"    '切捨て
                                    Text1(ptxKINGAKU).Text = Format(ToRoundDown(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                                    CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")
                                
                    
                                Case "5"    '四捨五入
                                
                                    Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                                    CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")
                                
                                
                                
                                
                                Case "9"    '切り上げ
                            
                            
                                    Text1(ptxKINGAKU).Text = Format(ToRoundUp(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                                    CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")
                        
                            
                            
                                Case Else    '四捨五入
                                
                                    Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                                    CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")
                            
                            
                            End Select
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                            '2007.11.01 他ｾﾝﾀｰ、時給は消費税計算しない
                            Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxORDER_CODE).Text)
                            sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                                    MsgBox "受払先マスタ未登録です。"
                                    Text1(ptxORDER_CODE).SetFocus
                                    Exit Function
                                
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
                                    Exit Function
                            End Select
                                               
                                               
                                               
                            If StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) <> P_TORI_ANOTHER And StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) <> P_TORI_JIKYU Then
                        
                        
                                If CLng(Text1(ptxKINGAKU).Text) >= 0 Then
                                    If Format(Text1(ptxUKEIRE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                                        ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                                                CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
                                    Else
                                        ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                                CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))
                                    End If
                                Else
                                    
                                    wkKINGAKU = CLng(Text1(ptxKINGAKU).Text) * -1
                                    
                                    If Format(Text1(ptxUKEIRE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                                        ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                                                CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
                                    Else
                                        ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                                CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))
                                    End If
                                    ZEI = ZEI * -1
                                End If
        
                                Text1(ptxZEI_KIN).Text = Format(ZEI, "#,##0")
                            Else
                                Text1(ptxZEI_KIN).Text = "0"
                            End If
                        
                        
                        
                        End If
                    End If
                        
                        
                End If
            
            End If
    
    
    
        Case ptxKINGAKU       '金額
    
            If Not NOUKI_MODE Then
    
    
                If Not IsNumeric(Text1(ptxKINGAKU).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxKINGAKU).Text = Format(CLng(Text1(ptxKINGAKU).Text), "#,##0")
                    '金額計算
                    If Trim(Text1(ptxZEI_KIN).Text) = "" Then
                        
                        
                        '2007.11.01 他ｾﾝﾀｰ、時給は消費税計算しない
                        Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxORDER_CODE).Text)
                        sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                MsgBox "受払先マスタ未登録です。"
                                Text1(ptxORDER_CODE).SetFocus
                                Exit Function
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
                                Exit Function
                        End Select
                                           
                                           
                                           
                        If StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) <> P_TORI_ANOTHER And StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) <> P_TORI_JIKYU Then
                        
                        
                            
                            If CLng(Text1(ptxKINGAKU).Text) >= 0 Then
                                If Format(Text1(ptxUKEIRE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                                    ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                                            CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
                                Else
                                    ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                            CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))
                                End If
                            Else
                                
                                wkKINGAKU = CLng(Text1(ptxKINGAKU).Text) * -1
                                
                                If Format(Text1(ptxUKEIRE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                                    ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                                            CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
                                Else
                                    ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                            CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))
                                End If
                                ZEI = ZEI * -1
                            End If
    
                            Text1(ptxZEI_KIN).Text = Format(ZEI, "#,##0")
                        Else
                            Text1(ptxZEI_KIN).Text = "0"
                        End If
                    
                    End If
                        
                        
                End If
            
            End If
    
    
    
        Case ptxZEI_KIN     '消費税額
    
            If Not NOUKI_MODE Then
    
    
                If Not IsNumeric(Text1(ptxZEI_KIN).Text) Then
                    MsgBox "入力した項目はエラーです。"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxZEI_KIN).Text = Format(CLng(Text1(ptxZEI_KIN).Text), "#,##0")
                        
                        
                End If
            
            End If


' 2012.12.28 Upd >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        Case ptxG_SYUSHI    '収支ｺｰﾄﾞ

            If Trim(GLB_SYUSHI_F) <> "" Then        '対象収支が有る場合のみチェック
                SYUSHI_ON = False
                For i = 0 To UBound(G_SYUSHI_TBL)
                    If Trim(Text1(ptxG_SYUSHI).Text) = G_SYUSHI_TBL(i) Then
                        SYUSHI_ON = True
                        Exit For
                    End If
                Next i

                If Not SYUSHI_ON Then           '2008.10.09
                    'MsgBox "収支違いです。"                     '2016.04.28
'                    MsgBox "収支対象外です。[対象収支]の押下で入力可能収支を確認しく下さい。(PI00041.INI) "
                    MsgBox "PI00041.iniに収支の登録がありません。" & Chr(13) & Chr(10) & "入力可能収支を確認して下さい。"
                    Text1(ptxG_SYUSHI).SetFocus
                    Exit Function
                End If
            End If


            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN03_CD)               '収支名
            Call UniCode_Conv(K0_P_CODE.C_Code, Text1(ptxG_SYUSHI).Text)
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                    'Text1(ptxSYUSHI_NM).Text = StrConv(P_CODEREC.C_NAME, vbUnicode)        '2018.04.09
                    Text1(ptxSYUSHI_NM).Text = StrConv(P_CODEREC.C_RNAME, vbUnicode)        '2018.04.09
                Case BtErrKeyNotFound
                    Text1(ptxSYUSHI_NM).Text = ""
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "商品ﾗﾍﾞﾙｺﾝﾄﾛｰﾙ ﾌｧｲﾙ")
                    Exit Function
            End Select
' 2012.12.28 Upd <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


    End Select



    Error_Check_Proc = False


End Function

Private Function Item_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   画面表示
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim i           As Integer

Dim SUMI_QTY    As Long
Dim MI_QTY      As Long

Dim ZEI         As Long
Dim wkKINGAKU   As Long

Dim wkDATE      As String * 8


    Item_Disp_Proc = True
    
    Call Input_Area_Proc(0)
    
    
    
        
    Text1(ptxORDER_NO).Text = StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode)        '注文№
                                                                                '注文日
    Text1(ptxORDER_DT).Text = Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 7, 2)
        
    Text1(ptxHIN_GAI).Text = StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)          '品番
        
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    
    End Select
    Text1(ptxHIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    If Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                            StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
        Exit Function
    
    End If
    Text1(ptxZAIKO_QTY).Text = Format(SUMI_QTY + MI_QTY, "#0")
        
        
    Text1(ptxG_SHIIRE_KBN).Text = StrConv(P_SHORDER_REC.G_SHIIRE_KBN, vbUnicode)    '仕入区分
    Combo1(pcmbG_SHIIRE_KBN).ListIndex = -1
    For i = 0 To Combo1(pcmbG_SHIIRE_KBN).ListCount - 1
    
        If Text1(ptxG_SHIIRE_KBN).Text = Trim(Left(Right(Combo1(pcmbG_SHIIRE_KBN).List(i), 3), 2)) Then
            Combo1(pcmbG_SHIIRE_KBN).ListIndex = i
            Exit For
        End If
    
    Next i
        
        
    If OSAKA_MODE Then                                                              '使用月 2008.01.10
    
        If Trim(StrConv(P_SHORDER_REC.USE_YM, vbUnicode)) <> "" Then
            Text1(ptxUSE_YM).Text = Mid(StrConv(P_SHORDER_REC.USE_YM, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SHORDER_REC.USE_YM, vbUnicode), 5, 2)
        Else
            Text1(ptxUSE_YM).Text = ""
        End If
    
    End If
        
    Text1(ptxTANTO_CODE).Text = StrConv(P_SHORDER_REC.TANTO_CODE, vbUnicode)        '担当者ｺｰﾄﾞ／名称
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

' 2012.12.28 Upd >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    If Trim(StrConv(P_SHORDER_REC.G_SYUSHI, vbUnicode)) = "" Then               '収支単位←ﾃﾞﾌｫﾙﾄ：品目ﾏｽﾀ値
        Call UniCode_Conv(P_SHORDER_REC.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))
    End If

    Text1(ptxG_SYUSHI).Text = Trim(StrConv(P_SHORDER_REC.G_SYUSHI, vbUnicode))  '収支単位

    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN03_CD)                           '収支名
    Call UniCode_Conv(K0_P_CODE.C_Code, Text1(ptxG_SYUSHI).Text)
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    Select Case sts
        Case BtNoErr
            'Text1(ptxSYUSHI_NM).Text = StrConv(P_CODEREC.C_NAME, vbUnicode)        '2018.04.09
            Text1(ptxSYUSHI_NM).Text = StrConv(P_CODEREC.C_RNAME, vbUnicode)        '2018.04.09
        Case BtErrKeyNotFound
            Text1(ptxSYUSHI_NM).Text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "商品ﾗﾍﾞﾙｺﾝﾄﾛｰﾙ ﾌｧｲﾙ")
            Exit Function
    End Select
' 2012.12.28 Upd <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

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
                                                                                    
                                                                                    '納期予定日
    
    If Trim(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode)) <> "" Then                '2007.09.06
        Text1(ptxY_NOUKI_DT).Text = Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 7, 2)
                                                                                    
    Else
        Text1(ptxY_NOUKI_DT).Text = ""
    End If
                                                                                    
                                                                                    
                                                                                    '回答納期日 2008.01.10
    If OSAKA_MODE Then
        If Trim(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode)) <> "" Then
            Text1(ptxANS_NOUKI_DT).Text = Mid(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 7, 2)
                                                                                        
        Else
            Text1(ptxANS_NOUKI_DT).Text = ""
        End If
    End If
                                                                                    '注文数
    Text1(ptxORDER_QTY).Text = Format(CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)), "#0")
                                                                                    '受入済数
    Text1(ptxUKEIRE_QTY).Text = Format(CLng(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)), "#0")
                                                                                    '単価
    Text1(ptxTANKA).Text = Format(CDbl(StrConv(P_SHORDER_REC.TANKA, vbUnicode)), "#0.00")
                                                                                    
                                                                                    
                                                                                    
'>>>>>>>>>>>>>>>>>>>>>>>>>> 前借情報    2016.09.08
    txtP_NYUKA_QTY.Text = ""
    txtP_NYUKA_DT.Text = ""
    
    
    Call UniCode_Conv(K0_P_NYU.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_P_NYU.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_P_NYU.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
    Call UniCode_Conv(K0_P_NYU.NYUKA_DT, "")


    sts = BTRV(BtOpGetGreater, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
    Select Case sts
        Case BtNoErr
        
            If StrConv(P_NYUREC.JGYOBU, vbUnicode) <> SHIZAI Or _
                StrConv(P_NYUREC.NAIGAI, vbUnicode) <> NAIGAI_NAI Or _
                Trim(StrConv(P_NYUREC.HIN_GAI, vbUnicode)) <> Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)) Then
            Else
                
                
                If Val(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode)) > Val(StrConv(P_NYUREC.SOUSAI_QTY, vbUnicode)) Then
                    wkDATE = StrConv(P_NYUREC.NYUKA_DT, vbUnicode)
                    txtP_NYUKA_DT.Text = Mid(wkDATE, 1, 4) & "/" & Mid(wkDATE, 5, 2) & "/" & Mid(wkDATE, 7, 2)
        
                    If IsNumeric(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode)) Then
                        txtP_NYUKA_QTY.Text = Format(Val(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode)), "#,##0")
                    Else
                        txtP_NYUKA_QTY.Text = ""
                    End If
                Else
                    txtP_NYUKA_QTY.Text = ""
                End If
            End If



        
        
        Case BtErrEOF
        Case Else
            Call File_Error(sts, BtOpGetGreater, "資材　前借データ")
            Exit Function
    
    End Select
    










'>>>>>>>>>>>>>>>>>>>>>>>>>> 前借情報    2016.09.08
                                                                                    '発注ﾛｯﾄ
    Text1(ptxLOT).Text = Format(CLng(StrConv(P_SHORDER_REC.LOT, vbUnicode)), "#0")
                                                                                    '今回受入数
    Text1(ptxKONKAI_UKEIRE_QTY).Text = Format(CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) - _
                                        CLng(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)), "#0")
                                                                                    '注文残
    Text1(ptxZAN_QTY).Text = "0"
                                                                                    '金額
    '2009.11.02
'    Text1(ptxKINGAKU).Text = Format(CDbl(StrConv(P_SHORDER_REC.TANKA, vbUnicode)) * _
'                                    CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text), "#,##0")
    
    Select Case StrConv(P_KANRIREC.SHI_MARUME, vbUnicode)
        Case "0"    '切捨て
            Text1(ptxKINGAKU).Text = Format(ToRoundDown(CCur(CCur(StrConv(P_SHORDER_REC.TANKA, vbUnicode)) * _
                                    CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")
        

        Case "5"    '四捨五入
        
            Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(StrConv(P_SHORDER_REC.TANKA, vbUnicode)) * _
                                    CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")

        
        
        
        
        Case "9"    '切り上げ
    
    
            Text1(ptxKINGAKU).Text = Format(ToRoundUp(CCur(CCur(StrConv(P_SHORDER_REC.TANKA, vbUnicode)) * _
                                    CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")


    
    
        Case Else    '四捨五入
        
            Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(StrConv(P_SHORDER_REC.TANKA, vbUnicode)) * _
                                    CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")
    
    
    End Select
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
                                                                                    '在庫計上有無
    If StrConv(ITEMREC.ZAIKO_F, vbUnicode) <> P_ZAIKO_F_OFF Then
        Check1(chkZAIKO_F).Value = vbChecked
    Else
        Check1(chkZAIKO_F).Value = vbUnchecked
    End If
    
                                                                                    '消費税
    
    '2007.11.01 他ｾﾝﾀｰ、時給は消費税計算しない
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxORDER_CODE).Text)
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            MsgBox "受払先マスタ未登録です。"
            Text1(ptxORDER_CODE).SetFocus
            Exit Function
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
            Exit Function
    End Select
                       
                       
                       
    If StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) <> P_TORI_ANOTHER And StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) <> P_TORI_JIKYU Then
    
    
        If CLng(Text1(ptxKINGAKU).Text) >= 0 Then
            
            If Format(Text1(ptxUKEIRE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                        CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
            Else
                ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                        CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))
            End If
        Else
            
            wkKINGAKU = CLng(Text1(ptxKINGAKU).Text) * -1
            
            If Format(Text1(ptxUKEIRE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                        CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
            Else
                ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                        CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))
            End If
            ZEI = ZEI * -1
        End If
    
        Text1(ptxZEI_KIN).Text = Format(ZEI, "#,##0")
    Else
        Text1(ptxZEI_KIN).Text = "0"
    End If
    
    
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

    
    
    
    '---------------------------------------------------    '資材注文ﾃﾞｰﾀｷｬﾝｾﾙ
    
    Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, Text1(ptxORDER_NO).Text)
    
    Do
    
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
            
        Select Case sts
            Case BtNoErr
            
                
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
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'    Call UniCode_Conv(P_SHORDER_REC.CANCEL_F, P_CANCEL_ON)  'ｷｬﾝｾﾙﾌﾗｸﾞ
'                                                            'ｷｬﾝｾﾙ日時
'    Call UniCode_Conv(P_SHORDER_REC.CANCEL_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    '2012.12.25     上記を下記に変更        M.T
    '                                            受入数の有無でフラグが異なる！
    If CLng(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)) <> 0 Then
        Call UniCode_Conv(P_SHORDER_REC.KAN_F, P_KAN_ON)                    '完了ﾌﾗｸﾞ
                                                                            '完了日
        Call UniCode_Conv(P_SHORDER_REC.KAN_DT, Format(Now, "YYYYMMDD"))
    
    Else
    
        Call UniCode_Conv(P_SHORDER_REC.CANCEL_F, P_CANCEL_ON)  'ｷｬﾝｾﾙﾌﾗｸﾞ
                                                                'ｷｬﾝｾﾙ日時
        Call UniCode_Conv(P_SHORDER_REC.CANCEL_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    End If
    '           ここまで
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                                                            
                                                            '更新日時
    Call UniCode_Conv(P_SHORDER_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpUpdate, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
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
                Call File_Error(sts, BtOpUpdate, "資材注文ﾃﾞｰﾀ")
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
    
    
    Cancel_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    
    
    Cancel_Proc = True

End Function

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                  資材注文ﾃﾞｰﾀ更新
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer

Dim SEQNO           As Integer


    Update_Proc = True
                                        
                                        
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
    
    
    '---------------------------------------------------    '資材注文データ更新
    If Input_Mode Then
    
                                            
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
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "管理マスタ")
                        GoTo Abort_Tran
                
                End Select
            
            
            Loop
        
            '注文書№＋１
        
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
                                                                                
                                                                                    
                                                                                
                                                                                '注文№
        Call UniCode_Conv(P_SHORDER_REC.ORDER_NO, StrConv(P_KANRIREC.ORDER_NO, vbUnicode))
                                                                                '注文日
        Call UniCode_Conv(P_SHORDER_REC.ORDER_DT, Format(Text1(ptxORDER_DT).Text, "YYYYMMDD"))
        
        Call UniCode_Conv(P_SHORDER_REC.Print_datetime, "")                     '発行日時
        Call UniCode_Conv(P_SHORDER_REC.TANTO_CODE, Text1(ptxTANTO_CODE).Text)  '担当者

' 2012.12.28 Upd >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Call UniCode_Conv(P_SHORDER_REC.G_SYUSHI, Trim(Text1(ptxG_SYUSHI).Text))    '収支単位
' 2012.12.28 Upd <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

        Call UniCode_Conv(P_SHORDER_REC.JGYOBU, SHIZAI)                         '事業部（＝資材）
        Call UniCode_Conv(P_SHORDER_REC.NAIGAI, NAIGAI_NAI)                     '国内外
        Call UniCode_Conv(P_SHORDER_REC.HIN_GAI, Text1(ptxHIN_GAI).Text)        '品番
        Call UniCode_Conv(P_SHORDER_REC.ORDER_CODE, Text1(ptxORDER_CODE).Text)  '注文先ｺｰﾄﾞ
        Call UniCode_Conv(P_SHORDER_REC.DELI_CODE, "")                          '納入先ｺｰﾄﾞ
        Call UniCode_Conv(P_SHORDER_REC.ORDER_QTY, Format(CDbl(Text1(ptxORDER_QTY).Text), _
                                                                "00000000.00")) '注文数
        Call UniCode_Conv(P_SHORDER_REC.Y_NOUKI_DT, Format(CDate(Text1(ptxY_NOUKI_DT).Text), _
                                                                "YYYYMMDD"))    '予定納期
        
        
        If Trim(Text1(ptxANS_NOUKI_DT).Text) <> "" Then                         '回答納期日 2008.01.10
            Call UniCode_Conv(P_SHORDER_REC.ANS_NOUKI_DT, Format(CDate(Text1(ptxANS_NOUKI_DT).Text), _
                                                                    "YYYYMMDD"))
        Else
'''2008.10.18            Call UniCode_Conv(P_SHORDER_REC.ANS_NOUKI_DT, "")
        
        
            Call UniCode_Conv(P_SHORDER_REC.ANS_NOUKI_DT, Format(Text1(ptxUKEIRE_DT).Text, "YYYYMMDD"))     '2008.10.18
        End If
        
        
        If Trim(Text1(ptxUSE_YM).Text) <> "" Then                               '使用月 2008.01.10
            Call UniCode_Conv(P_SHORDER_REC.USE_YM, Format(CDate(Text1(ptxUSE_YM).Text & "/01"), _
                                                                    "YYYYMMDD"))
        Else
            Call UniCode_Conv(P_SHORDER_REC.USE_YM, "")
        End If
        
        
        
        Call UniCode_Conv(P_SHORDER_REC.TANKA, Format(CDbl(Text1(ptxTANKA).Text), _
                                                                "00000000.00")) '単価
        Call UniCode_Conv(P_SHORDER_REC.LOT, "00000001")
    
        Call UniCode_Conv(P_SHORDER_REC.KAN_F, P_KAN_ON)                        '完了ﾌﾗｸﾞ（即完了）
        Call UniCode_Conv(P_SHORDER_REC.KAN_DT, Format(Now, "YYYYMMDD"))        '完了日
        Call UniCode_Conv(P_SHORDER_REC.BUNNOU_CNT, "01")                       '受入回数
        Call UniCode_Conv(P_SHORDER_REC.UKEIRE_QTY, Format(CDbl(Text1(ptxKONKAI_UKEIRE_QTY).Text), "00000000.00"))
    
        Call UniCode_Conv(P_SHORDER_REC.CANCEL_F, P_CANCEL_OFF)                 'ｷｬﾝｾﾙﾌﾗｸﾞ
        Call UniCode_Conv(P_SHORDER_REC.CANCEL_DATETIME, "")                    'ｷｬﾝｾﾙ日時
    
        Call UniCode_Conv(P_SHORDER_REC.PRINT_F, P_PRINT_ON)                    '印刷ﾌﾗｸﾞ(印刷済とする)
    
        Call UniCode_Conv(P_SHORDER_REC.WS_NO, WS_NO)                           '入力端末
                                                                                '仕入区分
        Call UniCode_Conv(P_SHORDER_REC.G_SHIIRE_KBN, Text1(ptxG_SHIIRE_KBN).Text)
        
        '品目ﾏｽﾀ読込み
        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                MsgBox "品目マスタが他端末で変更されました。更新処理を中止します。"
                GoTo Abort_Tran
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "資材注文ﾃﾞｰﾀ")
                GoTo Abort_Tran
        End Select
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
    
    
        Do
            
            DoEvents
            
            sts = BTRV(BtOpInsert, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
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
                    Call File_Error(sts, BtOpInsert, "資材注文ﾃﾞｰﾀ")
                    GoTo Abort_Tran
            End Select
        
        Loop
    
    
    
    Else
    
        '資材注文データ処理
        Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, Text1(ptxORDER_NO).Text)
        
        Do
        
            sts = BTRV(BtOpGetEqual + BtSNoWait, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
                
            Select Case sts
                Case BtNoErr
                
                    
                    Exit Do
                
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_SHORDER.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        GoTo Abort_Tran
                    End If
                
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "資材注文データ")
                    GoTo Abort_Tran
            End Select
    
        Loop
        
        If CDbl(Text1(ptxZAN_QTY).Text) = 0 Then
            Call UniCode_Conv(P_SHORDER_REC.KAN_F, P_KAN_ON)                   '完了ﾌﾗｸﾞ
            Call UniCode_Conv(P_SHORDER_REC.KAN_DT, Format(Now, "YYYYMMDD"))   '完了日
            '2008.10.18
            If Not IsNumeric(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode)) Or Trim(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode)) = "" Then
                Call UniCode_Conv(P_SHORDER_REC.ANS_NOUKI_DT, Format(Text1(ptxUKEIRE_DT).Text, "YYYYMMDD"))
            End If
            '2008.10.18
            If CInt(StrConv(P_SHORDER_REC.BUNNOU_CNT, vbUnicode)) = 0 Then     '分納回数
            Else
                Call UniCode_Conv(P_SHORDER_REC.BUNNOU_CNT, Format(CInt(CInt(StrConv(P_SHORDER_REC.BUNNOU_CNT, vbUnicode)) + 1), "000"))
            End If
        End If


' 2012.12.28 Upd >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        Call UniCode_Conv(P_SHORDER_REC.G_SYUSHI, Trim(Text1(ptxG_SYUSHI).Text))    '収支単位
' 2012.12.28 Upd <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


        Call UniCode_Conv(P_SHORDER_REC.UKEIRE_QTY, Format(CDbl(CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)) + CDbl(Text1(ptxKONKAI_UKEIRE_QTY).Text)), "00000000.00"))
                                                            '更新日時
        Call UniCode_Conv(P_SHORDER_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
        Do
            
            DoEvents
            
            sts = BTRV(BtOpUpdate, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
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
                    Call File_Error(sts, BtOpUpdate, "資材注文ﾃﾞｰﾀ")
                    GoTo Abort_Tran
            End Select
        
        Loop
    End If
    
    SEQNO = 0
    
    
    
    '資材受入履歴ﾃﾞｰﾀ処理
    Call UniCode_Conv(K0_P_SHUKEIRE.ORDER_NO, StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode))
    Call UniCode_Conv(K0_P_SHUKEIRE.SEQNO, "")
    
    com = BtOpGetGreater
    
    Do
    
        sts = BTRV(com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K0_P_SHUKEIRE, Len(K0_P_SHUKEIRE), 0)
            
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode) <> StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "資材受入履歴")
                GoTo Abort_Tran
        End Select
        
        
        
        SEQNO = SEQNO + 1
        
        
        com = BtOpGetNext
        
    Loop
        
                                                                                '注文№
    Call UniCode_Conv(P_SHUKEIRE_REC.ORDER_NO, StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode))                                                                                         '受入日
                                                                                '注文先
    Call UniCode_Conv(P_SHUKEIRE_REC.ORDER_CODE, StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode))
                                                                                '受入日
    Call UniCode_Conv(P_SHUKEIRE_REC.UKEIRE_DT, Format(Text1(ptxUKEIRE_DT).Text, "YYYYMMDD"))
                                                                                '受入数量
    Call UniCode_Conv(P_SHUKEIRE_REC.UKEIRE_QTY, Format(CDbl(Text1(ptxKONKAI_UKEIRE_QTY).Text), "00000000.00"))
                                                                                '受入単価
    Call UniCode_Conv(P_SHUKEIRE_REC.UKEIRE_TANKA, Format(CDbl(Text1(ptxTANKA).Text), _
                                                                "00000000.00"))
                                                                                '受入金額
'    Call UniCode_Conv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, Format(CLng(CDbl(Text1(ptxKONKAI_UKEIRE_QTY).Text) * _
'                                                        CDbl(StrConv(P_SHORDER_REC.TANKA, vbUnicode))), "00000000"))
        
    Call UniCode_Conv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, Format(CLng(Text1(ptxKINGAKU).Text), "00000000"))
    Call UniCode_Conv(P_SHUKEIRE_REC.ZEI_KIN, Format(CLng(Text1(ptxZEI_KIN).Text), "00000000"))
        
        
        
    If CDbl(Text1(ptxZAN_QTY).Text) = 0 Then
        Call UniCode_Conv(P_SHUKEIRE_REC.LAST_F, P_UKEIRE_END)
    Else
        Call UniCode_Conv(P_SHUKEIRE_REC.LAST_F, P_UKEIRE_CON)
    End If
                                                                                '計上年月
    Call UniCode_Conv(P_SHUKEIRE_REC.KEIJYO_YM, Mid(Text1(ptxKEIJYO_YM), 1, 4) & Mid(Text1(ptxKEIJYO_YM), 6, 2))
        
    Call UniCode_Conv(P_SHUKEIRE_REC.FILLER, "")
                                                        '更新日時
    Call UniCode_Conv(P_SHUKEIRE_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    Do
            
        SEQNO = SEQNO + 1
                                                '追番
        Call UniCode_Conv(P_SHUKEIRE_REC.SEQNO, Format(SEQNO, "000"))
            
        DoEvents
            
        sts = BTRV(BtOpInsert, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K0_P_SHUKEIRE, Len(K0_P_SHUKEIRE), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrDuplicates
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_SHUKEIRE.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "資材受入履歴")
                GoTo Abort_Tran
        End Select
        
    Loop
    '------------------------------------------------ POS入荷処理
    If POS_UMU Then
        If Check1(chkZAIKO_F).Value = vbChecked Then
    
            If POS_NYUKA_Update_Proc("  ", "  ", "  ", "  ") Then
                GoTo Abort_Tran
            End If
        End If
    
    Else
        
        If Check1(chkZAIKO_F).Value = vbChecked Then
    
            'POSなしは標準棚番に在庫計上    2006.04.24
    
    
            If POS_NYUKA_Update_Proc(StrConv(ITEMREC.ST_SOKO, vbUnicode), _
                                        StrConv(ITEMREC.ST_RETU, vbUnicode), _
                                        StrConv(ITEMREC.ST_REN, vbUnicode), _
                                        StrConv(ITEMREC.ST_DAN, vbUnicode)) Then
                GoTo Abort_Tran
            End If
        End If
        
    
    
    End If


End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    

End Function

Private Function NOUKI_Update_Proc() As Integer
'----------------------------------------------------------------------------
'                  資材注文ﾃﾞｰﾀ更新
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer




    NOUKI_Update_Proc = True
                                        
                                        
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
    '---------------------------------------------------    '資材注文データ更新
    '資材注文データ処理
    Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, Text1(ptxORDER_NO).Text)
    
    Do
    
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
            
        Select Case sts
            Case BtNoErr
            
                
                Exit Do
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_SHORDER.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            
            
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "資材注文データ")
                GoTo Abort_Tran
        End Select

    Loop
    
                                                        '予定納期
    Call UniCode_Conv(P_SHORDER_REC.Y_NOUKI_DT, Format(Text1(ptxY_NOUKI_DT).Text, "YYYYMMDD"))
                                                        
    If Trim(Text1(ptxANS_NOUKI_DT).Text) = "" Then      '回答納期     2008.01.10
        Call UniCode_Conv(P_SHORDER_REC.ANS_NOUKI_DT, "")
    Else
        Call UniCode_Conv(P_SHORDER_REC.ANS_NOUKI_DT, Format(CDate(Text1(ptxANS_NOUKI_DT).Text), "YYYYMMDD"))
    End If
                                                        
    If Trim(Text1(ptxUSE_YM).Text) = "" Then            '使用月       2008.01.10
        Call UniCode_Conv(P_SHORDER_REC.USE_YM, "")
    Else
        Call UniCode_Conv(P_SHORDER_REC.USE_YM, Format(CDate(Text1(ptxUSE_YM).Text & "/01"), "YYYYMMDD"))
    End If
                                                        '更新日時
    Call UniCode_Conv(P_SHORDER_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpUpdate, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<P_SHORDER.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
                    If sts Then
                        Call File_Error(sts, BtOpUnlock, "資材注文ﾃﾞｰﾀ")
                    End If
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "資材注文ﾃﾞｰﾀ")
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
    
    Call UniCode_Conv(P_SHORDER_REC.ORDER_NO, "")       '2016.01.18
    
    NOUKI_Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    

End Function



Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    Select Case Index
        Case pcmbG_SHIIRE_KBN   '仕入区分
            Text1(ptxG_SHIIRE_KBN).Text = Trim(Left(Right(Combo1(pcmbG_SHIIRE_KBN).Text, 3), 2))
        Case pcmbORDER          '注文先
            Text1(ptxORDER_CODE).Text = Trim(Right(Combo1(pcmbORDER).Text, 5))
        Case pcmbDELI           '納入先
            Text1(ptxDELI_CODE).Text = Trim(Right(Combo1(pcmbDELI).Text, 5))
    End Select
    
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    
    Select Case Index
        Case pcmbG_SHIIRE_KBN   '仕入区分
            Text1(ptxG_SHIIRE_KBN).Text = Trim(Left(Right(Combo1(pcmbG_SHIIRE_KBN).Text, 3), 2))
        Case pcmbORDER          '注文先
            Text1(ptxORDER_CODE).Text = Trim(Right(Combo1(pcmbORDER).Text, 5))
        Case pcmbDELI           '納入先
            Text1(ptxDELI_CODE).Text = Trim(Right(Combo1(pcmbDELI).Text, 5))
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans         As Integer
Dim i           As Integer


Dim sts         As Integer

Dim KAN_yn      As Integer                          '2012.12.25     受入数≠０の時のｙｎ


    Select Case Index
        Case P_CMD_Upd        '更新
            
            
'            For i = ptxORDER_NO To ptxKINGAKU
            For i = ptxORDER_NO To ptxG_SYUSHI
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
            
            
            If NOUKI_MODE Then
            
                ans = MsgBox("納期変更しますか？", vbYesNo + vbQuestion, "確認入力")
                If ans = vbYes Then
                    
                    
                    
                    Call Input_Lock
                    TDBGrid1.Enabled = False
                    
                    If NOUKI_Update_Proc() Then
                        Unload Me
                    End If
                    
''                    If List_Disp_Proc() Then
''                        Unload Me
''                    End If
                    
                    
                    Call Input_UnLock
                    TDBGrid1.Enabled = True
                    
                    
                    If Init_Proc() Then
                        Unload Me
                    End If
                
                    
                    
                    
                    Text1(ptxORDER_NO).SetFocus
                
                Else
                
                    If OSAKA_MODE Then
                        Text1(ptxUSE_YM).SetFocus
                    
                    Else
                        Text1(ptxY_NOUKI_DT).SetFocus
                
                    End If
                
                End If
            
            
            Else
                ans = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
                If ans = vbYes Then
                    
                    
                    Call Input_Lock
                    TDBGrid1.Enabled = False
                    
                    If Update_Proc() Then
                        Unload Me
                    End If
                    
''                    If List_Disp_Proc() Then
''                        Unload Me
''                    End If
                    
                    
                    Call Input_UnLock
                    TDBGrid1.Enabled = True
                    
                    
                    If Init_Proc() Then
                        Unload Me
                    End If
                
                    Text1(ptxORDER_NO).SetFocus
                Else
                    
                    
''''''''''''''''''''''''''' 2011.04.13  確認「いいえ」時の戻り先変更
''                    If OSAKA_MODE Then
''                        Text1(ptxUSE_YM).SetFocus
''
''                    Else
''                        Text1(ptxY_NOUKI_DT).SetFocus
''
''                    End If
                
                
                    If Text1(ptxHIN_GAI).Locked Then
                        Text1(ptxG_SHIIRE_KBN).SetFocus
                    Else
                        Text1(ptxHIN_GAI).SetFocus
                
                    End If
                
''''''''''''''''''''''''''' 2011.04.13
                
                End If
            
            
            
            End If
            
        Case P_CMD_DEL                      '削除
            
            '資材注文ﾃﾞｰﾀ
            Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, Text1(ptxORDER_NO))
            sts = BTRV(BtOpGetEqual, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
                
            Select Case sts
                Case BtNoErr
                                   
                    If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = P_KAN_ON Then
                        MsgBox "完了登録済です。"
                        Text1(ptxORDER_NO).SetFocus
                        Exit Sub
                    End If
                
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'                    If CLng(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)) <> 0 Then
'                        MsgBox "仕入実績が有ります。"
'                        Text1(ptxORDER_NO).SetFocus
'                        Exit Sub
'                    End If
                    '           2012.12.25 上記If文を下記に変更     M.T
                    If CLng(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)) <> 0 Then
                        KAN_yn = MsgBox("仕入実績が有ります。" & Chr(13) & Chr(10) & _
                                        "  完了にしますか？", vbYesNo + vbDefaultButton2 + vbQuestion, "確認入力")
                        If KAN_yn <> vbYes Then
                            Text1(ptxORDER_NO).SetFocus
                            Call Text1_GotFocus(ptxORDER_NO)
                            Exit Sub
                        End If
                        
                    End If
                    '>>>>>>>>>>>        ここまで
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

                    If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
                        MsgBox "キャンセル済です。"
                        Text1(ptxORDER_NO).SetFocus
                        Exit Sub
                    End If
                
                Case BtErrKeyNotFound
                    MsgBox "資材注文ﾃﾞｰﾀ未登録です。"
                    Text1(ptxORDER_NO).SetFocus
                    Exit Sub
                Case Else
                    Unload Me
            End Select
        
        
        
            ans = MsgBox("ｷｬﾝｾﾙしますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                
                Call Input_Lock
                TDBGrid1.Enabled = False
                
                
                If Cancel_Proc() Then
                    Unload Me
                End If
                
                
                If List_Disp_Proc() Then
                    Unload Me
                End If
                
                Call Input_UnLock
                TDBGrid1.Enabled = True
                
                
                If Init_Proc() Then
                    Unload Me
                End If
            
            
            
            End If
            
            Text1(ptxORDER_NO).SetFocus
    
        Case P_CMD_DSP                      '検索/表示
            
            If List_Disp_Proc() Then
                Unload Me
            End If
        
        
        Case cmdNOUKI
        
            If NOUKI_MODE Then
                Call Input_Area_Set(0)
                NOUKI_MODE = False
                Text1(ptxUKEIRE_DT).SetFocus
            Else
                Call Input_Area_Set(1)
                NOUKI_MODE = True
                
                If OSAKA_MODE Then
                    Text1(ptxUSE_YM).SetFocus
                
                Else
                    Text1(ptxY_NOUKI_DT).SetFocus
            
                End If
            End If
        
        Case P_CMD_OUT                      'ﾃﾞｰﾀ出力
        
        Case P_CMD_PRT                      '印刷
        Case P_CMD_End                      '終了
    
            Unload Me
    
    End Select

End Sub



Private Sub Command2_Click()

Dim i       As Integer
Dim wkMSG   As String

    wkMSG = ""
    
    For i = 0 To UBound(G_SYUSHI_TBL)
        wkMSG = wkMSG & "[" & G_SYUSHI_TBL(i) & "]" & Chr(13) & Chr(10)
    Next i

    MsgBox wkMSG




End Sub

Private Sub Form_DblClick()
'   PrintForm
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
Dim sBuffer As String

'    If App.PrevInstance Then
'        Beep
'        MsgBox "同一プログラム実行中です。"
'        End
'    End If

    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "資材仕入処理", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)



    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        WS_NO = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        WS_NO = "???"
    End If

                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止します。"
        End
    End If
    LOG_F = RTrim(c)
                                
'>>>>>>>>>>>>   P_SYS.INI --> PI00041.INI 2016.01.14
                                'POSｼｽﾃﾑ有無の取り込み
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "POS_UMU", "P_SYS", c) Then
    If GetIni(StrConv(App.EXEName, vbUpperCase), "POS_UMU", StrConv(App.EXEName, vbUpperCase), c) Then
        POS_UMU = False
    Else
        If RTrim(c) = "0" Then
            POS_UMU = False
        Else
            POS_UMU = True
        End If
    End If
'''     POSなしでも在庫計上する2006.04.24
'''    If POS_UMU Then
                                '入荷仮想倉庫の取り込み
'        If GetIni(StrConv(App.EXEName, vbUpperCase), "NYUKA_SOKO", "P_SYS", c) Then
        If GetIni(StrConv(App.EXEName, vbUpperCase), "NYUKA_SOKO", StrConv(App.EXEName, vbUpperCase), c) Then
            Beep
            MsgBox "入荷仮想倉庫番号の獲得に失敗しました。処理を中止します。"
            End
        End If
        KASO_NYUKA = RTrim(c)
    
    
                                '「資材通常入荷」の要因
'        If GetIni(StrConv(App.EXEName, vbUpperCase), "YOIN_TU_NYUKA", "P_SYS", c) Then
        If GetIni(StrConv(App.EXEName, vbUpperCase), "YOIN_TU_NYUKA", StrConv(App.EXEName, vbUpperCase), c) Then
            Call LOG_OUT(LOG_F, "[P_SYS.INI][" & StrConv(App.EXEName, vbUpperCase) & "[YOIN_TU_NYUKA] READ ERROR")
            MsgBox "資材通常入荷用要因の獲得に失敗しました。処理を中止します。"
            End
        End If
        P_YOIN_TU_NYUKA = Trim(c)
                                '「資材前借相殺」の要因
'        If GetIni(StrConv(App.EXEName, vbUpperCase), "YOIN_MAE_SOUSAI", "P_SYS", c) Then
        If GetIni(StrConv(App.EXEName, vbUpperCase), "YOIN_MAE_SOUSAI", StrConv(App.EXEName, vbUpperCase), c) Then
'            Call LOG_OUT(LOG_F, "[P_SYS.INI][" & StrConv(App.EXEName, vbUpperCase) & "[YOIN_MAE_SOUSAI] READ ERROR")
            Call LOG_OUT(LOG_F, "[" & StrConv(App.EXEName, vbUpperCase) & "][" & StrConv(App.EXEName, vbUpperCase) & "[YOIN_MAE_SOUSAI] READ ERROR")
            MsgBox "資材前借相殺用要因の獲得に失敗しました。処理を中止します。"
            End
        End If
        P_YOIN_MAE_SOUSAI = Trim(c)
    
    
                                    '履歴メモ取り込み
'        If GetIni(App.EXEName, "MEMO", "P_SYS", c) Then
        If GetIni(App.EXEName, "MEMO", App.EXEName, c) Then
            MEMO_TEXT = ""
        Else
            MEMO_TEXT = RTrim(c)
        End If
    
    
'''    End If
                                
                                
                                '回答納期日の入力有無 '2008.01.10
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
                                
                                
    PI000411.Caption = PI000411.Caption & LAST_UPDATE_DAY$
                                
                                
                                
    Label1(plblANS_NOUKI_DT).Visible = OSAKA_MODE
    Text1(ptxANS_NOUKI_DT).Visible = OSAKA_MODE
    Text1(ptxANS_NOUKI_DT).TabStop = OSAKA_MODE
                                    
    Label1(plblUSE_YM).Visible = OSAKA_MODE
    Text1(ptxUSE_YM).Visible = OSAKA_MODE
    Text1(ptxUSE_YM).TabStop = OSAKA_MODE
                                    
    TDBGrid1.Columns(colANS_NOUKI_DT).Visible = OSAKA_MODE
                                
                                
                                
                                '対象収支の獲得 2007.11.13
    
    If Trim(GLB_SYUSHI_F) = "" Then
    
    
        Command2.Visible = False
    
    Else
    
'        If GetIni(StrConv(App.EXEName, vbUpperCase), GLB_SYUSHI_F, "P_SYS", c) Then
        If GetIni(StrConv(App.EXEName, vbUpperCase), GLB_SYUSHI_F, App.EXEName, c) Then
            Beep
            MsgBox "対象収支の獲得に失敗しました。処理を中止して下さい。"
            End
        End If
    
        G_SYUSHI_TBL = Split(Trim(c), ",", -1)
    End If
                                
                                
                                
                                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    表示件数    2016.01.14
    If GetIni(App.EXEName, "LIST_MAX", App.EXEName, c) Then
        LIST_MAX = 0
    Else
        LIST_MAX = Val(RTrim(c))
    End If
                                
                                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    前借表示    2016.09.08
    
    If GetIni(App.EXEName, "P_NYUKA_DSP", App.EXEName, c) Then
        P_NYUKA_DSP = 0
    Else
        P_NYUKA_DSP = Val(RTrim(c))
    
    End If
                                
    If P_NYUKA_DSP = 1 Then
        LBLP_NYUKA_DT.Visible = True
        txtP_NYUKA_DT.Visible = True
    
        lblP_NYUKA_QTY.Visible = True
        txtP_NYUKA_QTY.Visible = True
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    前借表示    2016.09.08
                                
                                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    上下限設定  受入日　2017.04.25
    If GetIni(App.EXEName, "UKEIRE_DT", App.EXEName, c) Then
        UKEIRE_DT = 0
    Else
        UKEIRE_DT = Val(RTrim(c))
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    上下限設定  計上年月　2017.04.25
    If GetIni(App.EXEName, "KEIJYO_YM", App.EXEName, c) Then
        KEIJYO_YM = 0
    Else
        KEIJYO_YM = Val(RTrim(c))
    End If
                                
                                
                                
                                '棚マスタＯＰＥＮ
    If TANA_Open(BtOpenNomal) Then
        Unload Me
    End If
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
                                '資材受入履歴ﾃﾞｰﾀＯＰＥＮ
    If P_SHUKEIRE_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '在庫ﾃﾞｰﾀＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'ｺｰﾄﾞﾏｽﾀＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '資材前借ﾃﾞｰﾀ
    If P_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    '---------------------------    POSﾘﾝｸ用ﾌｧｲﾙ
                                '品目マスタＯＰＥＮ（データ更新用）
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '要因マスタＯＰＥＮ
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '発番マスタＯＰＥＮ
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '入荷予定データファイルＯＰＥＮ
    If Y_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫移動歴ＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
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
        
    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc
    
    '仕入区分のセット
    If Code_Set_Proc(pcmbG_SHIIRE_KBN, P_KBN01_CD, 0) Then
        Unload Me
    End If
    
    
    '注文先
    If Ukeharai_Set_Proc(pcmbORDER) Then
        Unload Me
    End If
    '納入先
    If Ukeharai_Set_Proc(pcmbDELI) Then
        Unload Me
    End If
    
    
    '画面初期設定
    If Init_Proc() Then
        Unload Me
    End If

    If List_Disp_Proc() Then
        Unload Me
    End If

    Text1(ptxKEIJYO_YM).Text = Left(Format(Now, "YYYY/MM/DD"), 7)   '2007.08.02

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
                                            
                                            
                                            
                                            '棚マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "棚マスタ")
        End If
    End If
                                            
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
                                            '資材受入履歴ﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K0_P_SHUKEIRE, Len(K0_P_SHUKEIRE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材受入履歴ﾃﾞｰﾀ")
        End If
    End If
                                            '在庫ﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫ﾃﾞｰﾀ")
        End If
    End If
                                            'ｺｰﾄﾞﾏｽﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ｺｰﾄﾞﾏｽﾀ")
        End If
    End If
    '-------------------------------------- POSﾘﾝｸ情報
                                            '品目マスタ（データ更新用）ＣＬＯＳＥ
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '要因マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "要因マスタ")
        End If
    End If
                                            '発番マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "発番マスタ")
        End If
    End If
                                            '入荷予定データファイルＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "入荷予定データファイル")
        End If
    End If
                                            '在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, Y_NYUREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PI000411 = Nothing

    End
End Sub

Private Sub TDBGrid1_DblClick()
Dim sts As Integer
    
    Text1(ptxORDER_NO).Text = SHORDER(TDBGrid1.Bookmark, colORDER_NO)
    '資材注文データのチェック
    sts = P_SHORDER_Read_Proc()
    Select Case sts
        Case False, BtNoErr
                    
            If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = P_PRINT_ON Then
                MsgBox "他端末で書き換えられています。"
                TDBGrid1.SetFocus
                Exit Sub
            End If
            
            
            '2007.09.06 予定納期未設定は受入不可
'            If Trim(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode)) = "" Then
'                MsgBox "予定納期が設定されていません。"
'                TDBGrid1.SetFocus
'                Exit Sub
'            End If
            
            
            
            Save_UKEIRE_QTY = 0
        
        Case BtErrKeyNotFound
            MsgBox "他端末で書き換えられています。"
            TDBGrid1.SetFocus
            Exit Sub
        Case Else
            Exit Sub
    End Select
    
    Text1(ptxUKEIRE_DT).SetFocus
    

End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)


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

Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If


    Select Case Index
    
        Case ptxTANKA
    
            wkTANKA = Trim(Text1(ptxTANKA).Text)
    
        Case ptxKONKAI_UKEIRE_QTY
            wkUKEIRE_QTY = Trim(Text1(ptxKONKAI_UKEIRE_QTY).Text)
    
    End Select
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
Dim fsw
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    fsw = False
    Select Case Index                                                               '2013.10.08
        Case ptxHIN_GAI                                                             '2013.10.08
            Text1(Index).Text = StrConv(RTrim(Text1(Index).Text), vbUpperCase)      '2013.10.08
            fsw = True
    End Select                                                                      '2013.10.08
        
    If Error_Check_Proc(Index) Then     'エラーチェック
        Exit Sub
    End If
            
        
    If fsw Then                             '2013.10.08
        Text1(ptxG_SHIIRE_KBN).SetFocus     '2013.10.08
    Else
        Call Tab_Ctrl(Shift)        '移動
    End If
End Sub


Private Function Init_Proc() As Integer
'----------------------------------------------------------------------------
'                   入力画面の初期設定
'----------------------------------------------------------------------------
Dim i       As Integer
Dim sts     As Integer


    Init_Proc = True
    
    
    
    For i = ptxORDER_NO To ptxZEI_KIN
        
        If i = ptxKEIJYO_YM Then    '2007.08.02
        Else
            Text1(i).Text = ""
        End If
    Next i
    '受入日＝当日
    Text1(ptxUKEIRE_DT).Text = Format(Now, "YYYY/MM/DD")
    '計上月＝当月
'2007.08.02    Text1(ptxKEIJYO_YM).Text = Left(Format(Now, "YYYY/MM/DD"), 7)


    Combo1(pcmbG_SHIIRE_KBN).ListIndex = 0


    For i = pcmbORDER To pcmbDELI
        
        Combo1(i).ListIndex = -1
    
    Next i


    Check1(chkZAIKO_F).Value = vbUnchecked


'>>>>>>>>>  前借項目クリアー    2016.09.08
    txtP_NYUKA_DT.Text = ""
    txtP_NYUKA_QTY.Text = ""


'>>>>>>>>>  前借項目クリアー    2016.09.08


    'ｿｰﾄ情報の初期化
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0             'ﾃﾞﾌｫﾙﾄ昇順
    Next i

    Sort_Tbl(colHIN_NAME) = 9       'ｿｰﾄ除外


    NOUKI_MODE = False
    Call Input_Area_Set(0)

    Call UniCode_Conv(ITEMREC.JGYOBU, "")
    Call UniCode_Conv(ITEMREC.NAIGAI, "")
    Call UniCode_Conv(ITEMREC.HIN_GAI, "")
    Save_UKEIRE_QTY = 0
    

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

Private Function P_SHORDER_Read_Proc() As Integer
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
        
        
        
        
        Case Else
            P_SHORDER_Read_Proc = sts
            Exit Function
    
    End Select
    
    
    If Item_Disp_Proc() Then
        Exit Function
    End If
    
    P_SHORDER_Read_Proc = False
        
    

End Function

Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'           資材注文ﾃﾞｰﾀの表示
'
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim com     As Integer

Dim Row     As Long

Dim SYUSHI_ON   As Boolean  '2008.10.09
Dim i           As Integer  '2008.10.09

    List_Disp_Proc = True
    PI000411.MousePointer = vbHourglass '2016.01.14
    PI000411.Enabled = False            '2016.01.14
    
    '>>>>>>>>>>>>>>>>>>>>>> ステータスウィンドウを作成する  2016.01.14
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "資材仕入処理 表示処理開始", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
    
    
    Set SHORDER = Nothing
    Tbl_Set_F = False
    
    
        
    
    
    
'    com = BtOpGetFirst         '2016.01.14
    
    
    Call UniCode_Conv(K3_P_SHORDER.KAN_F, "0")              '2016.01.14
    Call UniCode_Conv(K3_P_SHORDER.ORDER_DT, "99999999")    '2016.01.14
    Call UniCode_Conv(K3_P_SHORDER.ORDER_CODE, "99999")     '2016.01.14
    
    com = BtOpGetLess           '2016.01.14
    
    Row = Min_Row - 1
       
    Do
    
        DoEvents
    
'        sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K2_P_SHORDER, Len(K2_P_SHORDER), 2)  '2016.01.14
        sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K3_P_SHORDER, Len(K3_P_SHORDER), 3)   '2016.01.14
            
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                PI000411.Enabled = True         '2016.01.14
                Call File_Error(sts, com, "資材注文ﾃﾞｰﾀ")
                Exit Function
        End Select
If StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode) = "26299" Then
    Debug.Print
End If

    
        If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Or _
            StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = P_KAN_ON Then
        Else
    
    
            SYUSHI_ON = False               '2008.10.09
            If GLB_SYUSHI_F = "" Then       '2008.10.09
                SYUSHI_ON = True
            Else
                SYUSHI_ON = False
                
                For i = 0 To UBound(G_SYUSHI_TBL)
                
                    If Trim(StrConv(P_SHORDER_REC.G_SYUSHI, vbUnicode)) = G_SYUSHI_TBL(i) Then
                        SYUSHI_ON = True
                        Exit For
                    End If
                
                
                Next i
            End If

    
            If SYUSHI_ON Then
    
    
                Row = Row + 1
                
                If LIST_MAX <> 0 Then                   '2016.01.14
                    If Row > LIST_MAX Then              '2016.01.14
                        Exit Do                         '2016.01.14
                    End If                              '2016.01.14
                End If                                  '2016.01.14
                
                
                If Grid_Set_Proc(Row) Then
                    PI000411.Enabled = True         '2016.01.14
                    Exit Function
                End If
                Tbl_Set_F = True
            End If
        
        End If
        
'        com = BtOpGetNext                          '2016.01.14
        com = BtOpGetPrev                           '2016.01.14

    Loop
    
    Set TDBGrid1.Array = SHORDER
            
    If Row <> (Min_Row - 1) Then
        SHORDER.QuickSort Min_Row, SHORDER.UpperBound(1), colORDER_NO, XORDER_ASCEND, XTYPE_STRING
    End If
            
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    
    
    
    '>>>>>>>>>>>>>>>>>>>>>> ステータスウィンドウを作成する  2016.01.14
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "資材仕入処理 表示処理終了", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
    
    PI000411.Enabled = True         '2016.01.14
    
    PI000411.MousePointer = vbDefault
    List_Disp_Proc = False
    


End Function

Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           資材注文ﾃﾞｰﾀの内容をｸﾞﾘｯﾄﾞにｾｯﾄする
'----------------------------------------------------------------------------
Dim sts As Integer

Dim SUMI_QTY    As Long
Dim MI_QTY      As Long





    Grid_Set_Proc = True
    
    SHORDER.ReDim Min_Row, Row, Min_Col, Max_Col


    '注文日
    SHORDER(Row, colORDER_DT) = Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 7, 2)
    '注文№
    SHORDER(Row, colORDER_NO) = StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode)
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
    SHORDER(Row, colORDER_QTY) = Format(CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)), "#,##0")
    '注文残
    SHORDER(Row, colZAN_QTY) = Format(CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) - _
                                        CLng(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)), "#,##0")
    '在庫残
    If Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                            StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
        Exit Function
    
    End If
    SHORDER(Row, colZAIKO_QTY) = Format(SUMI_QTY + MI_QTY, "#,##0")
    '納期予定日
    If Trim(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode)) <> "" Then        '2007.09.06
        SHORDER(Row, colY_NOUKI_DT) = Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 7, 2)
    Else
        SHORDER(Row, colY_NOUKI_DT) = ""
    End If
    '回答納期
    If Trim(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode)) <> "" Then        '2007.09.06
        SHORDER(Row, colANS_NOUKI_DT) = Mid(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 7, 2)
    Else
        SHORDER(Row, colANS_NOUKI_DT) = ""
    End If
    
    
    
    'テスト収支 2016.01.19
    SHORDER(Row, colG_SYUSHI) = StrConv(P_SHORDER_REC.G_SYUSHI, vbUnicode)
    
    
    
    Grid_Set_Proc = False

End Function

Private Sub Input_Area_Set(Mode As Integer)
'----------------------------------------------------------------------------
'           入力エリアの切り替え
'----------------------------------------------------------------------------
                
                
    Select Case Mode
        Case 0  '納期--＞通常
                
            Text1(ptxG_SHIIRE_KBN).BackColor = G_INPUT_OK       '2008.01.10
            Text1(ptxG_SHIIRE_KBN).Locked = False
            Text1(ptxG_SHIIRE_KBN).TabStop = True
            
            Combo1(pcmbG_SHIIRE_KBN).BackColor = G_INPUT_OK     '2008.01.10
            Combo1(pcmbG_SHIIRE_KBN).Locked = False
            Combo1(pcmbG_SHIIRE_KBN).TabStop = True
                
                
            Text1(ptxY_NOUKI_DT).BackColor = G_INPUT_NG
            Text1(ptxY_NOUKI_DT).Locked = True
            Text1(ptxY_NOUKI_DT).TabStop = False


            If OSAKA_MODE Then      '2008.01.10
                
                Text1(ptxANS_NOUKI_DT).BackColor = G_INPUT_NG
                Text1(ptxANS_NOUKI_DT).Locked = True
                Text1(ptxANS_NOUKI_DT).TabStop = False

                Text1(ptxUSE_YM).BackColor = G_INPUT_NG
                Text1(ptxUSE_YM).Locked = True
                Text1(ptxUSE_YM).TabStop = False


            End If



            Text1(ptxUKEIRE_DT).BackColor = G_INPUT_OK
            Text1(ptxUKEIRE_DT).Locked = False
            Text1(ptxUKEIRE_DT).TabStop = True

            Text1(ptxKEIJYO_YM).BackColor = G_INPUT_OK
            Text1(ptxKEIJYO_YM).Locked = False
            Text1(ptxKEIJYO_YM).TabStop = True

            Text1(ptxKONKAI_UKEIRE_QTY).BackColor = G_INPUT_OK
            Text1(ptxKONKAI_UKEIRE_QTY).Locked = False
            Text1(ptxKONKAI_UKEIRE_QTY).TabStop = True

            Text1(ptxZAN_QTY).BackColor = G_INPUT_OK
            Text1(ptxZAN_QTY).Locked = False
            Text1(ptxZAN_QTY).TabStop = True

            Text1(ptxTANKA).BackColor = G_INPUT_OK      '2008.01.10
            Text1(ptxTANKA).Locked = False
            Text1(ptxTANKA).TabStop = True

            Text1(ptxKINGAKU).BackColor = G_INPUT_OK    '2008.01.10
            Text1(ptxKINGAKU).Locked = False
            Text1(ptxKINGAKU).TabStop = True

            Text1(ptxZEI_KIN).BackColor = G_INPUT_OK    '2008.01.10
            Text1(ptxZEI_KIN).Locked = False
            Text1(ptxZEI_KIN).TabStop = True

            Check1(chkZAIKO_F).Enabled = True           '2008.01.10
            Check1(chkZAIKO_F).TabStop = True




        Case 1  '通常--＞納期
                
            
            Text1(ptxG_SHIIRE_KBN).BackColor = G_INPUT_NG       '2008.01.10
            Text1(ptxG_SHIIRE_KBN).Locked = True
            Text1(ptxG_SHIIRE_KBN).TabStop = False
            
            Combo1(pcmbG_SHIIRE_KBN).BackColor = G_INPUT_NG     '2008.01.10
            Combo1(pcmbG_SHIIRE_KBN).Locked = True
            Combo1(pcmbG_SHIIRE_KBN).TabStop = False
            
            
            
            Text1(ptxY_NOUKI_DT).BackColor = G_INPUT_OK
            Text1(ptxY_NOUKI_DT).Locked = False
            Text1(ptxY_NOUKI_DT).TabStop = True


            If OSAKA_MODE Then      '2008.01.10
            
                Text1(ptxANS_NOUKI_DT).BackColor = G_INPUT_OK
                Text1(ptxANS_NOUKI_DT).Locked = False
                Text1(ptxANS_NOUKI_DT).TabStop = True
                
                Text1(ptxUSE_YM).BackColor = G_INPUT_OK
                Text1(ptxUSE_YM).Locked = False
                Text1(ptxUSE_YM).TabStop = True


            End If


            Text1(ptxUKEIRE_DT).BackColor = G_INPUT_NG
            Text1(ptxUKEIRE_DT).Locked = True
            Text1(ptxUKEIRE_DT).TabStop = False

            Text1(ptxKEIJYO_YM).BackColor = G_INPUT_NG
            Text1(ptxKEIJYO_YM).Locked = True
            Text1(ptxKEIJYO_YM).TabStop = False

            Text1(ptxKONKAI_UKEIRE_QTY).BackColor = G_INPUT_NG
            Text1(ptxKONKAI_UKEIRE_QTY).Locked = True
            Text1(ptxKONKAI_UKEIRE_QTY).TabStop = False

            Text1(ptxZAN_QTY).BackColor = G_INPUT_NG
            Text1(ptxZAN_QTY).Locked = True
            Text1(ptxZAN_QTY).TabStop = False

            Text1(ptxTANKA).BackColor = G_INPUT_NG      '2008.01.10
            Text1(ptxTANKA).Locked = True
            Text1(ptxTANKA).TabStop = False

            Text1(ptxKINGAKU).BackColor = G_INPUT_NG    '2008.01.10
            Text1(ptxKINGAKU).Locked = True
            Text1(ptxKINGAKU).TabStop = False

            Text1(ptxZEI_KIN).BackColor = G_INPUT_NG    '2008.01.10
            Text1(ptxZEI_KIN).Locked = True
            Text1(ptxZEI_KIN).TabStop = False

            Check1(chkZAIKO_F).Enabled = False          '2008.01.10
            Check1(chkZAIKO_F).TabStop = False


    End Select


End Sub

Private Function Code_Set_Proc(Index As Integer, KBN As String, Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   コードマスタをコンボにセットする。
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim Key_Len     As Integer
Dim OPTION1     As Integer
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
            wkOption = Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
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

Private Sub Input_Area_Proc(Mode As Integer)
'----------------------------------------------------------------------------
'                   入力可能領域の切り替え
'----------------------------------------------------------------------------
    
    
    Select Case Mode
        Case 0      'ノーマル
    
            Input_Mode = False
    
            '品番
            Text1(ptxHIN_GAI).BackColor = G_INPUT_NG
            Text1(ptxHIN_GAI).Locked = True
            Text1(ptxHIN_GAI).TabStop = False
    
            '担当者
            Text1(ptxTANTO_CODE).BackColor = G_INPUT_NG
            Text1(ptxTANTO_CODE).Locked = True
            Text1(ptxTANTO_CODE).TabStop = False
            '注文先
            Text1(ptxORDER_CODE).BackColor = G_INPUT_NG
            Text1(ptxORDER_CODE).Locked = True
            Text1(ptxORDER_CODE).TabStop = False
            
            Combo1(pcmbORDER).BackColor = G_INPUT_NG
            Combo1(pcmbORDER).Locked = True
            Combo1(pcmbORDER).TabStop = False
            '単価
            Text1(ptxTANKA).BackColor = G_INPUT_OK
            Text1(ptxTANKA).Locked = False
            Text1(ptxTANKA).TabStop = True
            '注文残
            Text1(ptxZAN_QTY).BackColor = G_INPUT_OK
            Text1(ptxZAN_QTY).Locked = False
            Text1(ptxZAN_QTY).TabStop = True
    
    
    

    
        Case 1      '注文なし時
    
            Input_Mode = True
    
    
            '品番
            Text1(ptxHIN_GAI).BackColor = G_INPUT_OK
            Text1(ptxHIN_GAI).Locked = False
            Text1(ptxHIN_GAI).TabStop = True
                            
            '担当者
            Text1(ptxTANTO_CODE).BackColor = G_INPUT_OK
            Text1(ptxTANTO_CODE).Locked = False
            Text1(ptxTANTO_CODE).TabStop = True
            '注文先
            Text1(ptxORDER_CODE).BackColor = G_INPUT_OK
            Text1(ptxORDER_CODE).Locked = False
            Text1(ptxORDER_CODE).TabStop = True
            
            Combo1(pcmbORDER).BackColor = G_INPUT_OK
            Combo1(pcmbORDER).Locked = False
            Combo1(pcmbORDER).TabStop = True
    
            '単価
            Text1(ptxTANKA).BackColor = G_INPUT_OK
            Text1(ptxTANKA).Locked = False
            Text1(ptxTANKA).TabStop = True
            '注文残
            Text1(ptxZAN_QTY).BackColor = G_INPUT_NG
            Text1(ptxZAN_QTY).Locked = True
            Text1(ptxZAN_QTY).TabStop = False
    
    End Select

End Sub

Private Function Hin_Item_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   品目マスタのﾁｪｯｸ＆内容表示
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long
Dim i           As Integer

    Hin_Item_Disp_Proc = True
    
    
    If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI And _
        StrConv(ITEMREC.NAIGAI, vbUnicode) = NAIGAI_NAI And _
        Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) = Trim(Text1(ptxHIN_GAI).Text) Then
    
        Hin_Item_Disp_Proc = False
        Exit Function
    End If
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
        
            Text1(ptxHIN_NAME).Text = ""
            Text1(ptxZAIKO_QTY).Text = ""

' 2012.12.28 Upd >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            Text1(ptxG_SYUSHI).Text = ""        '収支単位
            Text1(ptxSYUSHI_NM).Text = ""       '収支名
' 2012.12.28 Upd <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

            Hin_Item_Disp_Proc = BtErrKeyNotFound
            Exit Function
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    
    End Select
    Text1(ptxHIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)

' 2012.12.28 Upd >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Text1(ptxG_SYUSHI).Text = Trim(StrConv(ITEMREC.G_SYUSHI, vbUnicode))        '収支単位

    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN03_CD)                           '収支名
    Call UniCode_Conv(K0_P_CODE.C_Code, Text1(ptxG_SYUSHI).Text)
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxSYUSHI_NM).Text = StrConv(P_CODEREC.C_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text1(ptxSYUSHI_NM).Text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "商品ﾗﾍﾞﾙｺﾝﾄﾛｰﾙ ﾌｧｲﾙ")
            Exit Function
    End Select
' 2012.12.28 Upd <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    If Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                            StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
        Exit Function
    
    End If
    Text1(ptxZAIKO_QTY).Text = Format(SUMI_QTY + MI_QTY, "#0")
        
        
    Text1(ptxG_SHIIRE_KBN).Text = StrConv(ITEMREC.G_SHIIRE_KBN, vbUnicode)   '仕入区分
    Combo1(pcmbG_SHIIRE_KBN).ListIndex = -1
    For i = 0 To Combo1(pcmbG_SHIIRE_KBN).ListCount - 1
    
        If Text1(ptxG_SHIIRE_KBN).Text = Trim(Left(Right(Combo1(pcmbG_SHIIRE_KBN).List(i), 3), 2)) Then
            Combo1(pcmbG_SHIIRE_KBN).ListIndex = i
            Exit For
        End If
    
    Next i
    
    
    If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_ON Then
        Check1(chkZAIKO_F).Value = vbChecked
    Else
        Check1(chkZAIKO_F).Value = vbUnchecked
    End If
        
    Hin_Item_Disp_Proc = False
End Function

Private Function POS_NYUKA_Update_Proc(SOKO As String, Retu As String, Ren As String, Dan As String) As Integer
'----------------------------------------------------------------------------
'                   POS用在庫＆入荷予定更新
'           POSｼｽﾃﾑ無しは、標準棚番に在庫計上する2006.04.24
'----------------------------------------------------------------------------
                                            
Dim sts         As Integer
Dim com         As Integer


Dim DEN_NO      As String * 6
Dim ID_NO       As String * 9
Dim ans         As Integer
                                            
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long

Dim WK_Qty      As Long     '前借残ワーク
Dim WK_E_QTY    As Long     '先行出荷数ワーク
                                            
Dim MAEGARI_QTY As Long
                                            
Dim SOUSAI_QTY  As Long
                                            
Dim Upd_QTY     As Long     '2007.05.03
                                            
Dim TO_SOKO     As String * 2
Dim TO_RETU     As String * 2
Dim TO_REN      As String * 2
Dim TO_DAN      As String * 2
                                            
    POS_NYUKA_Update_Proc = True
                                        
'    Call Input_Lock

''    If CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text) <= 0 Then
''        POS_NYUKA_Update_Proc = False
''        Exit Function
''    End If
    

    If Trim(SOKO) = "" Then
        TO_SOKO = KASO_NYUKA
        TO_RETU = "01"
        TO_REN = "01"
        TO_DAN = "01"
    Else
        'ＰＯＳｼｽﾃﾑ無しは標準棚番へ
        TO_SOKO = SOKO
        TO_RETU = Retu
        TO_REN = Ren
        TO_DAN = Dan
    
    
        Call UniCode_Conv(K0_TANA.SOKO_NO, TO_SOKO)
        Call UniCode_Conv(K0_TANA.Retu, TO_RETU)
        Call UniCode_Conv(K0_TANA.Ren, TO_REN)
        Call UniCode_Conv(K0_TANA.Dan, TO_DAN)

    
        sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                '未登録は入荷仮想へ
                TO_SOKO = KASO_NYUKA
                TO_RETU = "01"
                TO_REN = "01"
                TO_DAN = "01"
                    
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "棚マスタ")
                Exit Function
        
        End Select
    
    
    End If








    WK_E_QTY = 0
                                            
    SUMI_QTY = 0
                            '資材品は全て未商品として扱う
    MI_QTY = CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode))
                
                
    '資材入荷ﾁｪｯｸﾃﾞｰﾀ(前借ﾃﾞｰﾀ)更新
    Call UniCode_Conv(K0_P_NYU.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_P_NYU.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_P_NYU.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
    Call UniCode_Conv(K0_P_NYU.NYUKA_DT, "")
    
    com = BtOpGetGreater
    
    Do
        DoEvents
                
        Do
            sts = BTRV(com + BtSNoWait, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
            Select Case sts
                Case BtNoErr
                    
                    If StrConv(P_NYUREC.JGYOBU, vbUnicode) <> SHIZAI Or _
                        StrConv(P_NYUREC.NAIGAI, vbUnicode) <> NAIGAI_NAI Or _
                        StrConv(P_NYUREC.HIN_GAI, vbUnicode) <> StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode) Then
                        
                        sts = BTRV(BtOpUnlock, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "資材前借ﾃﾞｰﾀ")
                            Exit Function
                        End If
                        sts = BtErrEOF
                        Exit Do
                    End If
                    If IsNumeric(StrConv(P_NYUREC.SOUSAI_QTY, vbUnicode)) Then
                        SOUSAI_QTY = CLng(StrConv(P_NYUREC.SOUSAI_QTY, vbUnicode))
                    Else
                        SOUSAI_QTY = 0
                    End If
                    MAEGARI_QTY = CLng(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode)) - SOUSAI_QTY
                    If MAEGARI_QTY > MI_QTY Then
                        SOUSAI_QTY = SOUSAI_QTY + MI_QTY        '2007.05.03
                        MI_QTY = MAEGARI_QTY - MI_QTY
                        Call UniCode_Conv(P_NYUREC.SOUSAI_DT, Format(Now, "YYYYMMDD"))
                '        Call UniCode_Conv(P_NYUREC.SOUSAI_QTY, Format(MI_QTY, "00000000"))
                        Call UniCode_Conv(P_NYUREC.SOUSAI_QTY, Format(SOUSAI_QTY, "00000000"))
                
                        Do
                        
                            sts = BTRV(BtOpUpdate, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("他端末でデータ使用中です。<P_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                    If ans = vbCancel Then
                                        Exit Function
                                    End If
                                Case Else
                                    Call File_Error(sts, BtOpUpdate, "資材前借ﾃﾞｰﾀ")
                                    Exit Function
                            End Select
                        
                        Loop
                        WK_E_QTY = CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode))  '先行処理分
                        
                        
                        sts = BtErrEOF      '2007.08.21
                        
                        Exit Do
                    Else
                        Do
                            sts = BTRV(BtOpDelete, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("他端末でデータ使用中です。<P_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                    If ans = vbCancel Then
                                        Exit Function
                                    End If
                                Case Else
                                    Call File_Error(sts, BtOpDelete, "資材前借ﾃﾞｰﾀ")
                                    Exit Function
                            End Select
                        Loop
                        
                        
                        MI_QTY = MI_QTY - MAEGARI_QTY
                        WK_E_QTY = WK_E_QTY + MAEGARI_QTY
                    
                        If MI_QTY = 0 Then
                            sts = BtErrEOF
                            Exit Do
                        End If
                    
                    End If
            
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                   End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "資材前借ﾃﾞｰﾀ")
                    Exit Function
            End Select
        
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
        
        com = BtOpGetNext
    
    Loop
                                            '入荷予定編集
    Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)            '完了区分
    Call UniCode_Conv(Y_NYUREC.DT_SYU, "R")                     'データ種別
    Call UniCode_Conv(Y_NYUREC.JGYOBU, SHIZAI)                  '事業部
    Call UniCode_Conv(Y_NYUREC.NAIGAI, NAIGAI_NAI)              '国内外
    Call UniCode_Conv(Y_NYUREC.JGYOBA, "")                      '事業場
    Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")                    'データ区分
    Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")                    '取引区分
                                                                'ＩＤ№
    sts = Den_No_Set_Proc(11, SHIZAI, ID_NO)
    If sts Then
        Exit Function
    End If
    
    Call UniCode_Conv(Y_NYUREC.ID_NO, ID_NO)
    Call UniCode_Conv(Y_NYUREC.TEXT_NO, ID_NO)
                                                                '品目番号
    Call UniCode_Conv(Y_NYUREC.HIN_NO, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
                                                                
                                                                '伝票№
    sts = Den_No_Set_Proc(10, SHIZAI, DEN_NO)
    If sts Then
        Exit Function
    End If
    Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
                                                                '予定数量
    Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode)), "0000000"))
    Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")                   '出庫先
    Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")                 '出庫収支
                                                                '出庫日付
    Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, Format(Now, "YYYYMMDD"))
    Call UniCode_Conv(Y_NYUREC.TANKA, "")                       '単価
    Call UniCode_Conv(Y_NYUREC.ODER_NO, "")                     'オーダー番号
    Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")                     'アイテム番号
    Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")                   'オーダー略号
    Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")                 '個装形態
                                                                '出荷日
    Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode))
                                                                '棚番１
    Call UniCode_Conv(Y_NYUREC.TANABAN1, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                            StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                            StrConv(ITEMREC.ST_DAN, vbUnicode))
        
    Call UniCode_Conv(Y_NYUREC.TANABAN2, "")                    '棚番２
    Call UniCode_Conv(Y_NYUREC.TANABAN3, "")                    '棚番３
    Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")                   '出庫先名称
    Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")                     '注文区分
    Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")                '注文区分名称
    Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")                     '原産国１
    Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")                     '原産国２
    Call UniCode_Conv(Y_NYUREC.BIKOU2, "")                      '備考２
    Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")                     '販売区分
    Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")                   '直送区分
    Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")                  'ﾕﾆｯﾄ修理ID-NO
    Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")               '在庫引当順序
    Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")              '合梱管理番号
    Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")                  '受注残数量
    Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")                  '供給区分
    Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")                '商品化納入先収支
    Call UniCode_Conv(Y_NYUREC.BIKOU1, "")                      '備考１
    Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")                   '帳端区分
    Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")                  '受注品目番号
                                                                '品名
    Call UniCode_Conv(Y_NYUREC.HIN_NAME, StrConv(ITEMREC.HIN_NAME, vbUnicode))
    Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")              '品番変更区分
    Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")             'モジュール交換区分
    Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")                 '残在庫まとめ在庫収支コード
    Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")                   '指定納期
    Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")            'サービス会社管理番号
    Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")                   '機種品目コード
    Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")             '環境規格部品区分
    Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD")) '完了日付
                                                                '先行入荷数
    Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, Format(WK_E_QTY, "00000000"))
    
'>>>>>>>>>>>>>>>>>>>>>> 2016.06.29
    Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, "")                  '予算単位（元）
    Call UniCode_Conv(Y_NYUREC.YOSAN_TO, "")                    '予算単位（先）
    Call UniCode_Conv(Y_NYUREC.HTANABAN, "")                    '標準棚番
    Call UniCode_Conv(Y_NYUREC.HIN_NAI, "")                     '品番（内部）
    Call UniCode_Conv(Y_NYUREC.H_SOKO, "")                      'ﾎｽﾄ倉庫 2006.10.17
            
    Call UniCode_Conv(Y_NYUREC.NYU_LIST_OUT, "")                '入庫予定出力ﾌﾗｸﾞ 2007.06.12    現在未使用 0:データ出力対象 9:出力済(もしくは出力対象外)
    
    Call UniCode_Conv(Y_NYUREC.GENSANKOKU, "")                  '原産国名
    Call UniCode_Conv(Y_NYUREC.GEN_GENSANKOKU, "")              '現物表示原産国名
    Call UniCode_Conv(Y_NYUREC.SHIIRE_WORK_CENTER, "")          '資材仕入先ﾜｰｸｾﾝﾀｰ
    Call UniCode_Conv(Y_NYUREC.KANKYO_KBN, "")                  '環境種類区分
    Call UniCode_Conv(Y_NYUREC.KANKYO_KBN_ST, "")               '環境種類区分適用開始
    Call UniCode_Conv(Y_NYUREC.KANKYO_KBN_SURYO, "")            '環境種類区分数量
    Call UniCode_Conv(Y_NYUREC.ID_NO2, "")                      'ID_NO
    Call UniCode_Conv(Y_NYUREC.AITESAKI_CODE, "")               '相手先ｺｰﾄﾞ
    Call UniCode_Conv(Y_NYUREC.JYUCHU_YMD, "")                  '受注年月日
    Call UniCode_Conv(Y_NYUREC.SHITEI_NOUKI_YMD, "")            '指定納期年月日
    Call UniCode_Conv(Y_NYUREC.LIST_OUT_END_F, "")              '入庫関連ﾘｽﾄ出力F    0:複数原産国部品入庫管理ﾘｽﾄまたは入庫／棚番ﾁｪｯｸﾘｽﾄが未処理
                                                                '9:複数原産国部品入庫管理ﾘｽﾄかつ入庫／棚番ﾁｪｯｸﾘｽﾄが処理済
    Call UniCode_Conv(Y_NYUREC.LIST_NYU_KANRI_F, "")            '入庫管理ﾘｽﾄ出力F　　「複数原産国部品入庫管理ﾘｽﾄ用」 0:印刷対象(未印刷) 8:印刷対象外　9:印刷済(0→9)
    Call UniCode_Conv(Y_NYUREC.LIST_NYU_CHECK_F, "")            '入庫ﾁｪｯｸﾘｽﾄ出力F    「入庫／棚番ﾁｪｯｸﾘｽﾄ用」　0:未印刷 9:印刷済
    Call UniCode_Conv(Y_NYUREC.NYUKO_TANABAN, "")               '入庫棚番
    Call UniCode_Conv(Y_NYUREC.MAEGARI_SURYO, "")               '前借相殺数
    
    Call UniCode_Conv(Y_NYUREC.INS_TANTO, Text1(ptxTANTO_CODE).Text)        '追加　担当者
    Call UniCode_Conv(Y_NYUREC.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS")) '追加　日時

    Call UniCode_Conv(Y_NYUREC.UPD_TANTO, "")                   '更新　担当者
    Call UniCode_Conv(Y_NYUREC.UPD_DATETIME, "")                '更新　日時
    Call UniCode_Conv(Y_NYUREC.MOTO_PROG_ID, "")                '発生元プログラム
    Call UniCode_Conv(Y_NYUREC.MOTO_TEXT_NO, "")                '元テキスト№
    
    Call UniCode_Conv(Y_NYUREC.JITU_SURYO, "")                  '実績数量
'>>>>>>>>>>>>>>>>>>>>>> 2016.06.29
    
    Call UniCode_Conv(Y_NYUREC.FILLER, "")
    
    Do
        sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
        
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case BtErrDuplicates
                                        '自動発番データ重複は再発行
                sts = Den_No_Set_Proc(11, SHIZAI, ID_NO)
                If sts Then
                    Exit Function
                End If

                Call UniCode_Conv(Y_NYUREC.ID_NO, ID_NO)
                Call UniCode_Conv(Y_NYUREC.TEXT_NO, ID_NO)
                
            Case Else
                Call File_Error(sts, BtOpInsert, "入荷予定データ")
                Exit Function
        End Select
    Loop
                            
'    sts = Nyuko_Update_Proc(SHIZAI, _
'                            NAIGAI_NAI, _
'                            StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode), _
'                            StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode), _
'                            (TO_SOKO & TO_RETU & TO_REN & TO_DAN), _
'                            P_YOIN_TU_NYUKA, _
'                            SUMI_QTY, _
'                            CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text), _
'                            WS_NO, _
'                            StrConv(P_SHORDER_REC.TANTO_CODE, vbUnicode), , _
'                            MEMO_TEXT, _
'                            StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode), _
'                            StrConv(P_SHORDER_REC.TANKA, vbUnicode), _
'                            StrConv(P_SHUKEIRE_REC.KEIJYO_YM, vbUnicode))
                            
                            
                            
                            
                            
                            
    sts = Nyuko_Update_Proc(SHIZAI, _
                            NAIGAI_NAI, _
                            StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode), _
                            StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode), _
                            (TO_SOKO & TO_RETU & TO_REN & TO_DAN), _
                            P_YOIN_TU_NYUKA, _
                            SUMI_QTY, _
                            CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text), _
                            WS_NO, _
                            StrConv(P_SHORDER_REC.TANTO_CODE, vbUnicode), , _
                            MEMO_TEXT, _
                            StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode), _
                            StrConv(P_SHUKEIRE_REC.UKEIRE_TANKA, vbUnicode), _
                            StrConv(P_SHUKEIRE_REC.KEIJYO_YM, vbUnicode))
                            
                            
                            
                            
                            
                            
    If sts Then
        Exit Function
    End If


    '前借り数で在庫データ更新（－）
    If WK_E_QTY <> 0 Then
    '在庫データLOCK
        If Zaiko_Lock_Proc((TO_SOKO & TO_RETU & TO_REN & TO_DAN), _
                            SHIZAI, _
                            NAIGAI_NAI, _
                            StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode), _
                            WS_NO) Then
            Exit Function

        End If

        MI_QTY = WK_E_QTY
        SUMI_QTY = 0

        If Syuko_Update_Proc(SHIZAI, _
                            NAIGAI_NAI, _
                            StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode), _
                            StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode), _
                            (TO_SOKO & TO_RETU & TO_REN & TO_DAN), _
                            P_YOIN_MAE_SOUSAI, _
                            0, WK_E_QTY, 0, _
                            WS_NO, WS_NO) Then
            Exit Function

        End If






    End If



    POS_NYUKA_Update_Proc = False
End Function

Private Sub Text1_LostFocus(Index As Integer)

Dim ZEI         As Long
Dim wkKINGAKU   As Long

Dim sts         As Integer

    Select Case Index
    
    
        
        
        
        Case ptxHIN_GAI                                                             '2013.10.08
            Text1(Index).Text = StrConv(RTrim(Text1(Index).Text), vbUpperCase)      '2013.10.08
                
                
'            If Error_Check_Proc(Index) Then                     'エラーチェック 2016.06.16     2017.05.06
'                Exit Sub                                        '2016.06.16                    2017.05.06
'            End If                                              '2016.06.16                    2017.05.06
                
    
        Case ptxTANKA
        
        
            If wkTANKA = Trim(ptxTANKA) Then
                Exit Sub
            End If
                    
            If IsNumeric(Text1(ptxTANKA).Text) And IsNumeric(Text1(ptxKONKAI_UKEIRE_QTY).Text) Then
                        
                '2009.11.02
'                Text1(ptxKINGAKU).Text = Format(CDbl(Text1(ptxTANKA).Text) * CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text), "#,##0")
                    
                    
                    
                    
                Select Case StrConv(P_KANRIREC.SHI_MARUME, vbUnicode)
                    Case "0"    '切捨て
                        Text1(ptxKINGAKU).Text = Format(ToRoundDown(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                        CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")
                    
        
                    Case "5"    '四捨五入
                    
                        Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                        CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")
                    
                    
                    
                    
                    Case "9"    '切り上げ
                
                
                        Text1(ptxKINGAKU).Text = Format(ToRoundUp(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                        CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")
            
                
                
                    Case Else    '四捨五入
                    
                        Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                        CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")
                
                
                End Select
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                '2007.11.01 他ｾﾝﾀｰ、時給は消費税計算しない
                Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxORDER_CODE).Text)
                sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        MsgBox "受払先マスタ未登録です。"
                        Text1(ptxORDER_CODE).SetFocus
                        Unload Me
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
                        Unload Me
                End Select
                    
                If StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) <> P_TORI_ANOTHER And StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) <> P_TORI_JIKYU Then
                    
                    
                    
                    If CLng(Text1(ptxKINGAKU).Text) >= 0 Then
                        If Format(Text1(ptxUKEIRE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                            ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                                    CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
                        Else
                            ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                    CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))
                        End If
                    Else
                        
                        wkKINGAKU = CLng(Text1(ptxKINGAKU).Text) * -1
                        
                        If Format(Text1(ptxUKEIRE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                            ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                                    CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
                        Else
                            ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                    CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))
                        End If
                        ZEI = ZEI * -1
                    End If
    
                    Text1(ptxZEI_KIN).Text = Format(ZEI, "#,##0")
                Else
                    Text1(ptxZEI_KIN).Text = "0"
        
                End If
            End If
        Case ptxKONKAI_UKEIRE_QTY
    
    
            If wkUKEIRE_QTY = Trim(ptxKONKAI_UKEIRE_QTY) Then
                Exit Sub
            End If
                    
            If IsNumeric(Text1(ptxTANKA).Text) And IsNumeric(Text1(ptxKONKAI_UKEIRE_QTY).Text) Then
    
                '2009.11.02
'                Text1(ptxKINGAKU).Text = Format(CDbl(Text1(ptxTANKA).Text) * CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text), "#,##0")
                Select Case StrConv(P_KANRIREC.SHI_MARUME, vbUnicode)
                    Case "0"    '切捨て
                        Text1(ptxKINGAKU).Text = Format(ToRoundDown(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                        CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")
                    
        
                    Case "5"    '四捨五入
                    
                        Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                        CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")
                    
                    
                    
                    
                    Case "9"    '切り上げ
                
                
                        Text1(ptxKINGAKU).Text = Format(ToRoundUp(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                        CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")
            
                
                
                    Case Else    '四捨五入
                    
                        Text1(ptxKINGAKU).Text = Format(ToHalfAdjust(CCur(CCur(Text1(ptxTANKA).Text) * _
                                                        CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)), 0), "#,##0")
                
                
                End Select
                    
                    
                    
                '2007.11.01 他ｾﾝﾀｰ、時給は消費税計算しない
                Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxORDER_CODE).Text)
                sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        MsgBox "受払先マスタ未登録です。"
                        Text1(ptxORDER_CODE).SetFocus
                    
                        Unload Me
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
                        Unload Me
                End Select
                    
                If StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) <> P_TORI_ANOTHER And StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) <> P_TORI_JIKYU Then
                    
                    
                    If CLng(Text1(ptxKINGAKU).Text) >= 0 Then
                        If Format(Text1(ptxUKEIRE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                            ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                                    CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
                        Else
                            ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                    CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))
                        End If
                    Else
                        
                        wkKINGAKU = CLng(Text1(ptxKINGAKU).Text) * -1
                        
                        If Format(Text1(ptxUKEIRE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                            ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                                    CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
                        Else
                            ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                    CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))
                        End If
                        ZEI = ZEI * -1
                    End If
    
                    Text1(ptxZEI_KIN).Text = Format(ZEI, "#,##0")
    
                Else
                    Text1(ptxZEI_KIN).Text = "0"
                End If
    
            End If
    End Select



End Sub

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


