VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form F1020251 
   BackColor       =   &H00FFFFFF&
   Caption         =   "入庫予定表作成処理 "
   ClientHeight    =   11490
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   15375
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
   ScaleHeight     =   11490
   ScaleWidth      =   15375
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   2
      Left            =   2865
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   120
      Width           =   330
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   1
      Left            =   2130
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   120
      Width           =   330
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   0
      Left            =   1185
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   120
      Width           =   645
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   12705
      ScaleHeight     =   315
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   10080
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton Command 
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
      Left            =   10320
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   10200
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   10200
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   10200
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "EXCEL"
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   10200
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "表 示"
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   10200
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   10200
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   10200
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   4
      Left            =   3960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   10200
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   10200
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   10200
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   10200
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "取 込"
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   10200
      Width           =   855
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   9255
      Left            =   480
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   720
      Width           =   14070
      _ExtentX        =   24818
      _ExtentY        =   16325
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "伝票№"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "ﾒｰｶｰCODE"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "ﾒｰｶｰ名"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "品番"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "予定数量"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "入庫数量"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "検収担当者"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "検収担当者名"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "注文№"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1667"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1561"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1667"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1561"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=512"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3889"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3784"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=512"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=3784"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3678"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2355"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2249"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2355"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2249"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=2275"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2170"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=513"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=3493"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=3387"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=2037"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=1931"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(9).Width=661"
      Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=556"
      Splits(0)._ColumnProps(49)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=10.5,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=10.5,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=14,.parent=53,.alignment=3"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=11,.parent=54,.alignment=2"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=12,.parent=55"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=13,.parent=57"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=110,.parent=53,.alignment=0"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=107,.parent=54,.alignment=2"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=108,.parent=55"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=109,.parent=57"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=18,.parent=53,.alignment=0"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=15,.parent=54,.alignment=2"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=16,.parent=55"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=17,.parent=57"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=48,.parent=53"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=45,.parent=54,.alignment=2"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=46,.parent=55"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=47,.parent=57"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=114,.parent=53,.alignment=1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=111,.parent=54,.alignment=2"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=112,.parent=55"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=113,.parent=57"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=66,.parent=53,.alignment=1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=54,.alignment=2"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=55"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=57"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=102,.parent=53,.alignment=2"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=19,.parent=54,.alignment=2"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=20,.parent=55"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=101,.parent=57"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=70,.parent=53"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=54,.alignment=2"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=55"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=57"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=78,.parent=53"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=75,.parent=54,.alignment=2"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=76,.parent=55"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=77,.parent=57"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=24,.parent=53"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=21,.parent=54"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=22,.parent=55"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=23,.parent=57"
      _StyleDefs(76)  =   "Named:id=29:Normal"
      _StyleDefs(77)  =   ":id=29,.parent=0"
      _StyleDefs(78)  =   "Named:id=30:Heading"
      _StyleDefs(79)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(80)  =   ":id=30,.wraptext=-1"
      _StyleDefs(81)  =   "Named:id=31:Footing"
      _StyleDefs(82)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(83)  =   "Named:id=32:Selected"
      _StyleDefs(84)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(85)  =   "Named:id=33:Caption"
      _StyleDefs(86)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(87)  =   "Named:id=34:HighlightRow"
      _StyleDefs(88)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(89)  =   "Named:id=35:EvenRow"
      _StyleDefs(90)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(91)  =   "Named:id=36:OddRow"
      _StyleDefs(92)  =   ":id=36,.parent=29"
      _StyleDefs(93)  =   "Named:id=89:RecordSelector"
      _StyleDefs(94)  =   ":id=89,.parent=30"
      _StyleDefs(95)  =   "Named:id=92:FilterBar"
      _StyleDefs(96)  =   ":id=92,.parent=29"
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "日"
      Height          =   255
      Index           =   3
      Left            =   3285
      TabIndex        =   21
      Top             =   240
      Width           =   330
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "月"
      Height          =   255
      Index           =   2
      Left            =   2550
      TabIndex        =   19
      Top             =   240
      Width           =   330
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "年"
      Height          =   255
      Index           =   1
      Left            =   1815
      TabIndex        =   17
      Top             =   240
      Width           =   330
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "入庫日"
      Height          =   375
      Index           =   0
      Left            =   450
      TabIndex        =   16
      Top             =   240
      Width           =   855
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
      TabIndex        =   12
      Top             =   10800
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
Attribute VB_Name = "F1020251"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Const In_Mode% = 1                  '入荷処理
Private Const Out_Mode% = 2                 '出荷処理


Private DEF_SOKO_NO         As String * 2   '倉庫№
Private GOODS_KBN           As String * 1   '商品化 要／不要

Private Type SHIMUKE_TBL
    SHIMUKE_CODE            As String * 2   '仕向け先
    JGYOBU                  As String * 1   '事業部
    NAIGAI                  As String * 1   '国内外
End Type

Private SHIMUKE_T()         As SHIMUKE_TBL

Private SHIMUKE_Flg         As Boolean


Private Const ptxNYUKO_YY% = 0
Private Const ptxNYUKO_MM% = 1
Private Const ptxNYUKO_DD% = 2


Private NYUKA               As New XArrayDB
Private Const Min_Row% = 1                  '最小行数
Private Max_Row             As Long         'グリッド最大表示件数
Private Const Min_Col% = 0                  '最小列数
Private Const Max_Col% = 9                  '最大列数
    
Private Const colDEN_NO% = 0                '伝票№
Private Const colMAKER_CODE% = 1            'ﾒｰｶｰｺｰﾄﾞ
Private Const colMAKER_NAME% = 2            'ﾒｰｶｰ名称
Private Const colHIN_NO% = 3                '品番
Private Const colY_SURYO% = 4               '予定数量
Private Const colJ_SURYO% = 5               '入庫数量
Private Const colTANTO_CODE% = 6            '検収担当者ｺｰﾄﾞ
Private Const colTANTO_NAME% = 7            '検収担当者名
Private Const colORDER_NO% = 8              '注文№
Private Const colKENPIN_F% = 9              '検品ﾌﾗｸﾞ

Private Sort_Tbl(Min_Col To Max_Col) _
                    As Integer              'ｿｰﾄの制御 0:昇順 1:降順

Private DATA_FLG    As Boolean

Dim Excel_Template      As String       '棚札 ﾃﾝﾌﾟﾚｰﾄ(ﾌﾙ･ﾊﾟｽ)
Dim Excel_PutPath       As String       '棚札 書き込み先ﾊﾟｽ
Dim Excel_Soko_Name     As String       'EXCEL出力用倉庫名称

Dim EXCEL_DIR           As String

'Dim ExcelApp            As Excel.Application
'Dim Excelbook           As Excel.Workbook
'Dim ExcelWorkSheet      As Excel.Worksheet

'Private Const LAST_UPDATE_DAY$ = "([F102025] 2011.04.11 09:00 ｽﾃｰﾀｽﾊﾞｰ&LOG出力 追加)"


Dim F102025_LOG         As String       '2017.01.05



'>>>>>>>>>>>>>>>>>>>    2017.12.28
Private Const EX_DEN_NO% = 1
Private Const EX_MAKER_CODE% = 2
Private Const EX_HIN_NO% = 3
Private Const EX_Y_SURYO% = 4
Private Const EX_ORDER_NO_1% = 5

Private Const EX_MAKER_NAME% = 2
Private Const EX_ORDER_NO_2% = 5



'Private Const LAST_UPDATE_DAY$ = "([F102025] 2017.12.28 15:15)"
Private Const LAST_UPDATE_DAY$ = "([F102025] 2018.03.29 16:15)"



Private Sub Command_Click(Index As Integer)

Dim ans As Integer
Dim c   As String * 128


    Select Case Index
        
        
        
        Case 0      'データ取り込み
        
        
            ans = MsgBox("データ取り込みを行います？", vbYesNo, "確認入力")
            If ans = vbYes Then
                If Data_Get_Proc() Then
                    Unload Me
                End If
            
                If List_Disp_Proc() Then
                    Unload Me
                End If
            
            
If Trim(F102025_LOG) <> "" Then                             '2018.03.29
    Call LOG_OUT(F102025_LOG, "F102025 データ取り込み")     '2018.03.29
End If                                                      '2018.03.29
            
            
            
            End If
        
            Command(0).SetFocus
        
        Case 3      '印刷
        
        
        
        
            If EXCEL_Put_Proc(1) Then
                Unload Me
            End If
        
            Command(3).SetFocus
        
        Case 4      '更新
        
            ans = MsgBox("データ更新を行います？", vbYesNo, "確認入力")
            If ans = vbYes Then
'                If Update_Proc() Then              '2017.01.27
                If NEW_Update_Proc() Then           '2017.01.27
                    Unload Me
                End If
                
                If List_Disp_Proc() Then
                    Unload Me
                End If
            
            
If Trim(F102025_LOG) <> "" Then                             '2018.03.29
    Call LOG_OUT(F102025_LOG, "F102025 データ更新")         '2018.03.29
End If                                                      '2018.03.29
            
            
            
            End If
            Command(4).SetFocus
            
        
        Case 7      '表示
        
        
        
        
            If List_Disp_Proc() Then
                Unload Me
            End If
        
        
If Trim(F102025_LOG) <> "" Then                             '2018.03.29
    Call LOG_OUT(F102025_LOG, "F102025 表示")               '2018.03.29
End If                                                      '2018.03.29
        
        
        
        
            Command(7).SetFocus
  
        Case 8      'EXCEL出力
        
        
        
        
            If EXCEL_Put_Proc(0) Then
                Unload Me
            End If
        
'            Command(8).SetFocus        DEL 2013.03.19
  
If Trim(F102025_LOG) <> "" Then                             '2018.03.29
    Call LOG_OUT(F102025_LOG, "F102025 EXCEL出力")          '2018.03.29
End If                                                      '2018.03.29

        
        Case 11                             '終了
            Unload Me
        Case Else
            Beep
    End Select
    
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

Dim i               As Integer
Dim c               As String * 128
Dim sts             As Integer
Dim com             As Integer


    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If



'---------------------------------------------------    2011.04.11
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[POSシステム]入庫予定表作成", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
'---------------------------------------------------    2011.04.11





    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                    '最大表示件数の獲得
    If GetIni(App.EXEName, "LISTMAX", App.EXEName, c) Then
        Max_Row = 99999
    Else
        Max_Row = CLng(RTrim(c))
    End If
                                '事業部取り込み
    If JGYOB_TB_Set(1) Then
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
            F1020251.Caption = "入庫予定表作成処理（" + RTrim(JGYOBU_T(i).NAME) + ")" & LAST_UPDATE_DAY
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)


    '出力倉庫名
    If GetIni(App.EXEName, "DEF_SOKO_NO", App.EXEName, c) Then
        Beep
        MsgBox "倉庫№の獲得に失敗しました。処理を中止して下さい。"
        End
    Else
        DEF_SOKO_NO = Trim(c)
    End If

    'EXCELテンプレート獲得
    If GetIni(App.EXEName, "EXCEL_TEMPLATE", App.EXEName, c) Then
        Beep
        MsgBox "ﾃﾝﾌﾟﾚｰﾄ(ﾌﾙ･ﾊﾟｽ)の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    Excel_Template = Trim(c)
    '書き込み先ﾊﾟｽ
    If GetIni(App.EXEName, "EXCEL_OUTPUT", App.EXEName, c) Then
        Beep
        MsgBox "書き込み先ﾊﾟｽの獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    Excel_PutPath = Trim(c)

    'EXCELフォルダ  DEL 2013.03.19
'    If GetIni(App.EXEName, "EXCEL_DIR", App.EXEName, c) Then
'        Beep
'        MsgBox "EXCELﾌｫﾙﾀﾞの獲得に失敗しました。処理を中止して下さい。"
'        End
'    End If
'    EXCEL_DIR = Trim(c)


    '出力倉庫名
    If GetIni(App.EXEName, "EXCEL_SOKO_NAME", App.EXEName, c) Then
        Excel_Soko_Name = ""
    Else
        Excel_Soko_Name = Trim(c)
    End If

    '専用ログ
    If GetIni(App.EXEName, "F102025_LOG", App.EXEName, c) Then
        F102025_LOG = ""
    Else
        F102025_LOG = Trim(c)
    End If




                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '構成マスタＯＰＥＮ
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '入荷予定 ＯＰＥＮ
    If Y_NYU_O_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'ｺｰﾄﾞﾏｽﾀ ＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If

    '仕向け先獲得       2005.12.30
    i = -1
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")
    com = BtOpGetGreater
    SHIMUKE_Flg = False
    
    Do
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> P_KBN04_CD Then
                    Exit Do
                End If
            
                i = i + 1
                ReDim Preserve SHIMUKE_T(0 To i)
            
            
                SHIMUKE_Flg = True
            
                SHIMUKE_T(i).SHIMUKE_CODE = StrConv(P_CODEREC.C_Code, vbUnicode)
                SHIMUKE_T(i).JGYOBU = StrConv(P_CODEREC.OPTION1, vbUnicode)
                SHIMUKE_T(i).NAIGAI = StrConv(P_CODEREC.OPTION2, vbUnicode)
                            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "コードマスタ")
                Unload Me
        End Select
    
        com = BtOpGetNext
    Loop

If Trim(F102025_LOG) <> "" Then                             '2018.03.29
    Call LOG_OUT(F102025_LOG, "F102025 Start")              '2018.03.29
End If                                                      '2018.03.29
    

                                '商品化要／不要の獲得
    GOODS_KBN = "0"
    If GetIni(App.EXEName, "GOODS_KBN", App.EXEName, c) Then
    Else
        GOODS_KBN = Trim(c)
    End If

    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0                 'ﾃﾞﾌｫﾙﾄ昇順
    Next i

    If List_Disp_Proc() Then
        Unload Me
    End If

    Command(0).SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            
If Trim(F102025_LOG) <> "" Then                             '2018.03.29
    Call LOG_OUT(F102025_LOG, "F102025 End")              '2018.03.29
End If                                                      '2018.03.29
                                            
                                            
                                            
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '構成マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
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
                                            '入荷予定ＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "入荷予定")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If


    Set F1020251 = Nothing


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
    F1020251.Caption = "入庫予定表作成処理（" & RTrim(JGYOBU_T(Index).NAME) & ")" & LAST_UPDATE_DAY
    SubMenu(Index).Checked = True
    If Last_JGYOBU <> JGYOBU_T(Index).CODE Then
        Last_JGYOBU = JGYOBU_T(Index).CODE
        LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
        LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)
    End If

End Sub



    

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1020251.MousePointer = vbHourglass

    Call Ctrl_Lock(F1020251)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1020251)


    F1020251.MousePointer = vbDefault

End Sub
Private Function Data_Get_Proc() As Integer
'********************************************************************
'*
'*              入荷予定データ取り込み
'*
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim In_Cnt      As Integer
Dim Out_Cnt     As Integer
    
    
Dim i           As Integer
    
Dim com         As Integer
Dim sts         As Integer
Dim ans         As Integer

Dim c           As String * 128
    
    Data_Get_Proc = True
    
    Call Input_Lock
                                            
                                            '入荷予定ＣＬＯＳＥ
''    sts = BTRV(BtOpClose, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
''    If sts Then
''        If sts <> BtErrNoOpen Then
''            Call File_Error(sts, BtOpClose, "入荷予定(大阪PC)")
''            Exit Function
''        End If
''    End If
''
''
''    sts = GetIni("FILE", Y_NYU_O_ID, "SYS", c)
''    If sts <> False Then
''        Call Log_Out(LOG_F, "SYS.INI [Y_NYU_O]読み込みエラー")
''        Exit Function
''    End If
''
''    On Error Resume Next
''
''    Kill Trim(c)
''
''    On Error GoTo 0
''
''                                '入荷予定 ＯＰＥＮ
''    If Y_NYU_O_Open(BtOpenNomal) Then
''        Exit Function
''    End If
    
    
    com = BtOpGetFirst
    
    Do
        
        DoEvents
        
        Do
            DoEvents
'            sts = BTRV(com + BtSNoWait, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)            '2017.01.27
            sts = BTRV(com, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)                         '2017.01.27
        
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrEOF
                    Exit Do
                                
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                
                Case Else
                    Call File_Error(sts, com, "入荷予定")
                    Exit Function
            End Select
        Loop
        
        If sts = BtErrEOF Then
            Exit Do
        End If
        
        Do
            DoEvents
            sts = BTRV(BtOpDelete, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
        
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                
                Case Else
                    Call File_Error(sts, BtOpDelete, "入荷予定")
                    Exit Function
            End Select
        Loop
        
        
        com = BtOpGetNext
    Loop
    
    In_Cnt = 0
    Out_Cnt = 0
    
'    For i = 0 To UBound(JGYOBU_T)  '2017.12.28 DELETE
        
        
        

'        DoEvents                   '2017.12.28 DELETE

'        If Nyuka_Update_Proc(JGYOBU_T(i).CODE, In_Cnt, Out_Cnt) Then   '入荷予定データ更新処理         2017.12.28
    If EX_Nyuka_Update_Proc("B", In_Cnt, Out_Cnt) Then   '入荷予定データ更新処理       2017.12.28
        Unload Me
    End If
    
'    Next i                         '2017.12.28 DELETE


    Call Input_UnLock

    MsgBox "データ取り込み終了。取り込み件数＝" & Format(Out_Cnt, "#0") & "件です。"
    


    Data_Get_Proc = False

End Function
Private Function Nyuka_Update_Proc(JGYOBU As String, In_Cnt As Integer, Out_Cnt As Integer) As Boolean
'----------------------------------------------------------------------------
'                   「入荷予定データ」更新処理
'----------------------------------------------------------------------------

Dim SOKO_NO         As String
Dim SOKO_T          As String

Dim DEN_NO          As String

Dim MAKER_CODE      As String
Dim MAKER_NAME      As String

Dim HINBAN          As String
Dim SURYO           As String
Dim ORDER_NO        As String
Dim ORDER_NO_1      As String
Dim ORDER_NO_2      As String


Dim NYUKO_YMD       As String
Dim NYUKO_YMD_T     As String

Dim DEN_NO_T        As String



Dim sts             As Integer
Dim ans             As Integer
Dim Ret             As Integer
    
Dim HS_NYUKANo      As Long
Dim HS_NYUKA_OP     As Boolean
    
Dim FileName        As String
Dim Input_Wk        As Variant
Dim Input_Buffer    As String

Dim SKIP_F          As Boolean
Dim FAST_F          As Boolean

Dim NEXT_F          As Integer

Dim c               As String * 128


Dim i               As Integer
    
Dim SEQ_NO          As Long
    
    
    
    Nyuka_Update_Proc = True



    '入荷予定ファイル名取り込み & ＯＰＥＮ
    If GetIni("FILE", "HS_NYUKA", "SYS", c) Then
        Beep
        MsgBox "入荷予定ファイル・ファイル名の獲得に失敗しました。処理を中止して下さい。"
        Exit Function
    End If
    FileName = Trim(c)

    HS_NYUKA_OP = False

    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & "_" & JGYOBU & Right(Trim(FileName), Len(Trim(FileName)) - Ret)
    
    On Error GoTo Exit_Proc
    
    HS_NYUKANo = FreeFile
    Open FileName For Input As #HS_NYUKANo
    On Error GoTo Exit_Proc
    HS_NYUKA_OP = True


    'ユニーク項目のﾀｲﾄﾙ取込み
    
    '倉庫
    SOKO_T = "倉庫"
    If GetIni(App.EXEName, "SOKO_T", App.EXEName, c) Then
    Else
        SOKO_T = Trim(c)
    End If
    
    '入庫日のﾀｲﾄﾙ取込み
    NYUKO_YMD_T = "入庫日 : "
    If GetIni(App.EXEName, "NYUKO_YMD_T", App.EXEName, c) Then
    Else
        NYUKO_YMD_T = Trim(c)
    End If
    
    '伝票№のﾀｲﾄﾙ取込み
    DEN_NO_T = "伝票№"
    If GetIni(App.EXEName, "DEN_NO_T", App.EXEName, c) Then
    Else
        DEN_NO_T = Trim(c)
    End If


    
    FAST_F = True
    NEXT_F = 0


    SEQ_NO = 0

    NYUKO_YMD = ""
    
    Do While Not EOF(HS_NYUKANo)
        Line Input #HS_NYUKANo, Input_Buffer
        
        Input_Wk = Split(Input_Buffer, vbTab, -1)
    
    
    
        In_Cnt = In_Cnt + 1
    
    
    
        If FAST_F Then
            
            
            
            If UBound(Input_Wk) > 1 Then
            
                If InStr(1, Input_Wk(0), SOKO_T) > 0 Then
            
            
                    SOKO_NO = Trim(Left(Input_Wk(1), 2))
            
                End If
            
            End If
            
            
            If UBound(Input_Wk) > 4 Then
            
                If InStr(1, Input_Wk(5), NYUKO_YMD_T) > 0 Then
            
            
                    NYUKO_YMD = Format(Right(Input_Wk(5), 11), "YYYYMMDD")
            
            
                End If
            
            End If
            
            
            If UBound(Input_Wk) >= 0 Then
                
                If InStr(1, Input_Wk(0), DEN_NO_T) > 0 Then
                    FAST_F = False
                End If
            End If
        Else
            If UBound(Input_Wk) < 4 Then
                SKIP_F = True
            Else
                Select Case NEXT_F
                    Case 0
                        MAKER_CODE = Trim(Input_Wk(1))
                        NEXT_F = 1
                
                    Case 1
                        
                        SURYO = Trim(Input_Wk(3))
                        
                        If Len(SURYO) > 2 Then
                            If Left(SURYO, 1) = """" And Right(SURYO, 1) = """" Then
                                SURYO = Mid(SURYO, 2, Len(SURYO) - 2)
                            End If
                        End If
                        
                        If Not IsNumeric(Trim(SURYO)) Then
                            SKIP_F = True
                        Else
                            DEN_NO = Trim(Input_Wk(0))
                            If Trim(DEN_NO) = "0" Then
                                DEN_NO = "000000"
                            End If
                            
                            MAKER_NAME = Trim(Input_Wk(1))
                            
                            HINBAN = Trim(Input_Wk(2))
                            
                            
                            
                            
                            ORDER_NO_1 = Trim(Input_Wk(4))
                
                        End If
                        NEXT_F = 2
                    
                    Case 2
                        ORDER_NO_2 = Trim(Input_Wk(4))
                        NEXT_F = 3
                End Select
            End If
        
            If NEXT_F = 3 Then
                If Not SKIP_F And Not FAST_F Then
                                                    
                    ORDER_NO = Trim(ORDER_NO_1) & Trim(ORDER_NO_2)
                                                    
                                                    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2017.01.27
'                                                    'トランザクション開始
'                    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                    If sts <> BtNoErr Then
'                        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
'                        Exit Function
'                    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2017.01.27
                                                '品目マスタチェック
                    If Item_Check_Proc(In_Mode, JGYOBU, NAIGAI_NAI, HINBAN, , , MAKER_CODE, MAKER_NAME) Then
                        GoTo Abort_Tran
                    End If
                                                
                                                '入荷データ作成
                    
                    
                    '事業部
                    Call UniCode_Conv(Y_NYU_O_REC.JGYOBU, JGYOBU)
                    '倉庫
                    Call UniCode_Conv(Y_NYU_O_REC.SOKO_NO, SOKO_NO)
                    'データSEQ
                    SEQ_NO = SEQ_NO + 1
                    Call UniCode_Conv(Y_NYU_O_REC.SEQ_NO, Format(SEQ_NO, "000"))
                    '入庫日
                    Call UniCode_Conv(Y_NYU_O_REC.NYUKO_YMD, NYUKO_YMD)
                    '伝票№
                    Call UniCode_Conv(Y_NYU_O_REC.DEN_NO, DEN_NO)
                    'ﾒｰｶｰｺｰﾄﾞ
                    Call UniCode_Conv(Y_NYU_O_REC.MAKER_CODE, MAKER_CODE)
                    '国内外
                    Call UniCode_Conv(Y_NYU_O_REC.NAIGAI, NAIGAI_NAI)
                    '品番
                    Call UniCode_Conv(Y_NYU_O_REC.HIN_NO, HINBAN)
                    '予定数量
                    Call UniCode_Conv(Y_NYU_O_REC.Y_SURYO, Format(CLng(SURYO), "00000000"))
                    '実績数量
                    Call UniCode_Conv(Y_NYU_O_REC.J_SURYO, "00000000")
                    '検収担当者
                    Call UniCode_Conv(Y_NYU_O_REC.TANTO_CODE, "")
                    '注文№
                    Call UniCode_Conv(Y_NYU_O_REC.ORDER_NO, ORDER_NO)
                    
                    '検品F
                    Call UniCode_Conv(Y_NYU_O_REC.KENPIN_F, KAN_KBN_UN)
                    
                    Call UniCode_Conv(Y_NYU_O_REC.ORDER_NO, ORDER_NO)
                    
                    
                    Call UniCode_Conv(Y_NYU_O_REC.WEL_ID, "")
                    
                    
                    Call UniCode_Conv(Y_NYU_O_REC.PRG_ID, "")
                    
                    Call UniCode_Conv(Y_NYU_O_REC.FILLER, "")
                    
                    Do
                        sts = BTRV(BtOpInsert, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpInsert, "入荷予定")
                                Exit Function
                        End Select
                    Loop
            
                            
            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2017.01.27
'                    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                    If sts <> BtNoErr Then
'                        GoTo Abort_Tran
'                    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2017.01.27
    
            
                    Out_Cnt = Out_Cnt + 1
                End If
                
                NEXT_F = 0
                SKIP_F = False
            
            End If
                        
            
            
            
            DoEvents

        End If
    
    Loop

    Nyuka_Update_Proc = False
'    Exit Function              '2017.01.27
    GoTo Exit_Proc              '2017.01.27

Abort_Tran:
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2017.01.27
'    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'    If sts <> BtNoErr Then
'        Call File_Error(sts, BtOpAbortTransaction, "")
'    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2017.01.27


Exit_Proc:
    
    If HS_NYUKA_OP Then
        Close #HS_NYUKANo
    End If
    

End Function


Private Function EX_Nyuka_Update_Proc(JGYOBU As String, In_Cnt As Integer, Out_Cnt As Integer) As Boolean
'----------------------------------------------------------------------------
'                   「入荷予定データ」更新処理
'                   EXCEL 対応  2017.12.28
'----------------------------------------------------------------------------

Dim SOKO_NO         As String
Dim DEN_NO          As String
Dim MAKER_CODE      As String
Dim MAKER_NAME      As String
Dim HIN_NO          As String
Dim SURYO           As String
Dim ORDER_NO        As String
Dim ORDER_NO_1      As String
Dim ORDER_NO_2      As String
Dim NYUKO_YMD       As String



Dim sts             As Integer
Dim ans             As Integer
Dim Ret             As Integer
    
Dim HS_NYUKANo      As Long
Dim HS_NYUKA_OP     As Boolean
    
Dim FileName        As String

Dim SKIP_F          As Boolean


Dim c               As String * 128


Dim i               As Long
    
Dim SEQ_NO          As Long
    
Dim END_GYO         As Long
    
Dim xlApp           As Object
Dim xlBook          As Object
Dim xlSheet         As Object
    
    
    
    EX_Nyuka_Update_Proc = True



    '入荷予定ファイル名取り込み & ＯＰＥＮ
    If GetIni(App.EXEName, "EX_FILE", App.EXEName, c) Then
        Beep
        MsgBox "入荷予定ファイル・ファイル名の獲得に失敗しました。処理を中止して下さい。"
        Exit Function
    End If
    FileName = Trim(c)



    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    
    
    On Error GoTo Error_Proc

    xlApp.Workbooks.Open (FileName), ReadOnly:=True, UpdateLinks:=0

    On Error GoTo 0


    Set xlSheet = xlApp.Worksheets(1)
    xlSheet.Activate





    SOKO_NO = Left(xlSheet.Application.Cells(4, 2), 2)                              '倉庫№
    NYUKO_YMD = Format(Right(xlSheet.Application.Cells(4, 6), 11), "YYYYMMDD")      '入庫日


    SEQ_NO = 0
    
    END_GYO = 0
    SKIP_F = False

    i = 6
    Do
        
        i = i + 1
        If Trim(xlSheet.Application.Cells(i, EX_MAKER_CODE)) = "" Then
        
            SKIP_F = True
            END_GYO = END_GYO + 1
            
            If END_GYO > 5 Then
                Exit Do
            End If
        Else

            In_Cnt = In_Cnt + 1
            
            MAKER_CODE = Trim(xlSheet.Application.Cells(i, EX_MAKER_CODE))
            i = i + 1
            DEN_NO = Trim(xlSheet.Application.Cells(i, EX_DEN_NO))
            MAKER_NAME = Trim(xlSheet.Application.Cells(i, EX_MAKER_NAME))
            HIN_NO = Trim(xlSheet.Application.Cells(i, EX_HIN_NO))
            SURYO = Trim(xlSheet.Application.Cells(i, EX_Y_SURYO))
            If Not IsNumeric(SURYO) Then
                SURYO = "0"
            End If
                        
            ORDER_NO_1 = Trim(xlSheet.Application.Cells(i, EX_ORDER_NO_1))
            i = i + 1
            ORDER_NO_2 = Trim(xlSheet.Application.Cells(i, EX_ORDER_NO_2))
            ORDER_NO = ORDER_NO_1 & ORDER_NO_2
                                        '品目マスタチェック
            If Item_Check_Proc(In_Mode, JGYOBU, NAIGAI_NAI, HIN_NO, , , MAKER_CODE, MAKER_NAME) Then
                GoTo Exit_Proc
            End If
                                                
            '入荷データ作成
            '事業部
            Call UniCode_Conv(Y_NYU_O_REC.JGYOBU, JGYOBU)
            '倉庫
            Call UniCode_Conv(Y_NYU_O_REC.SOKO_NO, SOKO_NO)
            'データSEQ
            SEQ_NO = SEQ_NO + 1
            Call UniCode_Conv(Y_NYU_O_REC.SEQ_NO, Format(SEQ_NO, "000"))
            '入庫日
            Call UniCode_Conv(Y_NYU_O_REC.NYUKO_YMD, NYUKO_YMD)
            '伝票№
            Call UniCode_Conv(Y_NYU_O_REC.DEN_NO, DEN_NO)
            'ﾒｰｶｰｺｰﾄﾞ
            Call UniCode_Conv(Y_NYU_O_REC.MAKER_CODE, MAKER_CODE)
            '国内外
            Call UniCode_Conv(Y_NYU_O_REC.NAIGAI, NAIGAI_NAI)
            '品番
            Call UniCode_Conv(Y_NYU_O_REC.HIN_NO, HIN_NO)
            '予定数量
            Call UniCode_Conv(Y_NYU_O_REC.Y_SURYO, Format(CLng(SURYO), "00000000"))
            '実績数量
            Call UniCode_Conv(Y_NYU_O_REC.J_SURYO, "00000000")
            '検収担当者
            Call UniCode_Conv(Y_NYU_O_REC.TANTO_CODE, "")
            '注文№
            Call UniCode_Conv(Y_NYU_O_REC.ORDER_NO, ORDER_NO)
            
            '検品F
            Call UniCode_Conv(Y_NYU_O_REC.KENPIN_F, KAN_KBN_UN)
            
            Call UniCode_Conv(Y_NYU_O_REC.ORDER_NO, ORDER_NO)
            
            
            Call UniCode_Conv(Y_NYU_O_REC.WEL_ID, "")
            
            
            Call UniCode_Conv(Y_NYU_O_REC.PRG_ID, "")
            
            Call UniCode_Conv(Y_NYU_O_REC.FILLER, "")
            
            Do
                sts = BTRV(BtOpInsert, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpInsert, "入荷予定")
                        Exit Function
                End Select
            Loop
            
    
            
            Out_Cnt = Out_Cnt + 1
            SKIP_F = False
            END_GYO = 0
        End If
        
        
        DoEvents
            
    Loop
                        
            
            

    EX_Nyuka_Update_Proc = False
    GoTo Exit_Proc              '2017.01.27

    


Exit_Proc:
    
    xlApp.DisplayAlerts = False

    xlBook.Close False
    xlApp.Quit                  'EXCELを閉じる
    Set xlApp = Nothing
    
    Exit Function
Error_Proc:
    

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
        Case 52, 53, 54, 55, 57, 59, 61, 62, 63, 68, 70, 71, 75, 76, 1004
            
            
            MsgBox "指定のファイルが見つかりません。" & Chr(13) & Chr(10) & "正しいファイル名を入力してください。"
            
            
            xlApp.DisplayAlerts = False
        
            xlBook.Close False
            xlApp.Quit                  'EXCELを閉じる
            Set xlApp = Nothing
            
            
            
            EX_Nyuka_Update_Proc = False
            


    '2011.12.03
        Case 13
        
            MsgBox "読込み対象のEXCELデータに異常が有ります。内容を確認後、再実行してください。"
            
            xlApp.DisplayAlerts = False
            xlBook.Close False
            xlApp.Quit 'EXCELを閉じる
            Set xlApp = Nothing
            
            
            EX_Nyuka_Update_Proc = False
    
            
            

        Case Else
    End Select
    

End Function



Private Function Item_Check_Proc(Mode As Integer, _
                                    JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    Optional HIN_NAME As String = "", _
                                    Optional LOCATION As String = "", _
                                    Optional MAKER_CODE As String = "", _
                                    Optional MAKER_NAME As String = "") As Integer
'----------------------------------------------------------------------------
'                   「品目マスタ」チェック＆更新処理
'----------------------------------------------------------------------------
Dim com         As Integer
Dim sts         As Integer
Dim ans         As Integer
        
Dim i           As Integer
    
    
    Item_Check_Proc = True

           

    Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)

    Do

        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                
                com = BtOpUpdate
                                
                If Trim(HIN_NAME) <> "" Then
                    Call UniCode_Conv(ITEMREC.HIN_NAME, HIN_NAME)   '品名
                End If
                
                
                If Trim(MAKER_CODE) <> "" Then
                    Call UniCode_Conv(ITEMREC.MAKER_CODE, MAKER_CODE)       'ﾒｰｶｰｺｰﾄﾞ
                End If
                
                If Trim(MAKER_NAME) <> "" Then
                    Call UniCode_Conv(ITEMREC.MAKER_NAME, MAKER_NAME)       'ﾒｰｶｰ名称
                End If
                
                
                
                Exit Do
            Case BtErrKeyNotFound
                
                com = BtOpInsert
                
                Call UniCode_Conv(ITEMREC.JGYOBU, JGYOBU)           '事業部
                Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI)           '国内外
                Call UniCode_Conv(ITEMREC.HIN_GAI, HIN_GAI)         '品番（外部）
    
                Call UniCode_Conv(ITEMREC.HIN_NAME, HIN_NAME)       '品名
    
                Call UniCode_Conv(ITEMREC.ST_SET_DT, "")            '標準棚番設定日
                
                
                                                                    '標準棚番
                If Len(Trim(LOCATION)) > 6 Then
                    Call UniCode_Conv(ITEMREC.ST_SOKO, Mid(LOCATION, 1, 2))
                    Call UniCode_Conv(ITEMREC.ST_RETU, Mid(LOCATION, 3, 2))
                    Call UniCode_Conv(ITEMREC.ST_REN, Mid(LOCATION, 5, 2))
                    Call UniCode_Conv(ITEMREC.ST_DAN, "01")
                
                Else
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
                    Call UniCode_Conv(ITEMREC.ST_REN, "")
                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
                End If
    
    
                Call UniCode_Conv(ITEMREC.BEF_SOKO, "")             '前回入庫倉庫
                Call UniCode_Conv(ITEMREC.BEF_RETU, "")
                Call UniCode_Conv(ITEMREC.BEF_REN, "")
                Call UniCode_Conv(ITEMREC.BEF_DAN, "")
    
                Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")          '最終入庫日
                Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")          '最終出庫日
    
                Call UniCode_Conv(ITEMREC.HIN_NAI, "")              '品番（内部）
    
                Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")           '備考 ホスト倉庫
                Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")           '備考 ホスト棚番
                
                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")      '補充点
                Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")    '月平均出荷数
                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "1")          'サンプル数
                
                Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")          '最終入荷日付
                Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")          '最終照合日付
                
                Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "")         '最終照合時在庫数
                
                Call UniCode_Conv(ITEMREC.BIKOU, "")                '印刷備考
                Call UniCode_Conv(ITEMREC.IRI_QTY, "")              '印刷入り数
                
                Call UniCode_Conv(ITEMREC.JAN_CODE, "")             'Janコード
                Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")           '品番読み替え
                
                Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_KBN)     '商品化有無（有）
                
                Call UniCode_Conv(ITEMREC.PACKING_NO, "")           '個装箱№
                
                Call UniCode_Conv(ITEMREC.RANK, "")                 '現在ﾗﾝｸ
                Call UniCode_Conv(ITEMREC.NEW_RANK, "")             '新ﾗﾝｸ
                
                
                Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")          'ｸﾞﾘｯｸｽ棚番1
                Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")          'ｸﾞﾘｯｸｽ棚番1
                Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")          'ｸﾞﾘｯｸｽ棚番1
                
                
                Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")             '業務管理　 仕入区分
                Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, "")             '           販売区分
                Call UniCode_Conv(ITEMREC.G_SYUSHI, "")                 '           収支単位
                Call UniCode_Conv(ITEMREC.G_KUMITATE, "")               '           組立製品
                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")              '           標準粗利売価単価　9(8)V99
                Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")           '           標準粗利売価設定日
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")              '           標準粗利原価単価  9(8)V99
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")           '           標準粗利原価設定日
                                            
                                            
                                                                        '           仕入先情報
                For i = 0 To 2
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).CODE, "")             'ｺｰﾄﾞ
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA, "")            '仕入単価
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, "")         '単価設定日
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LOT, "")              'ﾛｯﾄ数
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LEAD_TIME, "")        'ﾘｰﾄﾞﾀｲﾑ
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, "")    'ﾘｰﾄﾞﾀｲﾑ
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, "")   'ﾘｰﾄﾞﾀｲﾑ
                
                Next i
                                            
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")          '           前月在庫金額
                Call UniCode_Conv(ITEMREC.G_SHIZAI_KBN, "")             '           資材区分
                Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_ON)    '           ラベル貼付
                
                Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")             '商品ﾗﾍﾞﾙ   品名
                Call UniCode_Conv(ITEMREC.L_BIKOU, "")                  '           備考
                Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")            '           会社コード
                Call UniCode_Conv(ITEMREC.L_KISHU1, "")                 '           機種(1)
                Call UniCode_Conv(ITEMREC.xL_KISHU2, "")                '           機種(2)未使用
                Call UniCode_Conv(ITEMREC.L_KISHU3, "")                 '           機種(3)
                Call UniCode_Conv(ITEMREC.L_PAPER, "")                  '           紙
                Call UniCode_Conv(ITEMREC.L_PLASTIC, "")                '           プラスチック
                Call UniCode_Conv(ITEMREC.L_URIKIN1, "")                '           価格(1)
                Call UniCode_Conv(ITEMREC.L_URIKIN2, "")                '           価格(2)
                Call UniCode_Conv(ITEMREC.L_URIKIN3, "")                '           価格(3)
                Call UniCode_Conv(ITEMREC.L_LABEL, "")                  '           適用機種ﾗﾍﾞﾙ
                Call UniCode_Conv(ITEMREC.L_MAISU, "")                  '           枚数ﾗﾍﾞﾙ
                Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")            '           適用機種備考
                Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")            '           作業指示
                Call UniCode_Conv(ITEMREC.L_BIKOU3, "")                 '           備考３
                Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")            '           事業部コード
                Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")                '           入り数
                Call UniCode_Conv(ITEMREC.L_TANA1, "")                  '           棚番(1)
                Call UniCode_Conv(ITEMREC.L_TANA2, "")                  '           棚番(2)
                
                Call UniCode_Conv(ITEMREC.S_TANTO, "")                  '収単／担当者コード
                Call UniCode_Conv(ITEMREC.ZAIKO_F, P_ZAIKO_F_ON)        '在庫管理対象有無　（対象）
                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '機種(2)
                
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "00000000")  '前月在庫数
                Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, "00000000") '最終出荷数
                            
                Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, "00000000")     'S2 在庫
                Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, "00000000")     'P2 在庫
                            
                Call UniCode_Conv(ITEMREC.K_KEITAI, "")                 '個装形態
                            
                Call UniCode_Conv(ITEMREC.UNIT_BUHIN, "")               'ﾕﾆｯﾄ部品区分
                Call UniCode_Conv(ITEMREC.NAI_BUHIN, "")                '国内供給区分
                Call UniCode_Conv(ITEMREC.GAI_BUHIN, "")                '海外供給区分
                Call UniCode_Conv(ITEMREC.HYO_TANKA, "")                '標準単価
    
    
                Call UniCode_Conv(ITEMREC.MAKER_CODE, MAKER_CODE)       'ﾒｰｶｰｺｰﾄﾞ
                Call UniCode_Conv(ITEMREC.MAKER_NAME, MAKER_NAME)       'ﾒｰｶｰ名称
            
    
                            
                Call UniCode_Conv(ITEMREC.FILLER, "")
                                                                        '更新担当者
                Call UniCode_Conv(ITEMREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
                                                                        '更新日時
                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                Exit Function
        End Select
    Loop
    
    Do
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Exit Function
        End Select
    Loop
        
    If SHIMUKE_Flg Then
        If com = BtOpInsert Then
            '構成マスタの追加       2005.12.30
            For i = 0 To UBound(SHIMUKE_T)
            
                If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIMUKE_T(i).JGYOBU And _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) = SHIMUKE_T(i).NAIGAI Then
                                                                            '仕向け先ｺｰﾄﾞ
                    Call UniCode_Conv(P_COMPO_O_REC.SHIMUKE_CODE, SHIMUKE_T(i).SHIMUKE_CODE)
                                                                            '事業部
                    Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, SHIMUKE_T(i).JGYOBU)
                                                                            '国内外
                    Call UniCode_Conv(P_COMPO_O_REC.NAIGAI, SHIMUKE_T(i).NAIGAI)
                                                                            '品番
                    Call UniCode_Conv(P_COMPO_O_REC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                                                            'ﾃﾞｰﾀ区分
                    Call UniCode_Conv(P_COMPO_O_REC.DATA_KBN, P_HEAD)
                                                                            '追番
                    Call UniCode_Conv(P_COMPO_O_REC.SEQNO, "000")
                                                                            '基本クラス
                    Call UniCode_Conv(P_COMPO_O_REC.CLASS_CODE, "")
                                                                            '備考
                    Call UniCode_Conv(P_COMPO_O_REC.BIKOU, "")
                    
                    Call UniCode_Conv(P_COMPO_O_REC.FILLER, "")
                                                                            '更新担当者
                    Call UniCode_Conv(P_COMPO_O_REC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
                                                                            '更新日時
                    Call UniCode_Conv(P_COMPO_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                
                
                    Do
                        sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Select Case sts
                            Case BtNoErr, BtErrDuplicates
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpInsert, "構成マスタ")
                                Exit Function
                        End Select
                    Loop
                
                
                End If
            Next i
        
        End If
        
    End If

    Item_Check_Proc = False

End Function


Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   データ表示
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim Row         As Long
    
    
Dim Skip_flg    As Boolean
    
    
    List_Disp_Proc = True
                                    
    F1020251.MousePointer = vbHourglass
                                    
    DATA_FLG = False
                                    
                                    'テーブルリセット
    Set NYUKA = Nothing
    
    Row = Min_Row - 1
    
    com = BtOpGetFirst
    
    Do
        
        DoEvents
        
        sts = BTRV(com, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
    
        Select Case sts
            Case BtNoErr
        
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "入荷予定")
                List_Disp_Proc = SYS_ERR
                Exit Function
        End Select
            
        If Row = (Min_Row - 1) Then
            Text1(ptxNYUKO_YY).Text = Mid(StrConv(Y_NYU_O_REC.NYUKO_YMD, vbUnicode), 1, 4)
            Text1(ptxNYUKO_MM).Text = Mid(StrConv(Y_NYU_O_REC.NYUKO_YMD, vbUnicode), 5, 2)
            Text1(ptxNYUKO_DD).Text = Mid(StrConv(Y_NYU_O_REC.NYUKO_YMD, vbUnicode), 7, 2)
        End If
            
        Row = Row + 1
        If Row > Max_Row Then
            Beep
            MsgBox "最大表示行数を超えました。"
            Exit Do
        End If
        DATA_FLG = True
                
        If Grid_Set_Proc(Row) Then
            Exit Function
        End If
        
        com = BtOpGetNext
        
    Loop
    
    Set TDBGrid1.Array = NYUKA
    
    TDBGrid1.style.Locked = False
    
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    
'''    Call Input_UnLock
    F1020251.MousePointer = vbDefault
    
    
    
    List_Disp_Proc = False

    
End Function


Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'                   FILE--->GLID
'----------------------------------------------------------------------------

Dim sts As Integer

    
    Grid_Set_Proc = True

    

    NYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
    
    '伝票№
    NYUKA(Row, colDEN_NO) = Trim(StrConv(Y_NYU_O_REC.DEN_NO, vbUnicode))
    'ﾒｰｶｰ
    NYUKA(Row, colMAKER_CODE) = Trim(StrConv(Y_NYU_O_REC.MAKER_CODE, vbUnicode))
    If Trim(StrConv(Y_NYU_O_REC.MAKER_CODE, vbUnicode)) = "" Then
        Call UniCode_Conv(ITEMREC.MAKER_NAME, "")
    Else
    
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_NYU_O_REC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_NYU_O_REC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_NYU_O_REC.HIN_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    
        Select Case sts
            Case BtNoErr
            
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.MAKER_NAME, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                Exit Function
        End Select
    End If
    'ﾒｰｶｰ名称
    NYUKA(Row, colMAKER_NAME) = Trim(StrConv(ITEMREC.MAKER_NAME, vbUnicode))
    '品番
    NYUKA(Row, colHIN_NO) = Trim(StrConv(Y_NYU_O_REC.HIN_NO, vbUnicode))
    '予定数量
    If IsNumeric(StrConv(Y_NYU_O_REC.Y_SURYO, vbUnicode)) Then
        NYUKA(Row, colY_SURYO) = Format(CLng(StrConv(Y_NYU_O_REC.Y_SURYO, vbUnicode)), "#0")
    Else
        NYUKA(Row, colY_SURYO) = ""
    End If
    
    '入庫数量
    If IsNumeric(StrConv(Y_NYU_O_REC.J_SURYO, vbUnicode)) Then
        NYUKA(Row, colJ_SURYO) = Format(CLng(StrConv(Y_NYU_O_REC.J_SURYO, vbUnicode)), "#0")
    Else
        NYUKA(Row, colJ_SURYO) = ""
    End If
    
    '担当者
    NYUKA(Row, colTANTO_CODE) = Trim(StrConv(Y_NYU_O_REC.TANTO_CODE, vbUnicode))
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(Y_NYU_O_REC.TANTO_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)

    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(TANTOREC.TANTO_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ ")
            Exit Function
    End Select
    NYUKA(Row, colTANTO_NAME) = Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
    
    '注文№
    NYUKA(Row, colORDER_NO) = Trim(StrConv(Y_NYU_O_REC.ORDER_NO, vbUnicode))
    
    '注文№
    NYUKA(Row, colKENPIN_F) = Trim(StrConv(Y_NYU_O_REC.KENPIN_F, vbUnicode))
    
    
    
    
If Trim(F102025_LOG) <> "" Then                             '2018.03.29
    Call LOG_OUT(F102025_LOG, "伝票№=" & NYUKA(Row, colDEN_NO) & "　メーカー=" & NYUKA(Row, colMAKER_CODE) & "　予定数量=" & NYUKA(Row, colY_SURYO) & "　注文№=" & NYUKA(Row, colORDER_NO))        '2018.03.29
End If                                                      '2018.03.29
    
    
    
    
    Grid_Set_Proc = False
End Function


Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    If Sort_Tbl(ColIndex) = 0 Then
        Sort_Tbl(ColIndex) = 1
    Else
        If Sort_Tbl(ColIndex) = 1 Then
            Sort_Tbl(ColIndex) = 0
        End If
    
    End If
        
    If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then
                    
        NYUKA.QuickSort Min_Row, NYUKA.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
        
        Set TDBGrid1.Array = NYUKA
        
        TDBGrid1.ReBind
        TDBGrid1.Update
        TDBGrid1.MoveFirst


    End If

End Sub
Private Function EXCEL_Put_Proc(Print_Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   「入庫予定表」出力ﾒｲﾝ処理
'----------------------------------------------------------------------------


Dim strExelFile     As String
Dim Rec_Cnt         As Long
Dim Page_Offset     As Long
Dim posG            As Long

Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer
Dim i               As Integer
Dim Skip_flg        As Boolean

Dim MyChan          As Long


'2011.04.17
Dim ExcelApp        As Object
Dim Excelbook       As Object
Dim ExcelWorkSheet  As Object
'2011.04.17


'On Error GoTo ERR_PRT


    EXCEL_Put_Proc = True





hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
"EXCEL_Put_Proc Start", Me.hwnd, 0)


'If Trim(F102025_LOG) <> "" Then                             '2017.01.05
'    Call LOG_OUT(F102025_LOG, "EXCEL_Put_Proc Start")       '2017.01.05
'End If                                                      '2017.01.05



    Call Input_Lock

                                    '出力ﾌｧｲﾙ名編集
    strExelFile = Excel_PutPath & Text1(ptxNYUKO_YY).Text & _
                                    Text1(ptxNYUKO_MM).Text & _
                                    Text1(ptxNYUKO_DD).Text & ".xls"

    'Excelｱﾌﾟﾘｹｰｼｮﾝｵﾌﾞｼﾞｪｸﾄ取得
    
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
"CreateObject(" & "Excel.Application" & ")", Me.hwnd, 0)

'If Trim(F102025_LOG) <> "" Then                             '2017.01.05
'    Call LOG_OUT(F102025_LOG, "CreateObject(" & "Excel.Application" & ")")
'End If
    
    Set ExcelApp = CreateObject("Excel.Application")


    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
"ExcelApp.Workbooks.Open(Excel_Template)", Me.hwnd, 0)

'If Trim(F102025_LOG) <> "" Then                             '2017.01.05
'    Call LOG_OUT(F102025_LOG, "ExcelApp.Workbooks.Open(Excel_Template)")
'End If
    
    
    Set Excelbook = ExcelApp.Workbooks.Open(Excel_Template)         'ﾃﾝﾌﾟﾚｰﾄﾌﾞｯｸを開く
    
    
    
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
"Excelbook.Worksheets(1)", Me.hwnd, 0)

'If Trim(F102025_LOG) <> "" Then                                '2017.01.05
'    Call LOG_OUT(F102025_LOG, "Excelbook.Worksheets(1)")
'End If
    
    
    Set ExcelWorkSheet = Excelbook.Worksheets(1)                    '１ｼｰﾄ目を選択




    '発行日
    ExcelWorkSheet.Application.Cells(4, 10).Value = "発行日：" & Format(Now, "yyyy年mm月dd日")
    '倉庫
    ExcelWorkSheet.Application.Cells(6, 4).Value = Excel_Soko_Name
    '入庫日
    ExcelWorkSheet.Application.Cells(6, 10).Value = "入庫日：" & Text1(ptxNYUKO_YY).Text & "年" & _
                                                                 Text1(ptxNYUKO_MM).Text & "月" & _
                                                                 Text1(ptxNYUKO_DD).Text & "日"
    Rec_Cnt = 0
    Page_Offset = 9
    posG = 9

    com = BtOpGetFirst
    Do
        sts = BTRV(com, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "入庫予定")
                Exit Function
        End Select


        '2011.04.17
'''        If EXCEL_Set_Proc(posG, Page_Offset) Then     '１行分編集
        If EXCEL_Set_Proc(posG, Page_Offset, ExcelApp, Excelbook, ExcelWorkSheet) Then     '１行分編集
        '2011.04.17
            Exit Function
        End If
        Rec_Cnt = Rec_Cnt + 1
        com = BtOpGetNext
        DoEvents
    Loop

    '当該ページの残り行をクリア
    If posG <= Page_Offset + 33 Then
        Call UniCode_Conv(Y_NYU_O_REC.DEN_NO, "")
        Call UniCode_Conv(Y_NYU_O_REC.MAKER_CODE, "")
        Call UniCode_Conv(Y_NYU_O_REC.HIN_NO, "")
        Call UniCode_Conv(Y_NYU_O_REC.Y_SURYO, "")
        Call UniCode_Conv(Y_NYU_O_REC.J_SURYO, "")
        Call UniCode_Conv(Y_NYU_O_REC.TANTO_CODE, "")
        Call UniCode_Conv(Y_NYU_O_REC.ORDER_NO, "")
        Do
            If posG > Page_Offset + 33 Then
                Exit Do
            End If
        '2011.04.17
'''             If EXCEL_Set_Proc(posG, Page_Offset) Then     '１行分編集
            If EXCEL_Set_Proc(posG, Page_Offset, ExcelApp, Excelbook, ExcelWorkSheet) Then     '１行分編集
        '2011.04.17
                Exit Function
            End If
        Loop
    End If


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
"ExcelApp.Visible = False", Me.hwnd, 0)


'If Trim(F102025_LOG) <> "" Then                                '2017.01.05
'    Call LOG_OUT(F102025_LOG, "ExcelApp.Visible = False")
'End If




    '編集時にﾜｰｸｼｰﾄの先頭が表示される様に「A1」をｱｸﾃｨﾌﾞにする
    ExcelWorkSheet.Application.Range("A1").Activate

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
"ExcelWorkSheet.Application.Range(A1).Activate", Me.hwnd, 0)

'If Trim(F102025_LOG) <> "" Then                            '2017.01.05
'    Call LOG_OUT(F102025_LOG, "ExcelWorkSheet.Application.Range(A1).Activate")
'End If

    ExcelApp.Visible = False


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
"ExcelApp.DisplayAlerts = False", Me.hwnd, 0)

'If Trim(F102025_LOG) <> "" Then                            '2017.01.05
'    Call LOG_OUT(F102025_LOG, "ExcelApp.DisplayAlerts = False")
'End If


    ExcelApp.DisplayAlerts = False              'ﾏｸﾛ実行ｴﾗｰは表示しない

    If Print_Mode = 1 Then
        ExcelWorkSheet.PrintOut
    End If

    If Rec_Cnt > 0 Then
        On Error Resume Next
        Kill strExelFile
        ExcelWorkSheet.SaveAs strExelFile
        On Error GoTo 0
    End If
    

    ExcelApp.ScreenUpdating = True              'INS 2013.03.19

    ExcelApp.Visible = True                     'INS 2013.03.19


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
"ExcelApp.Workbooks.Close", Me.hwnd, 0)

'If Trim(F102025_LOG) <> "" Then                            '2017.01.05
'    Call LOG_OUT(F102025_LOG, "ExcelApp.Workbooks.Close")
'End If

'    ExcelApp.Workbooks.Close                    'ﾜｰｸﾌﾞｯｸ閉じる DEL 2013.03.19
    
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
"ExcelApp.Quit", Me.hwnd, 0)

'If Trim(F102025_LOG) <> "" Then                            '2017.01.05
'    Call LOG_OUT(F102025_LOG, "ExcelApp.Quit")
'End If
    
'    ExcelApp.Quit                              DEL 2013.03.19


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
"Set ExcelWorkSheet = Nothing", Me.hwnd, 0)

'If Trim(F102025_LOG) <> "" Then                            '2017.01.05
'    Call LOG_OUT(F102025_LOG, "Set ExcelWorkSheet = Nothing")
'End If


    Set ExcelWorkSheet = Nothing                                    'ﾜｰｸｼｰﾄ開放
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
"Set Excelbook = Nothing", Me.hwnd, 0)

'If Trim(F102025_LOG) <> "" Then                            '2017.01.05
'    Call LOG_OUT(F102025_LOG, "Set Excelbook = Nothing")
'End If
    
    Set Excelbook = Nothing                                         'ﾜｰｸﾌﾞｯｸ開放
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
"Set ExcelApp = Nothing", Me.hwnd, 0)

'If Trim(F102025_LOG) <> "" Then                            '2017.01.05
'    Call LOG_OUT(F102025_LOG, "Set ExcelApp = Nothing")
'End If
    
    Set ExcelApp = Nothing                                          'ﾜｰｸﾌﾞｯｸ開放


'hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
'"sts = Shell(EXCEL_DIR &  strExelFile, vbMaximizedFocus)", Me.hwnd, 0)
'
'
'If Trim(F102025_LOG) <> "" Then                            '2017.01.05
'    Call LOG_OUT(F102025_LOG, "sts = Shell(EXCEL_DIR &  strExelFile, vbMaximizedFocus)")
'End If

'    sts = Shell(EXCEL_DIR & " " & strExelFile, vbMaximizedFocus)       DEL 2013.03.19




'    MyChan = DDEInitiate("Excel", "System")
'    DDEExecute MyChan, "[open(""" & strExelFile & """)]" '--- B
'    DDETerminate MyChan




    Call Input_UnLock

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
"EXCEL_Put_Proc End", Me.hwnd, 0)

'If Trim(F102025_LOG) <> "" Then                            '2017.01.05
'    Call LOG_OUT(F102025_LOG, "EXCEL_Put_Proc End")
'End If

    EXCEL_Put_Proc = False

End Function

Private Function EXCEL_Set_Proc(posG As Long, Page_Offset As Long, ExcelApp As Object, Excelbook As Object, ExcelWorkSheet As Object) As Integer




Dim c   As String * 128
Dim sts As Integer

    EXCEL_Set_Proc = True
    
    
    '１頁分編集完了⇒次頁分のフォーマットをコピー
    If posG > Page_Offset + 33 Then
        ExcelWorkSheet.Application.Range(Page_Offset - 8 & ":" & Page_Offset - 8 + 46).Copy
        ExcelWorkSheet.Application.Range(Page_Offset - 8 + 46 & ":" & Page_Offset - 8 + 92).Select
        ExcelWorkSheet.Paste

        Page_Offset = Page_Offset + 46
        posG = Page_Offset
    End If
    
    
    
    
    'ﾒｰｶｰｺｰﾄﾞ
    ExcelWorkSheet.Application.Cells(posG, 4).Value = Trim(StrConv(Y_NYU_O_REC.MAKER_CODE, vbUnicode))

    '伝票№
    ExcelWorkSheet.Application.Cells(posG + 1, 3).Value = Trim(StrConv(Y_NYU_O_REC.DEN_NO, vbUnicode))

    If Trim(StrConv(Y_NYU_O_REC.MAKER_CODE, vbUnicode)) = "" Then
        Call UniCode_Conv(ITEMREC.MAKER_NAME, "")
    Else
        'ﾒｰｶｰ名
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_NYU_O_REC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_NYU_O_REC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_NYU_O_REC.HIN_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.MAKER_NAME, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
                Exit Function
        End Select
    End If
    ExcelWorkSheet.Application.Cells(posG + 1, 4).Value = Trim(StrConv(ITEMREC.MAKER_NAME, vbUnicode))

    ExcelWorkSheet.Application.Cells(posG + 1, 5).Value = Trim(StrConv(Y_NYU_O_REC.HIN_NO, vbUnicode))

    If IsNumeric(StrConv(Y_NYU_O_REC.Y_SURYO, vbUnicode)) Then
        ExcelWorkSheet.Application.Cells(posG + 1, 6).Value = Format(CLng(StrConv(Y_NYU_O_REC.Y_SURYO, vbUnicode)), "#0")
    Else
        ExcelWorkSheet.Application.Cells(posG + 1, 6).Value = ""
    End If
    
    If IsNumeric(StrConv(Y_NYU_O_REC.J_SURYO, vbUnicode)) Then
        ExcelWorkSheet.Application.Cells(posG + 1, 7).Value = Format(CLng(StrConv(Y_NYU_O_REC.J_SURYO, vbUnicode)), "#0")
    Else
        ExcelWorkSheet.Application.Cells(posG + 1, 7).Value = ""
    End If

    Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(Y_NYU_O_REC.TANTO_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(TANTOREC.TANTO_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "担当者ﾏｽﾀ")
            Exit Function
    End Select
    ExcelWorkSheet.Application.Cells(posG + 1, 8).Value = Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))

    ExcelWorkSheet.Application.Cells(posG + 1, 9).Value = Left(StrConv(Y_NYU_O_REC.ORDER_NO, vbUnicode), 5)

    If Trim(StrConv(Y_NYU_O_REC.HIN_NO, vbUnicode)) = "" Then
        ExcelWorkSheet.Application.Cells(posG + 1, 10).Value = ""
    Else
        ExcelWorkSheet.Application.Cells(posG + 1, 10).Value = "*" & Trim(StrConv(Y_NYU_O_REC.HIN_NO, vbUnicode)) & "*"
    End If

    ExcelWorkSheet.Application.Cells(posG + 2, 9).Value = Right(StrConv(Y_NYU_O_REC.ORDER_NO, vbUnicode), 5)



    posG = posG + 3

    EXCEL_Set_Proc = False

End Function


Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   入荷予定ﾃﾞｰﾀ更新
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
    
    
Dim i           As Integer
    
Dim ans         As Integer
    
    
    Update_Proc = True
    F1020251.MousePointer = vbHourglass


                                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        Exit Function
    End If


    On Error GoTo Abort_Tran


    com = BtOpGetFirst
    Do
    
        DoEvents
        
        Do
            sts = BTRV(com + BtSNoWait, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
                
                
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrEOF
                    Exit Do
                
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                
                    ans = MsgBox("他端末でデータ使用中です。<Y_NYU_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        GoTo Abort_Tran
                    End If
                
                Case Else
                    Call File_Error(sts, com, "入荷予定")
                    GoTo Abort_Tran
            End Select
        
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
                
        
        Do
            sts = BTRV(BtOpDelete, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
            
            Select Case sts
                Case BtNoErr
                    Exit Do
                
                Case BtErrKeyNotFound
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    ans = MsgBox("他端末でデータ使用中です。<Y_NYU_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        GoTo Abort_Tran
                    End If
            
                Case Else
                    Call File_Error(sts, BtOpDelete, "入荷予定")
                    GoTo Abort_Tran
            End Select
        Loop
    
    
        com = BtOpGetNext
    
    Loop
    
    
    
    
    
    
    
    
    Set TDBGrid1.Array = NYUKA
    TDBGrid1.Refresh
    TDBGrid1.Update
    
    
    
    
    
    For i = 1 To NYUKA.UpperBound(1)
    
    
        
        Call UniCode_Conv(Y_NYU_O_REC.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(Y_NYU_O_REC.SOKO_NO, DEF_SOKO_NO)
        Call UniCode_Conv(Y_NYU_O_REC.SEQ_NO, Format(i, "000"))
        Call UniCode_Conv(Y_NYU_O_REC.NYUKO_YMD, Text1(ptxNYUKO_YY).Text & Text1(ptxNYUKO_MM).Text & Text1(ptxNYUKO_DD).Text)
        Call UniCode_Conv(Y_NYU_O_REC.DEN_NO, NYUKA(i, colDEN_NO))
        Call UniCode_Conv(Y_NYU_O_REC.MAKER_CODE, NYUKA(i, colMAKER_CODE))
        Call UniCode_Conv(Y_NYU_O_REC.HIN_NO, NYUKA(i, colHIN_NO))
        
        If IsNumeric(NYUKA(i, colY_SURYO)) Then
            Call UniCode_Conv(Y_NYU_O_REC.Y_SURYO, Format(Val(NYUKA(i, colY_SURYO)), "00000000"))
        Else
            Call UniCode_Conv(Y_NYU_O_REC.Y_SURYO, "00000000")
        End If
        If IsNumeric(NYUKA(i, colJ_SURYO)) Then
            Call UniCode_Conv(Y_NYU_O_REC.J_SURYO, Format(Val(NYUKA(i, colJ_SURYO)), "00000000"))
        Else
            Call UniCode_Conv(Y_NYU_O_REC.J_SURYO, "00000000")
        End If
    
        Call UniCode_Conv(Y_NYU_O_REC.TANTO_CODE, NYUKA(i, colTANTO_CODE))
        Call UniCode_Conv(Y_NYU_O_REC.ORDER_NO, NYUKA(i, colORDER_NO))
        Call UniCode_Conv(Y_NYU_O_REC.KENPIN_F, NYUKA(i, colKENPIN_F))
    
        Call UniCode_Conv(Y_NYU_O_REC.WEL_ID, "")
        Call UniCode_Conv(Y_NYU_O_REC.PRG_ID, "")
    
        Call UniCode_Conv(Y_NYU_O_REC.FILLER, "")
    
    
        sts = BTRV(BtOpInsert, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
        If sts <> BtNoErr Then
            Call File_Error(sts, BtOpInsert, "入荷予定")
            GoTo Abort_Tran
        End If
    
    
                                    '品目マスタチェック
        If Item_Check_Proc(In_Mode, Last_JGYOBU, NAIGAI_NAI, NYUKA(i, colHIN_NO), , , NYUKA(i, colMAKER_CODE), NYUKA(i, colMAKER_NAME)) Then
            GoTo Abort_Tran
        End If
    
    
    Next i
    
                                'トランザクション終了
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
    
    F1020251.MousePointer = vbDefault

    Update_Proc = False


    Exit Function


Abort_Tran:
    
    On Error GoTo 0
    
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


End Function



Private Function NEW_Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   入荷予定ﾃﾞｰﾀ更新
'       2017.01.27
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
    
    
Dim i           As Integer
    
Dim ans         As Integer
    
    
    NEW_Update_Proc = True
    F1020251.MousePointer = vbHourglass




    
    
    
    
    
    
    
    
    Set TDBGrid1.Array = NYUKA
    TDBGrid1.Refresh
    TDBGrid1.Update
    
    
    
    
    
    For i = 1 To NYUKA.UpperBound(1)
    
    
    
    
        Do
            Call UniCode_Conv(K0_Y_NYU_O.SEQ_NO, Format(i, "000"))
            sts = BTRV(BtOpGetEqual, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound
                    com = BtOpInsert
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    ans = MsgBox("他端末でデータ使用中です。<Y_NYU_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        GoTo Abort_Tran
                    End If
            
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "入荷予定")
                    GoTo Abort_Tran
            End Select
        Loop
        
        Call UniCode_Conv(Y_NYU_O_REC.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(Y_NYU_O_REC.SOKO_NO, DEF_SOKO_NO)
        Call UniCode_Conv(Y_NYU_O_REC.SEQ_NO, Format(i, "000"))
        Call UniCode_Conv(Y_NYU_O_REC.NYUKO_YMD, Text1(ptxNYUKO_YY).Text & Text1(ptxNYUKO_MM).Text & Text1(ptxNYUKO_DD).Text)
        Call UniCode_Conv(Y_NYU_O_REC.DEN_NO, NYUKA(i, colDEN_NO))
        Call UniCode_Conv(Y_NYU_O_REC.MAKER_CODE, NYUKA(i, colMAKER_CODE))
        Call UniCode_Conv(Y_NYU_O_REC.HIN_NO, NYUKA(i, colHIN_NO))
        
        If IsNumeric(NYUKA(i, colY_SURYO)) Then
            Call UniCode_Conv(Y_NYU_O_REC.Y_SURYO, Format(Val(NYUKA(i, colY_SURYO)), "00000000"))
        Else
            Call UniCode_Conv(Y_NYU_O_REC.Y_SURYO, "00000000")
        End If
        If IsNumeric(NYUKA(i, colJ_SURYO)) Then
            Call UniCode_Conv(Y_NYU_O_REC.J_SURYO, Format(Val(NYUKA(i, colJ_SURYO)), "00000000"))
        Else
            Call UniCode_Conv(Y_NYU_O_REC.J_SURYO, "00000000")
        End If
    
        Call UniCode_Conv(Y_NYU_O_REC.TANTO_CODE, NYUKA(i, colTANTO_CODE))
        Call UniCode_Conv(Y_NYU_O_REC.ORDER_NO, NYUKA(i, colORDER_NO))
        Call UniCode_Conv(Y_NYU_O_REC.KENPIN_F, NYUKA(i, colKENPIN_F))
    
        Call UniCode_Conv(Y_NYU_O_REC.WEL_ID, "")
        Call UniCode_Conv(Y_NYU_O_REC.PRG_ID, "")
    
        Call UniCode_Conv(Y_NYU_O_REC.FILLER, "")
    
    
        sts = BTRV(com, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                ans = MsgBox("他端末でデータ使用中です。<Y_NYU_O.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
        
            Case Else
                Call File_Error(sts, com, "入荷予定")
                GoTo Abort_Tran
        End Select
    
    
                                    '品目マスタチェック
        If Item_Check_Proc(In_Mode, Last_JGYOBU, NAIGAI_NAI, NYUKA(i, colHIN_NO), , , NYUKA(i, colMAKER_CODE), NYUKA(i, colMAKER_NAME)) Then
            GoTo Abort_Tran
        End If
    
    
    Next i
    
    
    
    
    F1020251.MousePointer = vbDefault

    NEW_Update_Proc = False


    Exit Function


Abort_Tran:
    


End Function

