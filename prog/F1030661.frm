VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form F1030661 
   BackColor       =   &H00FFFFFF&
   Caption         =   "出荷問い合わせ(問合せ��) 検品取り消し機能付き"
   ClientHeight    =   8715
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   15240
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
   ScaleHeight     =   8715
   ScaleWidth      =   15240
   StartUpPosition =   2  '画面の中央
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
      Left            =   10320
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   8040
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   8040
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
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8040
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
      Index           =   8
      Left            =   7800
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8040
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   8040
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   8040
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   8040
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
      Left            =   3960
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   8040
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   8040
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8040
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   8040
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
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   8040
      Width           =   855
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   5655
      Left            =   105
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1920
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   9975
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "問合せ��"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "��"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "出荷日"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "送り先名"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "売伝"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "伝票番号"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "品番／品名"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "数量"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "注文��"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "得意先"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "備考"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "運送会社"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "口"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "検品日時"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "検品担当者"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "ID_NO"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "仕掛中　品番"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "仕掛数　バラ"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "仕掛数　箱"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "品番読込み実績"
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   21
      Splits(0)._UserFlags=   0
      Splits(0).Locked=   -1  'True
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=21"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2196"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2090"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=900"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=794"
      Splits(0)._ColumnProps(8)=   "Column(1)._ColStyle=2"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=1958"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1852"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=2143"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2037"
      Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(18)=   "Column(4).Width=900"
      Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=794"
      Splits(0)._ColumnProps(21)=   "Column(4)._ColStyle=1"
      Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(23)=   "Column(5).Width=1931"
      Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=1826"
      Splits(0)._ColumnProps(26)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(27)=   "Column(6).Width=2831"
      Splits(0)._ColumnProps(28)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(6)._WidthInPix=2725"
      Splits(0)._ColumnProps(30)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(31)=   "Column(7).Width=1191"
      Splits(0)._ColumnProps(32)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(7)._WidthInPix=1085"
      Splits(0)._ColumnProps(34)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(35)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(36)=   "Column(8).Width=1640"
      Splits(0)._ColumnProps(37)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(8)._WidthInPix=1535"
      Splits(0)._ColumnProps(39)=   "Column(8)._ColStyle=0"
      Splits(0)._ColumnProps(40)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(41)=   "Column(9).Width=2408"
      Splits(0)._ColumnProps(42)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(9)._WidthInPix=2302"
      Splits(0)._ColumnProps(44)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(45)=   "Column(10).Width=1931"
      Splits(0)._ColumnProps(46)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(10)._WidthInPix=1826"
      Splits(0)._ColumnProps(48)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(49)=   "Column(11).Width=1667"
      Splits(0)._ColumnProps(50)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(11)._WidthInPix=1561"
      Splits(0)._ColumnProps(52)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(53)=   "Column(12).Width=582"
      Splits(0)._ColumnProps(54)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(12)._WidthInPix=476"
      Splits(0)._ColumnProps(56)=   "Column(12)._ColStyle=2"
      Splits(0)._ColumnProps(57)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(58)=   "Column(13).Width=1667"
      Splits(0)._ColumnProps(59)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(60)=   "Column(13)._WidthInPix=1561"
      Splits(0)._ColumnProps(61)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(62)=   "Column(14).Width=2461"
      Splits(0)._ColumnProps(63)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(14)._WidthInPix=2355"
      Splits(0)._ColumnProps(65)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(66)=   "Column(15).Width=2514"
      Splits(0)._ColumnProps(67)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(15)._WidthInPix=2408"
      Splits(0)._ColumnProps(69)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(70)=   "Column(16).Width=3810"
      Splits(0)._ColumnProps(71)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(72)=   "Column(16)._WidthInPix=3704"
      Splits(0)._ColumnProps(73)=   "Column(16).Visible=0"
      Splits(0)._ColumnProps(74)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(75)=   "Column(17).Width=3810"
      Splits(0)._ColumnProps(76)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(77)=   "Column(17)._WidthInPix=3704"
      Splits(0)._ColumnProps(78)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(79)=   "Column(18).Width=3810"
      Splits(0)._ColumnProps(80)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(81)=   "Column(18)._WidthInPix=3704"
      Splits(0)._ColumnProps(82)=   "Column(18)._ColStyle=2"
      Splits(0)._ColumnProps(83)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(84)=   "Column(19).Width=3810"
      Splits(0)._ColumnProps(85)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(86)=   "Column(19)._WidthInPix=3704"
      Splits(0)._ColumnProps(87)=   "Column(19)._ColStyle=2"
      Splits(0)._ColumnProps(88)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(89)=   "Column(20).Width=3810"
      Splits(0)._ColumnProps(90)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(20)._WidthInPix=3704"
      Splits(0)._ColumnProps(92)=   "Column(20)._ColStyle=2"
      Splits(0)._ColumnProps(93)=   "Column(20).Order=21"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=10.5,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=10.5,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=14,.parent=53"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=11,.parent=54"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=12,.parent=55"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=13,.parent=57"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=18,.parent=53,.alignment=1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=15,.parent=54"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=16,.parent=55"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=17,.parent=57"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=48,.parent=53"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=45,.parent=54"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=46,.parent=55"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=47,.parent=57"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=66,.parent=53"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=54"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=55"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=57"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=102,.parent=53,.alignment=2"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=19,.parent=54"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=20,.parent=55"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=101,.parent=57"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=70,.parent=53"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=54"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=55"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=57"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=78,.parent=53"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=75,.parent=54"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=76,.parent=55"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=77,.parent=57"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=86,.parent=53,.alignment=1,.locked=0"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=83,.parent=54,.alignment=3"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=84,.parent=55,.alignment=3"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=85,.parent=57"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=24,.parent=53,.alignment=0,.locked=0"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=21,.parent=54,.alignment=3"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=22,.parent=55,.alignment=3"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=23,.parent=57"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=28,.parent=53"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=25,.parent=54"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=26,.parent=55"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=27,.parent=57"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=44,.parent=53"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=41,.parent=54"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=42,.parent=55"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=43,.parent=57"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=74,.parent=53"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=71,.parent=54"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=72,.parent=55"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=73,.parent=57"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=40,.parent=53,.alignment=1"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=37,.parent=54"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=38,.parent=55"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=39,.parent=57"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=52,.parent=53"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=49,.parent=54"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=50,.parent=55"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=51,.parent=57"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=96,.parent=53"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=93,.parent=54"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=94,.parent=55"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=95,.parent=57"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=82,.parent=53"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=79,.parent=54"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=80,.parent=55"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=81,.parent=57"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=118,.parent=53"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=115,.parent=54"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=116,.parent=55"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=117,.parent=57"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=100,.parent=53"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=97,.parent=54"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=98,.parent=55"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=99,.parent=57"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=106,.parent=53,.alignment=1"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=103,.parent=54"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=104,.parent=55"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=105,.parent=57"
      _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=110,.parent=53,.alignment=1"
      _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=107,.parent=54"
      _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=108,.parent=55"
      _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=109,.parent=57"
      _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=114,.parent=53,.alignment=1"
      _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=111,.parent=54"
      _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=112,.parent=55"
      _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=113,.parent=57"
      _StyleDefs(120) =   "Named:id=29:Normal"
      _StyleDefs(121) =   ":id=29,.parent=0"
      _StyleDefs(122) =   "Named:id=30:Heading"
      _StyleDefs(123) =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(124) =   ":id=30,.wraptext=-1"
      _StyleDefs(125) =   "Named:id=31:Footing"
      _StyleDefs(126) =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(127) =   "Named:id=32:Selected"
      _StyleDefs(128) =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(129) =   "Named:id=33:Caption"
      _StyleDefs(130) =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(131) =   "Named:id=34:HighlightRow"
      _StyleDefs(132) =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(133) =   "Named:id=35:EvenRow"
      _StyleDefs(134) =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(135) =   "Named:id=36:OddRow"
      _StyleDefs(136) =   ":id=36,.parent=29"
      _StyleDefs(137) =   "Named:id=89:RecordSelector"
      _StyleDefs(138) =   ":id=89,.parent=30"
      _StyleDefs(139) =   "Named:id=92:FilterBar"
      _StyleDefs(140) =   ":id=92,.parent=29"
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   105
      TabIndex        =   26
      Top             =   120
      Width           =   14925
      Begin VB.OptionButton Option1 
         Height          =   255
         Index           =   1
         Left            =   4515
         TabIndex        =   36
         Top             =   360
         Width           =   225
      End
      Begin VB.TextBox Text 
         Height          =   360
         Index           =   1
         Left            =   5670
         MaxLength       =   7
         TabIndex        =   35
         Top             =   240
         Width           =   930
      End
      Begin VB.CheckBox Check1 
         Caption         =   "検品済"
         Height          =   255
         Index           =   1
         Left            =   2835
         TabIndex        =   11
         Top             =   1440
         Width           =   1065
      End
      Begin VB.CheckBox Check1 
         Caption         =   "未検品"
         Height          =   255
         Index           =   0
         Left            =   1575
         TabIndex        =   10
         Top             =   1440
         Width           =   1065
      End
      Begin VB.ComboBox Combo 
         Height          =   360
         Index           =   1
         Left            =   12810
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   9
         Top             =   840
         Width           =   1965
      End
      Begin VB.TextBox Text 
         Height          =   360
         Index           =   6
         Left            =   9345
         MaxLength       =   20
         TabIndex        =   8
         Top             =   840
         Width           =   2505
      End
      Begin VB.ComboBox Combo 
         Height          =   360
         Index           =   0
         Left            =   5565
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   7
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox Text 
         Height          =   360
         Index           =   5
         Left            =   4515
         MaxLength       =   8
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text 
         Height          =   360
         Index           =   2
         Left            =   1470
         MaxLength       =   4
         TabIndex        =   3
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text 
         Height          =   360
         Index           =   3
         Left            =   2415
         MaxLength       =   2
         TabIndex        =   4
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text 
         Height          =   360
         Index           =   4
         Left            =   3060
         MaxLength       =   2
         TabIndex        =   5
         Top             =   840
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Height          =   255
         Index           =   2
         Left            =   105
         TabIndex        =   2
         Top             =   960
         Width           =   225
      End
      Begin VB.TextBox Text 
         Height          =   360
         Index           =   0
         Left            =   1470
         MaxLength       =   20
         TabIndex        =   1
         Top             =   240
         Width           =   2505
      End
      Begin VB.OptionButton Option1 
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   27
         Top             =   360
         Width           =   225
      End
      Begin VB.Line Line2 
         X1              =   4305
         X2              =   4305
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "伝票��"
         Height          =   240
         Index           =   8
         Left            =   4830
         TabIndex        =   37
         Top             =   360
         Width           =   720
      End
      Begin VB.Line Line1 
         X1              =   105
         X2              =   14700
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "運送会社"
         Height          =   240
         Index           =   3
         Left            =   11865
         TabIndex        =   34
         Top             =   960
         Width           =   960
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "品番"
         Height          =   240
         Index           =   2
         Left            =   8715
         TabIndex        =   33
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "得意先"
         Height          =   240
         Index           =   0
         Left            =   3780
         TabIndex        =   32
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "出 荷 日"
         Height          =   240
         Index           =   4
         Left            =   420
         TabIndex        =   31
         Top             =   960
         Width           =   960
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "年"
         Height          =   240
         Index           =   5
         Left            =   2100
         TabIndex        =   30
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "月"
         Height          =   240
         Index           =   6
         Left            =   2820
         TabIndex        =   29
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "日"
         Height          =   240
         Index           =   7
         Left            =   3420
         TabIndex        =   28
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "問合せ��"
         Height          =   240
         Index           =   1
         Left            =   420
         TabIndex        =   0
         Top             =   375
         Width           =   960
      End
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
      Left            =   120
      TabIndex        =   25
      Top             =   6600
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
Attribute VB_Name = "F1030661"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const poptSEL_OKURI_NO% = 0     '問合せ�ｑI択
Private Const poptSEL_DEN_NO% = 1       '伝票�ｑI択
Private Const poptSEL_ETC% = 2          'その他選択



Private Const ptxSEL_OKURI_NO% = 0      '問合せ��
Private Const ptxSEL_DEN_NO% = 1        '伝票��
Private Const ptxSEL_SYUKA_YY% = 2      '出荷日(年)
Private Const ptxSEL_SYUKA_MM% = 3      '出荷日(年)
Private Const ptxSEL_SYUKA_DD% = 4      '出荷日(年)

Private Const ptxSEL_MUKE_CODE% = 5     '出荷先
Private Const ptxSEL_HIN_NO% = 6        '品番

Private Const pcmbSEL_MUKE_NAME% = 0    '出荷先名
Private Const pcmbSEL_UNSOU_KAISHA% = 1 '運送会社

Private Const pchkMI% = 0               '未処理
Private Const pchkSUMI% = 1             '処理済

Private Const UNSOU_KAISHA_ALL = "全  て"


Dim SYUKA As New XArrayDB

Private Const Min_Row% = 1              '最小行数
Dim Max_Row    As Integer               'グリッド最大表示件数



Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 20             '最大列数

Private Const ColOKURI_NO% = 0              '問合せ��
Private Const ColSYUKA_NO% = 1              '��
Private Const ColSYUKA_YMD% = 2             '出荷日

Private Const ColOKURISAKI% = 3             '送り先
Private Const ColURIDEN% = 4                '売伝
Private Const ColDEN_NO% = 5                '伝票番号
Private Const ColHIN_NO% = 6                '品番／品名
Private Const ColSURYO% = 7                 '数量
Private Const ColORDER_NO% = 8              '注文��
Private Const ColMUKE_CODE% = 9             '得意先
Private Const ColBIKOU% = 10                '備考
Private Const ColUNSOU_KAISHA% = 11         '運送会社
Private Const ColKUTI_SU% = 12              '口数
Private Const ColKENPIN_NOW% = 13           '検品日時
Private Const ColKENPIN_TANTO_CODE% = 14    '検品担当者
Private Const ColID_NO% = 15                'ID_NO

Private Const ColKEY_HIN_NO% = 16           '品番のみ

Private Const ColKEN_HINBAN% = 17           '仕掛中　品番       2012.10.24
Private Const ColCNT_BARA_SU% = 18          '仕掛中　バラ       2012.10.24
Private Const ColCNT_HAKO_SU% = 19          '仕掛中  箱         2012.10.24
Private Const ColJ_HIN_CHK_CNT% = 20        '品番読込み実績     2012.10.24


Private Sort_Tbl(Min_Col To Max_Col) _
                As Integer                  'ｿｰﾄの制御 0:昇順 1:降順
                
Private Const LAST_UPDATE_DAY$ = "[F103066] 2012.10.25 16:00 検品取り消し機能付き 注文ﾃﾞｰﾀ対応"



Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    

    Call Tab_Ctrl(Shift)
End Sub


Private Sub Command_Click(Index As Integer)

Dim ans As Integer

    Select Case Index
        Case 7
            
            If Option1(poptSEL_OKURI_NO).Value Then
                If Trim(Text(ptxSEL_OKURI_NO).Text) = "" Then
                    MsgBox "問合せ�ｂ�指定して下さい。"
                    Exit Sub
                End If
            
            
            Else
                If IsNumeric(Trim(Text(ptxSEL_SYUKA_MM).Text)) Then
                    Text(ptxSEL_SYUKA_MM).Text = Format(CInt(Trim(Text(ptxSEL_SYUKA_MM).Text)), "00")
                End If
                If IsNumeric(Trim(Text(ptxSEL_SYUKA_DD).Text)) Then
                    Text(ptxSEL_SYUKA_DD).Text = Format(CInt(Trim(Text(ptxSEL_SYUKA_DD).Text)), "00")
                End If
                Text(ptxSEL_MUKE_CODE).Text = Right(Combo(pcmbSEL_MUKE_NAME), 8)
            End If
            
            
            If Option1(poptSEL_OKURI_NO).Value Or Option1(poptSEL_ETC).Value Then
                If List_Disp_Proc Then
                    Unload Me
                End If
        
            Else
        
                If List_Disp_DEN_Proc Then
                    Unload Me
                End If

            End If
        
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


Dim UNSOU_KAISHA    As Variant

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
    Max_Row = CInt(RTrim(c))
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
            F1030661.Caption = "出荷問合わせ(問合せ��)　検品取り消し機能付き（" + RTrim(JGYOBU_T(i).NAME) + ") " & LAST_UPDATE_DAY
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)

                                '向け先管理マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)ＯＰＥＮ
    If Y_SYU_H_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定ＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '邸別注文ﾃﾞｰﾀＯＰＥＮ
    If Y_SYU_TEI_Open(BtOpenNomal) Then
        Unload Me
    End If




    Option1(poptSEL_OKURI_NO).Value = True
    Option1(poptSEL_ETC).Value = False

'出荷日付
    Text(ptxSEL_SYUKA_YY).Text = Left(Format(Now, "YYYYMMDD"), 4)
    Text(ptxSEL_SYUKA_MM).Text = Mid(Format(Now, "YYYYMMDD"), 5, 2)
    Text(ptxSEL_SYUKA_DD).Text = Right(Format(Now, "YYYYMMDD"), 2)

'向け先設定
    If MTS_Set_Proc() Then
        Unload Me
    End If


'運送会社
    Combo(pcmbSEL_UNSOU_KAISHA).Clear
    Combo(pcmbSEL_UNSOU_KAISHA).AddItem UNSOU_KAISHA_ALL
                    
                    
                                '運送会社名称獲得
    If GetIni(App.EXEName, "LOGF", "UNSOU_KAISHA", c) Then
    Else
        UNSOU_KAISHA = Split(Trim(c), ",", -1)
        For i = 0 To UBound(UNSOU_KAISHA)
            Combo(pcmbSEL_UNSOU_KAISHA).AddItem UNSOU_KAISHA(i)
        Next i
    End If
    Combo(pcmbSEL_UNSOU_KAISHA).ListIndex = 0
    
'表示条件
    Check1(pchkMI).Value = vbChecked
    Check1(pchkSUMI).Value = vbChecked
    
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0                 'ﾃﾞﾌｫﾙﾄ昇順
    Next i
    

    Text(ptxSEL_OKURI_NO).SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
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
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
                                            '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
        End If
    End If
    
    sts = BTRV(BtOpReset, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    Set F1030661 = Nothing


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
    F1030661.Caption = "出荷問い合わせ(問合せ��)　検品取り消し機能付き（" + RTrim(JGYOBU_T(Index).NAME) + ") " & LAST_UPDATE_DAY
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
    
    
    Combo(pcmbSEL_MUKE_NAME).Clear
    
    Combo(pcmbSEL_MUKE_NAME).AddItem "全て　　　" & "   " & Space(8)
        
    
    
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
        Edit = Edit & StrConv(MTSREC.MUKE_CODE, vbUnicode)
        
        
        Combo(pcmbSEL_MUKE_NAME).AddItem Edit
    
        com = BtOpGetNext
    
    Loop

    Combo(pcmbSEL_MUKE_NAME).ListIndex = 0

    Call Input_UnLock

    MTS_Set_Proc = False
End Function


Private Function List_Disp_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim Row         As Long
    
Dim DEN_MAISU   As Long
Dim KAN_MAISU   As Long
    
Dim Skip_Flg    As Boolean
    
    
    List_Disp_Proc = True
                                    
    Call Input_Lock
                                    
                                    
                                    'テーブルリセット
    Set SYUKA = Nothing
                                    
'---------------------------------- '当日分出荷予定読み込み開始
    If Option1(poptSEL_OKURI_NO).Value Then
    
        Call UniCode_Conv(K2_Y_SYU_H.OKURI_NO, Text(ptxSEL_OKURI_NO).Text)
    
    Else
    
        Call UniCode_Conv(K3_Y_SYU_H.SYUKA_YMD, Text(ptxSEL_SYUKA_YY).Text & _
                                                Text(ptxSEL_SYUKA_MM).Text & _
                                                Text(ptxSEL_SYUKA_DD).Text)
    
    End If
    
    
    
    
    
    
    Row = Min_Row - 1
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        Skip_Flg = False
        If Option1(poptSEL_OKURI_NO).Value Then
                
        
            sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K2_Y_SYU_H, Len(K2_Y_SYU_H), 2)
    
            Select Case sts
                Case BtNoErr
            
                    If Trim(StrConv(Y_SYU_HREC.OKURI_NO, vbUnicode)) <> Trim(Text(ptxSEL_OKURI_NO).Text) Then
                        sts = BtErrEOF
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                    List_Disp_Proc = SYS_ERR
                    Exit Function
            End Select
    
    
    
        Else
    
            sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K3_Y_SYU_H, Len(K3_Y_SYU_H), 3)
    
            Select Case sts
                Case BtNoErr
                        
                    If Trim(Text(ptxSEL_SYUKA_YY).Text) <> "" Then
                        If Text(ptxSEL_SYUKA_YY).Text <> Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 1, 4) Then
                            sts = BtErrEOF
                            Exit Do
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_SYUKA_MM).Text) <> "" Then
                        If Text(ptxSEL_SYUKA_MM).Text <> Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 5, 2) Then
                            sts = BtErrEOF
                            Exit Do
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_SYUKA_DD).Text) <> "" Then
                        If Text(ptxSEL_SYUKA_DD).Text <> Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 7, 2) Then
                            sts = BtErrEOF
                            Exit Do
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_MUKE_CODE).Text) <> "" Then
                        If Trim(Text(ptxSEL_MUKE_CODE).Text) <> Trim(StrConv(Y_SYU_HREC.MUKE_CODE, vbUnicode)) Then
                            Skip_Flg = True
                        End If
                    End If
                
                    If Trim(Text(ptxSEL_HIN_NO).Text) <> "" Then
                        If Trim(Text(ptxSEL_HIN_NO).Text) <> Trim(StrConv(Y_SYU_HREC.HIN_NO, vbUnicode)) Then
                            Skip_Flg = True
                        End If
                    End If
                        
                    If Trim(Combo(pcmbSEL_UNSOU_KAISHA).Text) <> UNSOU_KAISHA_ALL Then
                        If Trim(Combo(pcmbSEL_UNSOU_KAISHA).Text) <> Trim(StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)) Then
                            Skip_Flg = True
                        End If
                    End If
                
                    If Check1(pchkMI).Value <> vbChecked Then
                        If Trim(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode)) = "" Then
                            Skip_Flg = True
                        End If
                    End If
                
                    If Check1(pchkSUMI).Value <> vbChecked Then
                        If Trim(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode)) <> "" Then
                            Skip_Flg = True
                        End If
                    End If
                
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                    List_Disp_Proc = SYS_ERR
                    Exit Function
            End Select
    
    
        End If
                
                
        If Not Skip_Flg Then
            
            Row = Row + 1
            If Row > Max_Row Then
                Beep
                MsgBox "最大表示行数を超えました。"
                Exit Do
            End If
                    
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
        End If
        
        com = BtOpGetNext
        
    Loop
    
    
    
    
    Set TDBGrid1.Array = SYUKA
    
    TDBGrid1.Style.Locked = True
    
    
    
    
    
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    
    
    
    
    Call Input_UnLock
    
    
    
    List_Disp_Proc = False

    
End Function
Private Function List_Disp_DEN_Proc() As Integer
'----------------------------------------------------------------------------
'                   伝票�ｂﾅの検索
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim Row         As Long
    
Dim DEN_MAISU   As Long
Dim KAN_MAISU   As Long
    
Dim Skip_Flg    As Boolean
    
Dim svID_No     As String * 7
    
    
    
    
    List_Disp_DEN_Proc = True
                                    
    Call Input_Lock
                                    
                                    
                                    'テーブルリセット
    Set SYUKA = Nothing
                                    
'---------------------------------- '当日分出荷予定読み込み開始
    
    Call UniCode_Conv(K0_Y_SYU_H.DEN_NO, Text(ptxSEL_DEN_NO).Text)
    Call UniCode_Conv(K0_Y_SYU_H.SEQ_NO, "")
    sts = BTRV(BtOpGetGreaterEqual, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)

    Select Case sts
        Case BtNoErr
    
            If Left(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode), 7) <> Trim(Text(ptxSEL_DEN_NO).Text) Then
                sts = BtErrEOF
            End If
        
        Case BtErrEOF
        
        Case Else
            Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
            List_Disp_DEN_Proc = SYS_ERR
            Exit Function
    
    End Select
    
    
    If sts = BtNoErr Then
        svID_No = StrConv(Y_SYU_HREC.ID_NO, vbUnicode)
        
        Call UniCode_Conv(K4_Y_SYU_H.ID_NO, svID_No)
        
        Row = Min_Row - 1
    
        com = BtOpGetGreaterEqual
    
        Do
        
            DoEvents
        
        
            sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
    
            Select Case sts
                Case BtNoErr
            
                    If Left(StrConv(Y_SYU_HREC.ID_NO, vbUnicode), 7) <> svID_No Then
                        sts = BtErrEOF
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                    List_Disp_DEN_Proc = SYS_ERR
                    Exit Function
            End Select
                
            If Not Skip_Flg Then
                
                Row = Row + 1
                If Row > Max_Row Then
                    Beep
                    MsgBox "最大表示行数を超えました。"
                    Exit Do
                End If
                        
                If Grid_Set_Proc(Row) Then
                    Exit Function
                End If
            End If
        
            com = BtOpGetNext
        
        Loop
    
    End If
    
    
    
    
    Set TDBGrid1.Array = SYUKA
    
    TDBGrid1.Style.Locked = True
    
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    
    
    
    Call Input_UnLock
    
    
    
    List_Disp_DEN_Proc = False

    
End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1030661.MousePointer = vbHourglass

    Call Ctrl_Lock(F1030661)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1030661)


    F1030661.MousePointer = vbDefault

End Sub

Private Function Grid_Set_Proc(Row As Long) As Integer

Dim sts As Integer

    
    Grid_Set_Proc = True

    

    SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
    
    '問合せ��
    SYUKA(Row, ColOKURI_NO) = Trim(StrConv(Y_SYU_HREC.OKURI_NO, vbUnicode))
    '出荷��
    SYUKA(Row, ColSYUKA_NO) = Trim(StrConv(Y_SYU_HREC.SYUKA_NO, vbUnicode))
    '出荷日
    SYUKA(Row, ColSYUKA_YMD) = Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode), 7, 2)
    '送り先
    SYUKA(Row, ColOKURISAKI) = Trim(StrConv(Y_SYU_HREC.OKURISAKI, vbUnicode))
    '売り伝
    If StrConv(Y_SYU_HREC.URIDEN, vbUnicode) = "1" Then
        SYUKA(Row, ColURIDEN) = "有"
    Else
        SYUKA(Row, ColURIDEN) = ""
    End If
    '伝票番号
    SYUKA(Row, ColDEN_NO) = Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode))
    '品番＆品名
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYU_HREC.HIN_NO, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
            Exit Function
    End Select
    SYUKA(Row, ColHIN_NO) = Left(StrConv(Y_SYU_HREC.HIN_NO, vbUnicode), 12) & " " & Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
    '数量
    SYUKA(Row, ColSURYO) = Format(CLng(StrConv(Y_SYU_HREC.SURYO, vbUnicode)), "#0")
    '注文��
    SYUKA(Row, ColORDER_NO) = StrConv(Y_SYU_HREC.ODER_NO, vbUnicode)
    '得意先ｺｰﾄﾞ＆名称
    SYUKA(Row, ColMUKE_CODE) = StrConv(Y_SYU_HREC.MUKE_CODE, vbUnicode) & " " & Trim(StrConv(Y_SYU_HREC.MUKE_NAME, vbUnicode))
    '備考
    SYUKA(Row, ColBIKOU) = StrConv(Y_SYU_HREC.BIKOU, vbUnicode)
    '運送会社
    SYUKA(Row, ColUNSOU_KAISHA) = StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)
    '口数
    If IsNumeric(StrConv(Y_SYU_HREC.KUTI_SU, vbUnicode)) Then
        SYUKA(Row, ColKUTI_SU) = Format(CInt(StrConv(Y_SYU_HREC.KUTI_SU, vbUnicode)), "#")
    Else
        SYUKA(Row, ColKUTI_SU) = ""
    End If
    '検品日時
    If Trim(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode)) = "" Then
        SYUKA(Row, ColKENPIN_NOW) = ""
    Else
        SYUKA(Row, ColKENPIN_NOW) = Mid(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode), 7, 2) & " " & _
                                    Mid(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode), 9, 2) & ":" & _
                                    Mid(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode), 11, 2) & ":" & _
                                    Mid(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode), 13, 2)


    End If
    '担当者
    If Trim(StrConv(Y_SYU_HREC.KENPIN_TANTO_CODE, vbUnicode)) = "" Then
        SYUKA(Row, ColKENPIN_TANTO_CODE) = ""
    Else
        Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(Y_SYU_HREC.KENPIN_TANTO_CODE, vbUnicode))
        sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    
        Select Case sts
            Case BtNoErr
            
            Case BtErrKeyNotFound
                Call UniCode_Conv(TANTOREC.TANTO_NAME, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "担当者ﾏｽﾀ")
                Exit Function
        End Select
    
        SYUKA(Row, ColKENPIN_TANTO_CODE) = Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
    
    End If
    
    'ID_NO
    SYUKA(Row, ColID_NO) = Trim(StrConv(Y_SYU_HREC.ID_NO, vbUnicode))
    '品番
    SYUKA(Row, ColKEY_HIN_NO) = Trim(StrConv(Y_SYU_HREC.HIN_NO, vbUnicode))
    
    '仕掛中　品番   2012.10.24
    SYUKA(Row, ColKEN_HINBAN) = Trim(StrConv(Y_SYU_HREC.KEN_HINBAN, vbUnicode))
    '仕掛中　バラ   2012.10.24
    SYUKA(Row, ColCNT_BARA_SU) = Trim(StrConv(Y_SYU_HREC.CNT_BARA_SU, vbUnicode))
    '仕掛中　箱   2012.10.24
    SYUKA(Row, ColCNT_HAKO_SU) = Trim(StrConv(Y_SYU_HREC.CNT_HAKO_SU, vbUnicode))
    '品番読込み回数   2012.10.24
    SYUKA(Row, ColJ_HIN_CHK_CNT) = Trim(StrConv(Y_SYU_HREC.J_HIN_CHK_CNT, vbUnicode))
    
    Grid_Set_Proc = False
End Function

Private Sub TDBGrid1_DblClick()
    
Dim yn  As Integer
    
    
    If TDBGrid1.Bookmark = -1 Then
    Else
    
        yn = MsgBox("検品実績の取り消しを行いますか？(該当するＩＤ全てが取り消されます。)", vbYesNo + vbDefaultButton2, "確認入力")
        
        
        Select Case yn
        
    
            Case vbYes
    
                If KENPIN_Update_Proc() Then
                    Unload Me
                End If
        
                
                If Option1(poptSEL_OKURI_NO).Value Or Option1(poptSEL_ETC).Value Then
                
                
                    If List_Disp_Proc() Then
                        Unload Me
                    End If
    
    
                Else
    
                    If List_Disp_DEN_Proc() Then
                        Unload Me
                    End If
    
                End If
    
        End Select
    End If

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
        
        Case ptxSEL_SYUKA_YY
            If Len(Trim(Text(ptxSEL_SYUKA_YY).Text)) = 0 Then
            Else
                If Not IsNumeric(Text(ptxSEL_SYUKA_YY).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Exit Sub
                End If
            End If
        Case ptxSEL_SYUKA_MM
            If Len(Trim(Text(ptxSEL_SYUKA_MM).Text)) = 0 Then
            Else
                If Not IsNumeric(Text(ptxSEL_SYUKA_MM).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Exit Sub
                End If
                Text(ptxSEL_SYUKA_MM).Text = Format(CInt(Text(ptxSEL_SYUKA_MM).Text), "00")
            End If
        Case ptxSEL_SYUKA_DD
            If Len(Trim(Text(ptxSEL_SYUKA_DD).Text)) = 0 Then
            Else
                If Not IsNumeric(Text(ptxSEL_SYUKA_DD).Text) Then
                    Beep
                    MsgBox "入力した項目はエラーです。"
                    Exit Sub
                End If
                Text(ptxSEL_SYUKA_DD).Text = Format(CInt(Text(ptxSEL_SYUKA_DD).Text), "00")
            End If
        
        
        
        Case ptxSEL_MUKE_CODE

            For i = 0 To Combo(pcmbSEL_MUKE_NAME).ListCount - 1 '向け先
    
                If Right(Combo(pcmbSEL_MUKE_NAME).List(i), 8) = Trim(Text(ptxSEL_MUKE_CODE)) Then
                    Combo(pcmbSEL_MUKE_NAME).ListIndex = i
                    Exit For
                End If
            
    
            Next
        
            If i > Combo(pcmbSEL_MUKE_NAME).ListCount - 1 Then
                Beep
                MsgBox "入力した項目はエラーです。"
                Combo(pcmbSEL_MUKE_NAME).ListIndex = -1
                Exit Sub
            End If
            
            Combo(pcmbSEL_MUKE_NAME).SetFocus
    End Select
    
    Call Tab_Ctrl(Shift)

End Sub

Private Function KENPIN_Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   検品済取り消し更新
'----------------------------------------------------------------------------
Dim sts As Integer
Dim ans As Integer
Dim com As Integer

    If TDBGrid1.Bookmark = -1 Then
        Exit Function
    End If
    
    
    KENPIN_Update_Proc = True
                                     'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
    '------------------------------- 出荷予定の処理
    Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU)     '事業部
    Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Left(SYUKA(TDBGrid1.Bookmark, ColID_NO), 7))  ' ID��
    
    com = BtOpGetGreaterEqual
    Do
    
        DoEvents
        
        Do
        
            sts = BTRV(com + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            Select Case sts
                Case BtNoErr
                    
                    If Left(SYUKA(TDBGrid1.Bookmark, ColID_NO), 7) <> Left(StrConv(Y_SYUREC.ID_NO, vbUnicode), 7) Then
                        sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                        If sts <> BtNoErr Then
                            Call File_Error(sts, BtOpUnlock, "出荷予定")
                            GoTo Abort_Tran
                        
                        End If
                        
                        sts = BtErrEOF
                    End If
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                
                    ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        KENPIN_Update_Proc = False
                        GoTo Abort_Tran
                    End If
                
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定")
                    GoTo Abort_Tran
            End Select
    
        Loop
        
        If sts = BtErrEOF Then
            Exit Do
        End If
        
'        If Trim(SYUKA(TDBGrid1.Bookmark, ColKEY_HIN_NO)) <> Trim(StrConv(Y_SYUREC.HIN_NO, vbUnicode)) Then
'        Else
        
        
'2012.10.03            If Trim(StrConv(Y_SYUREC.WEL_ID, vbUnicode)) <> "" Then
'2012.10.03                MsgBox "他端末で処理中です。当画面では処理できません。。"
'2012.10.03                KENPIN_Update_Proc = False
'2012.10.03                GoTo Abort_Tran
'2012.10.03            End If
        
        
        
                                        
        
        
        
            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
            
            '未検品する
            Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
            Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, "")
            Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, "")
            Call UniCode_Conv(Y_SYUREC.G_KENPIN_F, "")
                                            
                                            
                                            
                                            
                                            '出荷予定書込み
            Do
                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            KENPIN_Update_Proc = False
                            GoTo Abort_Tran
                        End If
                
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "出荷予定")
                        GoTo Abort_Tran
                End Select
            Loop
                                            
                                            
        '------------------------------- 出荷予定(ﾎｽﾄｲﾒｰｼﾞ)の処理 「品番毎の検品解除」
            Call UniCode_Conv(K4_Y_SYU_H.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))  ' ID��
        
            Do
            
                sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound
                        MsgBox "他端末で内容が変更されています。最新表示を行ってください。"
                        KENPIN_Update_Proc = False
                        GoTo Abort_Tran
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    
                        ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            KENPIN_Update_Proc = False
                            GoTo Abort_Tran
                        End If
                    
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                        GoTo Abort_Tran
                End Select
        
            Loop
            
            
            '未検品する
            Call UniCode_Conv(Y_SYU_HREC.KENPIN_NOW, "")
            Call UniCode_Conv(Y_SYU_HREC.KENPIN_TANTO_CODE, "")
                                            
            Call UniCode_Conv(Y_SYU_HREC.JURYO, "")
            Call UniCode_Conv(Y_SYU_HREC.SAI_SU, "")
            Call UniCode_Conv(Y_SYU_HREC.KUTI_SU, "")
                                            
            Call UniCode_Conv(Y_SYU_HREC.OKURI_NO, "")
                                            
                                            
            Call UniCode_Conv(Y_SYU_HREC.OKURI_NO_SEQ, "")
                                            
            Call UniCode_Conv(Y_SYU_HREC.KONPOU_F, "")
                                            
            Call UniCode_Conv(Y_SYU_HREC.KUTI_SU_TAN, "")
            Call UniCode_Conv(Y_SYU_HREC.SAI_SU_TAN, "")
                                            
            Call UniCode_Conv(Y_SYU_HREC.OKURI_NO_SEQ_TO, "")
                                            
                                            
            '仕掛検品を解除する   2012.10.24
            Call UniCode_Conv(Y_SYU_HREC.CNT_BARA_SU, "")
            Call UniCode_Conv(Y_SYU_HREC.CNT_HAKO_SU, "")
            Call UniCode_Conv(Y_SYU_HREC.GAISO_IRI_QTY, "")
            Call UniCode_Conv(Y_SYU_HREC.Y_HIN_CHK_CNT, "")
            Call UniCode_Conv(Y_SYU_HREC.J_HIN_CHK_CNT, "")
            Call UniCode_Conv(Y_SYU_HREC.KEN_HINBAN, "")
    
    
    
    
    
                                            
                                            
                                            '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)書込み
            Do
                sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            KENPIN_Update_Proc = False
                            GoTo Abort_Tran
                        End If
                
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                        GoTo Abort_Tran
                End Select
            Loop
    
''        End If
        
        
        
        '------------------------------- 注文ﾃﾞｰﾀの処理 「品番毎の検品解除」
            Call UniCode_Conv(K2_Y_SYU_TEI.KEN_NO, StrConv(Y_SYU_HREC.SEK_KEN_NO, vbUnicode))
            Call UniCode_Conv(K2_Y_SYU_TEI.HIN_NO, StrConv(Y_SYU_HREC.SEK_HIN_NO, vbUnicode))
        
            Do
            
                sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K2_Y_SYU_TEI, Len(K2_Y_SYU_TEI), 2)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    
                        ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA_TEI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            KENPIN_Update_Proc = False
                            GoTo Abort_Tran
                        End If
                    
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "邸別注文ﾃﾞｰﾀ")
                        GoTo Abort_Tran
                End Select
        
            Loop
        
        
            If sts = BtNoErr Then
            
            
                Call UniCode_Conv(Y_SYU_TEI_REC.KENPIN_TANTO, "")
                Call UniCode_Conv(Y_SYU_TEI_REC.KENPIN_DATETIME, "")
            
            
                
                If Trim(StrConv(Y_SYU_TEI_REC.KONPO_ID, vbUnicode)) <> "" Then
                    Call UniCode_Conv(Y_SYU_TEI_REC.SAI_SU, "999999")
                Else
                    Call UniCode_Conv(Y_SYU_TEI_REC.SAI_SU, "000.00")
                End If
            
            
            
                Call UniCode_Conv(Y_SYU_TEI_REC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
                Call UniCode_Conv(Y_SYU_TEI_REC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
            
            
            
                Do
                    sts = BTRV(BtOpUpdate, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K2_Y_SYU_TEI, Len(K2_Y_SYU_TEI), 2)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                            ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA_TEI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                KENPIN_Update_Proc = False
                                GoTo Abort_Tran
                            End If
                    
                        Case Else
                            Call File_Error(sts, BtOpUpdate, "邸別注文ﾃﾞｰﾀ")
                            GoTo Abort_Tran
                    End Select
                Loop
            
            End If
        
        
        com = BtOpGetNext
    
    Loop
        
                                        
End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
    
    KENPIN_Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If


End Function


