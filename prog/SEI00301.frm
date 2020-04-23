VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form SEI00301 
   Caption         =   "[請求システム]輸送箱実績入力"
   ClientHeight    =   9960
   ClientLeft      =   2025
   ClientTop       =   -3360
   ClientWidth     =   14340
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
   ScaleHeight     =   9960
   ScaleWidth      =   14340
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   4
      Left            =   5145
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   3
      Left            =   5145
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終 了"
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
      Left            =   3570
      TabIndex        =   11
      Top             =   120
      Width           =   1380
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   1365
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1365
      MaxLength       =   5
      TabIndex        =   0
      Top             =   960
      Width           =   750
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
      Height          =   375
      Index           =   1
      Left            =   2100
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   1365
      MaxLength       =   10
      TabIndex        =   2
      Top             =   2280
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   5355
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton Command1 
      Caption         =   "表 示"
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
      Left            =   1890
      TabIndex        =   5
      Top             =   120
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "更 新"
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
      TabIndex        =   4
      Top             =   120
      Width           =   1380
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   6735
      Left            =   420
      TabIndex        =   3
      Top             =   3000
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   11880
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "事業部"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "国内外"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "品番"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "品名"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "才数"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).DataField=   ""
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).DataField=   ""
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).DataField=   ""
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).DataField=   ""
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).DataField=   ""
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(26)._VlistStyle=   0
      Columns(26)._MaxComboItems=   5
      Columns(26).DataField=   ""
      Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(27)._VlistStyle=   0
      Columns(27)._MaxComboItems=   5
      Columns(27).DataField=   ""
      Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(28)._VlistStyle=   0
      Columns(28)._MaxComboItems=   5
      Columns(28).DataField=   ""
      Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(29)._VlistStyle=   0
      Columns(29)._MaxComboItems=   5
      Columns(29).DataField=   ""
      Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(30)._VlistStyle=   0
      Columns(30)._MaxComboItems=   5
      Columns(30).DataField=   ""
      Columns(30)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(31)._VlistStyle=   0
      Columns(31)._MaxComboItems=   5
      Columns(31).DataField=   ""
      Columns(31)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(32)._VlistStyle=   0
      Columns(32)._MaxComboItems=   5
      Columns(32).DataField=   ""
      Columns(32)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(33)._VlistStyle=   0
      Columns(33)._MaxComboItems=   5
      Columns(33).DataField=   ""
      Columns(33)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(34)._VlistStyle=   0
      Columns(34)._MaxComboItems=   5
      Columns(34).DataField=   ""
      Columns(34)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(35)._VlistStyle=   0
      Columns(35)._MaxComboItems=   5
      Columns(35).DataField=   ""
      Columns(35)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   36
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=36"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4366"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
      Splits(0)._ColumnProps(4)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=4366"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4233"
      Splits(0)._ColumnProps(9)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=4366"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4233"
      Splits(0)._ColumnProps(14)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2646"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2514"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=4842"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=4710"
      Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=8196"
      Splits(0)._ColumnProps(24)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=1323"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1191"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=8194"
      Splits(0)._ColumnProps(30)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(32)=   "Column(6).Width=2328"
      Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=2196"
      Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=8194"
      Splits(0)._ColumnProps(36)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(38)=   "Column(7).Width=2328"
      Splits(0)._ColumnProps(39)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(7)._WidthInPix=2196"
      Splits(0)._ColumnProps(41)=   "Column(7)._ColStyle=8194"
      Splits(0)._ColumnProps(42)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(44)=   "Column(8).Width=2328"
      Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=2196"
      Splits(0)._ColumnProps(47)=   "Column(8)._ColStyle=8194"
      Splits(0)._ColumnProps(48)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(49)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(50)=   "Column(9).Width=2328"
      Splits(0)._ColumnProps(51)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(9)._WidthInPix=2196"
      Splits(0)._ColumnProps(53)=   "Column(9)._ColStyle=8194"
      Splits(0)._ColumnProps(54)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(55)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(56)=   "Column(10).Width=2328"
      Splits(0)._ColumnProps(57)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(10)._WidthInPix=2196"
      Splits(0)._ColumnProps(59)=   "Column(10)._ColStyle=8194"
      Splits(0)._ColumnProps(60)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(61)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(62)=   "Column(11).Width=2328"
      Splits(0)._ColumnProps(63)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(11)._WidthInPix=2196"
      Splits(0)._ColumnProps(65)=   "Column(11)._ColStyle=8194"
      Splits(0)._ColumnProps(66)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(67)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(68)=   "Column(12).Width=2328"
      Splits(0)._ColumnProps(69)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(70)=   "Column(12)._WidthInPix=2196"
      Splits(0)._ColumnProps(71)=   "Column(12)._ColStyle=8194"
      Splits(0)._ColumnProps(72)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(73)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(74)=   "Column(13).Width=2328"
      Splits(0)._ColumnProps(75)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(76)=   "Column(13)._WidthInPix=2196"
      Splits(0)._ColumnProps(77)=   "Column(13)._ColStyle=8194"
      Splits(0)._ColumnProps(78)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(79)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(80)=   "Column(14).Width=2328"
      Splits(0)._ColumnProps(81)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(82)=   "Column(14)._WidthInPix=2196"
      Splits(0)._ColumnProps(83)=   "Column(14)._ColStyle=8194"
      Splits(0)._ColumnProps(84)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(85)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(86)=   "Column(15).Width=2328"
      Splits(0)._ColumnProps(87)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(88)=   "Column(15)._WidthInPix=2196"
      Splits(0)._ColumnProps(89)=   "Column(15)._ColStyle=8194"
      Splits(0)._ColumnProps(90)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(91)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(92)=   "Column(16).Width=2328"
      Splits(0)._ColumnProps(93)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(94)=   "Column(16)._WidthInPix=2196"
      Splits(0)._ColumnProps(95)=   "Column(16)._ColStyle=8194"
      Splits(0)._ColumnProps(96)=   "Column(16).Visible=0"
      Splits(0)._ColumnProps(97)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(98)=   "Column(17).Width=2328"
      Splits(0)._ColumnProps(99)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(100)=   "Column(17)._WidthInPix=2196"
      Splits(0)._ColumnProps(101)=   "Column(17)._ColStyle=8194"
      Splits(0)._ColumnProps(102)=   "Column(17).Visible=0"
      Splits(0)._ColumnProps(103)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(104)=   "Column(18).Width=2328"
      Splits(0)._ColumnProps(105)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(106)=   "Column(18)._WidthInPix=2196"
      Splits(0)._ColumnProps(107)=   "Column(18)._ColStyle=8194"
      Splits(0)._ColumnProps(108)=   "Column(18).Visible=0"
      Splits(0)._ColumnProps(109)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(110)=   "Column(19).Width=2328"
      Splits(0)._ColumnProps(111)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(112)=   "Column(19)._WidthInPix=2196"
      Splits(0)._ColumnProps(113)=   "Column(19)._ColStyle=8194"
      Splits(0)._ColumnProps(114)=   "Column(19).Visible=0"
      Splits(0)._ColumnProps(115)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(116)=   "Column(20).Width=2328"
      Splits(0)._ColumnProps(117)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(118)=   "Column(20)._WidthInPix=2196"
      Splits(0)._ColumnProps(119)=   "Column(20)._ColStyle=8194"
      Splits(0)._ColumnProps(120)=   "Column(20).Visible=0"
      Splits(0)._ColumnProps(121)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(122)=   "Column(21).Width=2328"
      Splits(0)._ColumnProps(123)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(124)=   "Column(21)._WidthInPix=2196"
      Splits(0)._ColumnProps(125)=   "Column(21)._ColStyle=2"
      Splits(0)._ColumnProps(126)=   "Column(21).Visible=0"
      Splits(0)._ColumnProps(127)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(128)=   "Column(22).Width=2328"
      Splits(0)._ColumnProps(129)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(130)=   "Column(22)._WidthInPix=2196"
      Splits(0)._ColumnProps(131)=   "Column(22)._ColStyle=2"
      Splits(0)._ColumnProps(132)=   "Column(22).Visible=0"
      Splits(0)._ColumnProps(133)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(134)=   "Column(23).Width=2328"
      Splits(0)._ColumnProps(135)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(136)=   "Column(23)._WidthInPix=2196"
      Splits(0)._ColumnProps(137)=   "Column(23)._ColStyle=2"
      Splits(0)._ColumnProps(138)=   "Column(23).Visible=0"
      Splits(0)._ColumnProps(139)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(140)=   "Column(24).Width=2328"
      Splits(0)._ColumnProps(141)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(142)=   "Column(24)._WidthInPix=2196"
      Splits(0)._ColumnProps(143)=   "Column(24)._ColStyle=2"
      Splits(0)._ColumnProps(144)=   "Column(24).Visible=0"
      Splits(0)._ColumnProps(145)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(146)=   "Column(25).Width=2328"
      Splits(0)._ColumnProps(147)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(148)=   "Column(25)._WidthInPix=2196"
      Splits(0)._ColumnProps(149)=   "Column(25)._ColStyle=2"
      Splits(0)._ColumnProps(150)=   "Column(25).Visible=0"
      Splits(0)._ColumnProps(151)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(152)=   "Column(26).Width=2328"
      Splits(0)._ColumnProps(153)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(154)=   "Column(26)._WidthInPix=2196"
      Splits(0)._ColumnProps(155)=   "Column(26)._ColStyle=2"
      Splits(0)._ColumnProps(156)=   "Column(26).Visible=0"
      Splits(0)._ColumnProps(157)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(158)=   "Column(27).Width=2328"
      Splits(0)._ColumnProps(159)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(160)=   "Column(27)._WidthInPix=2196"
      Splits(0)._ColumnProps(161)=   "Column(27)._ColStyle=2"
      Splits(0)._ColumnProps(162)=   "Column(27).Visible=0"
      Splits(0)._ColumnProps(163)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(164)=   "Column(28).Width=2328"
      Splits(0)._ColumnProps(165)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(166)=   "Column(28)._WidthInPix=2196"
      Splits(0)._ColumnProps(167)=   "Column(28)._ColStyle=2"
      Splits(0)._ColumnProps(168)=   "Column(28).Visible=0"
      Splits(0)._ColumnProps(169)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(170)=   "Column(29).Width=2328"
      Splits(0)._ColumnProps(171)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(172)=   "Column(29)._WidthInPix=2196"
      Splits(0)._ColumnProps(173)=   "Column(29)._ColStyle=2"
      Splits(0)._ColumnProps(174)=   "Column(29).Visible=0"
      Splits(0)._ColumnProps(175)=   "Column(29).Order=30"
      Splits(0)._ColumnProps(176)=   "Column(30).Width=2328"
      Splits(0)._ColumnProps(177)=   "Column(30).DividerColor=0"
      Splits(0)._ColumnProps(178)=   "Column(30)._WidthInPix=2196"
      Splits(0)._ColumnProps(179)=   "Column(30)._ColStyle=2"
      Splits(0)._ColumnProps(180)=   "Column(30).Visible=0"
      Splits(0)._ColumnProps(181)=   "Column(30).Order=31"
      Splits(0)._ColumnProps(182)=   "Column(31).Width=2328"
      Splits(0)._ColumnProps(183)=   "Column(31).DividerColor=0"
      Splits(0)._ColumnProps(184)=   "Column(31)._WidthInPix=2196"
      Splits(0)._ColumnProps(185)=   "Column(31)._ColStyle=2"
      Splits(0)._ColumnProps(186)=   "Column(31).Visible=0"
      Splits(0)._ColumnProps(187)=   "Column(31).Order=32"
      Splits(0)._ColumnProps(188)=   "Column(32).Width=2328"
      Splits(0)._ColumnProps(189)=   "Column(32).DividerColor=0"
      Splits(0)._ColumnProps(190)=   "Column(32)._WidthInPix=2196"
      Splits(0)._ColumnProps(191)=   "Column(32)._ColStyle=2"
      Splits(0)._ColumnProps(192)=   "Column(32).Visible=0"
      Splits(0)._ColumnProps(193)=   "Column(32).Order=33"
      Splits(0)._ColumnProps(194)=   "Column(33).Width=2328"
      Splits(0)._ColumnProps(195)=   "Column(33).DividerColor=0"
      Splits(0)._ColumnProps(196)=   "Column(33)._WidthInPix=2196"
      Splits(0)._ColumnProps(197)=   "Column(33)._ColStyle=2"
      Splits(0)._ColumnProps(198)=   "Column(33).Visible=0"
      Splits(0)._ColumnProps(199)=   "Column(33).Order=34"
      Splits(0)._ColumnProps(200)=   "Column(34).Width=2328"
      Splits(0)._ColumnProps(201)=   "Column(34).DividerColor=0"
      Splits(0)._ColumnProps(202)=   "Column(34)._WidthInPix=2196"
      Splits(0)._ColumnProps(203)=   "Column(34)._ColStyle=2"
      Splits(0)._ColumnProps(204)=   "Column(34).Visible=0"
      Splits(0)._ColumnProps(205)=   "Column(34).Order=35"
      Splits(0)._ColumnProps(206)=   "Column(35).Width=2328"
      Splits(0)._ColumnProps(207)=   "Column(35).DividerColor=0"
      Splits(0)._ColumnProps(208)=   "Column(35)._WidthInPix=2196"
      Splits(0)._ColumnProps(209)=   "Column(35)._ColStyle=2"
      Splits(0)._ColumnProps(210)=   "Column(35).Visible=0"
      Splits(0)._ColumnProps(211)=   "Column(35).Order=36"
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
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "使用実績"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=ＭＳ Ｐゴシック"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
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
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFFF80&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFFFF00&"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=87,.parent=1,.bgcolor=&H80FF80&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=96,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=89,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=91,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=94,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=95,.parent=10,.bgcolor=&HFFFF00&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=110,.parent=87"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=107,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=108,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=109,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=106,.parent=87"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=103,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=104,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=105,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=102,.parent=87"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=99,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=100,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=101,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=114,.parent=87"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=111,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=112,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=113,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=118,.parent=87,.locked=-1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=115,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=116,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=117,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=16,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=13,.parent=88"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=14,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=15,.parent=91"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=20,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=17,.parent=88"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=18,.parent=89"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=19,.parent=91"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=24,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=21,.parent=88"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=22,.parent=89"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=23,.parent=91"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=28,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=25,.parent=88"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=26,.parent=89"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=27,.parent=91"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=32,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=29,.parent=88"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=30,.parent=89"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=31,.parent=91"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=46,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=43,.parent=88"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=44,.parent=89"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=45,.parent=91"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=50,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=47,.parent=88"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=48,.parent=89"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=49,.parent=91"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=54,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=51,.parent=88"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=52,.parent=89"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=53,.parent=91"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=58,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=55,.parent=88"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=56,.parent=89"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=57,.parent=91"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=62,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=59,.parent=88"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=60,.parent=89"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=61,.parent=91"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=66,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=63,.parent=88"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=64,.parent=89"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=65,.parent=91"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=70,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=67,.parent=88"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=68,.parent=89"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=69,.parent=91"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=74,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=71,.parent=88"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=72,.parent=89"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=73,.parent=91"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=78,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=75,.parent=88"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=76,.parent=89"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=77,.parent=91"
      _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=82,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=79,.parent=88"
      _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=80,.parent=89"
      _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=81,.parent=91"
      _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=86,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=83,.parent=88"
      _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=84,.parent=89"
      _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=85,.parent=91"
      _StyleDefs(120) =   "Splits(0).Columns(21).Style:id=122,.parent=87,.alignment=1"
      _StyleDefs(121) =   "Splits(0).Columns(21).HeadingStyle:id=119,.parent=88"
      _StyleDefs(122) =   "Splits(0).Columns(21).FooterStyle:id=120,.parent=89"
      _StyleDefs(123) =   "Splits(0).Columns(21).EditorStyle:id=121,.parent=91"
      _StyleDefs(124) =   "Splits(0).Columns(22).Style:id=126,.parent=87,.alignment=1"
      _StyleDefs(125) =   "Splits(0).Columns(22).HeadingStyle:id=123,.parent=88"
      _StyleDefs(126) =   "Splits(0).Columns(22).FooterStyle:id=124,.parent=89"
      _StyleDefs(127) =   "Splits(0).Columns(22).EditorStyle:id=125,.parent=91"
      _StyleDefs(128) =   "Splits(0).Columns(23).Style:id=130,.parent=87,.alignment=1"
      _StyleDefs(129) =   "Splits(0).Columns(23).HeadingStyle:id=127,.parent=88"
      _StyleDefs(130) =   "Splits(0).Columns(23).FooterStyle:id=128,.parent=89"
      _StyleDefs(131) =   "Splits(0).Columns(23).EditorStyle:id=129,.parent=91"
      _StyleDefs(132) =   "Splits(0).Columns(24).Style:id=134,.parent=87,.alignment=1"
      _StyleDefs(133) =   "Splits(0).Columns(24).HeadingStyle:id=131,.parent=88"
      _StyleDefs(134) =   "Splits(0).Columns(24).FooterStyle:id=132,.parent=89"
      _StyleDefs(135) =   "Splits(0).Columns(24).EditorStyle:id=133,.parent=91"
      _StyleDefs(136) =   "Splits(0).Columns(25).Style:id=138,.parent=87,.alignment=1"
      _StyleDefs(137) =   "Splits(0).Columns(25).HeadingStyle:id=135,.parent=88"
      _StyleDefs(138) =   "Splits(0).Columns(25).FooterStyle:id=136,.parent=89"
      _StyleDefs(139) =   "Splits(0).Columns(25).EditorStyle:id=137,.parent=91"
      _StyleDefs(140) =   "Splits(0).Columns(26).Style:id=142,.parent=87,.alignment=1"
      _StyleDefs(141) =   "Splits(0).Columns(26).HeadingStyle:id=139,.parent=88"
      _StyleDefs(142) =   "Splits(0).Columns(26).FooterStyle:id=140,.parent=89"
      _StyleDefs(143) =   "Splits(0).Columns(26).EditorStyle:id=141,.parent=91"
      _StyleDefs(144) =   "Splits(0).Columns(27).Style:id=146,.parent=87,.alignment=1"
      _StyleDefs(145) =   "Splits(0).Columns(27).HeadingStyle:id=143,.parent=88"
      _StyleDefs(146) =   "Splits(0).Columns(27).FooterStyle:id=144,.parent=89"
      _StyleDefs(147) =   "Splits(0).Columns(27).EditorStyle:id=145,.parent=91"
      _StyleDefs(148) =   "Splits(0).Columns(28).Style:id=150,.parent=87,.alignment=1"
      _StyleDefs(149) =   "Splits(0).Columns(28).HeadingStyle:id=147,.parent=88"
      _StyleDefs(150) =   "Splits(0).Columns(28).FooterStyle:id=148,.parent=89"
      _StyleDefs(151) =   "Splits(0).Columns(28).EditorStyle:id=149,.parent=91"
      _StyleDefs(152) =   "Splits(0).Columns(29).Style:id=154,.parent=87,.alignment=1"
      _StyleDefs(153) =   "Splits(0).Columns(29).HeadingStyle:id=151,.parent=88"
      _StyleDefs(154) =   "Splits(0).Columns(29).FooterStyle:id=152,.parent=89"
      _StyleDefs(155) =   "Splits(0).Columns(29).EditorStyle:id=153,.parent=91"
      _StyleDefs(156) =   "Splits(0).Columns(30).Style:id=158,.parent=87,.alignment=1"
      _StyleDefs(157) =   "Splits(0).Columns(30).HeadingStyle:id=155,.parent=88"
      _StyleDefs(158) =   "Splits(0).Columns(30).FooterStyle:id=156,.parent=89"
      _StyleDefs(159) =   "Splits(0).Columns(30).EditorStyle:id=157,.parent=91"
      _StyleDefs(160) =   "Splits(0).Columns(31).Style:id=162,.parent=87,.alignment=1"
      _StyleDefs(161) =   "Splits(0).Columns(31).HeadingStyle:id=159,.parent=88"
      _StyleDefs(162) =   "Splits(0).Columns(31).FooterStyle:id=160,.parent=89"
      _StyleDefs(163) =   "Splits(0).Columns(31).EditorStyle:id=161,.parent=91"
      _StyleDefs(164) =   "Splits(0).Columns(32).Style:id=166,.parent=87,.alignment=1"
      _StyleDefs(165) =   "Splits(0).Columns(32).HeadingStyle:id=163,.parent=88"
      _StyleDefs(166) =   "Splits(0).Columns(32).FooterStyle:id=164,.parent=89"
      _StyleDefs(167) =   "Splits(0).Columns(32).EditorStyle:id=165,.parent=91"
      _StyleDefs(168) =   "Splits(0).Columns(33).Style:id=170,.parent=87,.alignment=1"
      _StyleDefs(169) =   "Splits(0).Columns(33).HeadingStyle:id=167,.parent=88"
      _StyleDefs(170) =   "Splits(0).Columns(33).FooterStyle:id=168,.parent=89"
      _StyleDefs(171) =   "Splits(0).Columns(33).EditorStyle:id=169,.parent=91"
      _StyleDefs(172) =   "Splits(0).Columns(34).Style:id=174,.parent=87,.alignment=1"
      _StyleDefs(173) =   "Splits(0).Columns(34).HeadingStyle:id=171,.parent=88"
      _StyleDefs(174) =   "Splits(0).Columns(34).FooterStyle:id=172,.parent=89"
      _StyleDefs(175) =   "Splits(0).Columns(34).EditorStyle:id=173,.parent=91"
      _StyleDefs(176) =   "Splits(0).Columns(35).Style:id=178,.parent=87,.alignment=1"
      _StyleDefs(177) =   "Splits(0).Columns(35).HeadingStyle:id=175,.parent=88"
      _StyleDefs(178) =   "Splits(0).Columns(35).FooterStyle:id=176,.parent=89"
      _StyleDefs(179) =   "Splits(0).Columns(35).EditorStyle:id=177,.parent=91"
      _StyleDefs(180) =   "Named:id=33:Normal"
      _StyleDefs(181) =   ":id=33,.parent=0"
      _StyleDefs(182) =   "Named:id=34:Heading"
      _StyleDefs(183) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(184) =   ":id=34,.wraptext=-1"
      _StyleDefs(185) =   "Named:id=35:Footing"
      _StyleDefs(186) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(187) =   "Named:id=36:Selected"
      _StyleDefs(188) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(189) =   "Named:id=37:Caption"
      _StyleDefs(190) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(191) =   "Named:id=38:HighlightRow"
      _StyleDefs(192) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(193) =   "Named:id=39:EvenRow"
      _StyleDefs(194) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(195) =   "Named:id=40:OddRow"
      _StyleDefs(196) =   ":id=40,.parent=33"
      _StyleDefs(197) =   "Named:id=41:RecordSelector"
      _StyleDefs(198) =   ":id=41,.parent=34"
      _StyleDefs(199) =   "Named:id=42:FilterBar"
      _StyleDefs(200) =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      Caption         =   "才数合計"
      Height          =   255
      Index           =   4
      Left            =   3990
      TabIndex        =   14
      Top             =   2400
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "枚数合計"
      Height          =   255
      Index           =   3
      Left            =   3990
      TabIndex        =   12
      Top             =   1800
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "仕向け先"
      Height          =   255
      Index           =   2
      Left            =   315
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "担当者"
      Height          =   375
      Index           =   0
      Left            =   210
      TabIndex        =   9
      Top             =   960
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "売上日付"
      Height          =   375
      Index           =   1
      Left            =   315
      TabIndex        =   8
      Top             =   2400
      Width           =   1065
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "更新"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "表 示"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   2
      End
      Begin VB.Menu SHORI 
         Caption         =   "画面印刷"
         Index           =   3
      End
   End
End
Attribute VB_Name = "SEI00301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Const pcmbSHIMUKE% = 0          '仕向け先


Private Const ptxTanto_Code% = 0        '担当者コード
Private Const ptxTanto_Name% = 1        '担当者名称
Private Const ptxJITU_Date% = 2         '実績日付

Private Const ptxMAISU% = 3             '枚数合計
Private Const ptxSAISU% = 4             '才数合計



Dim SE_USOU_HAKO As New XArrayDB

Private Const Min_Row% = 1              '最小行数

Dim Max_Row    As Integer               'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 35             '最大列数

Private Const ColSE_USOU_F% = 0         '事業部
Private Const ColJGYOBU% = 1            '事業部
Private Const ColNAIGAI% = 2            '内外
Private Const ColHIN_GAI% = 3           '品番
Private Const ColHIN_NAME% = 4          '品名
Private Const ColSAI_SU% = 5            '才数
Private Const ColMUKE% = 6              '向け先




Private MUKE_TBL()  As String * 8       '対象向け先コード


Private Type SE_JGYOBU_TBL

    SHIMUKE_CODE    As String * 2
    
    JGYOBU          As String * 1
    NAIGAI          As String * 1

End Type


Private SE_JGYOBU_T()       As SE_JGYOBU_TBL
Private Last_SHIMUKE_CODE   As String * 2

Private Const LAST_UPDATE_DAY$ = "[SEI0030]2013.04.01 14:00"

Private Sub Command1_Click(Index As Integer)

Dim yn  As Integer
Dim i   As Integer

    Select Case Index
    
        Case 0
        
            For i = ptxTanto_Code To ptxJITU_Date
                If Error_Check_Proc(Index) Then     'エラーチェック
                    Exit Sub
                End If
            Next i
        
        
            If Grid_Error_Check_Proc() Then
                Exit Sub
            End If
        
        
        
            yn = MsgBox("更新しますか？", vbYesNo, "確認入力")
            If yn = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            
                If List_Disp_Proc() Then
                    Unload Me
                End If
            
            
            
            End If
        Case 1
            If List_Disp_Proc() Then
                Unload Me
            End If
        Case 2
            
            Unload Me
    End Select



End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
Private Sub Form_Load()
Dim i   As Integer
Dim c   As String * 128
Dim sts As Integer

Dim com As Integer


    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If


    
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]輸送箱実績入力", Me.hwnd, 0)
    'ペイン複数作る
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

    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If


    SEI00301.Caption = SEI00301.Caption & " " & LAST_UPDATE_DAY

    Max_Row = 9999
                                

                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '構成マスタＯＰＥＮ
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '向け先マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '輸送箱実績マスタＯＰＥＮ
    If SE_USOU_HAKO_Open(BtOpenNomal) Then
        Unload Me
    End If



    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc
    '仕向け先のセット
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD, 0) Then
        Unload Me
    End If
    Combo1(pcmbSHIMUKE).ListIndex = 0



    '仕向け先をﾃｰﾌﾞﾙにｾｯﾄ
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")

    com = BtOpGetGreater

    i = -1
    Do
        DoEvents
        
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> P_KBN04_CD Then
                    Exit Do
                End If
            
                i = i + 1
                ReDim Preserve SE_JGYOBU_T(0 To i)
                SE_JGYOBU_T(i).SHIMUKE_CODE = Trim(StrConv(P_CODEREC.C_Code, vbUnicode))
                SE_JGYOBU_T(i).JGYOBU = Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
                SE_JGYOBU_T(i).NAIGAI = Trim(StrConv(P_CODEREC.OPTION2, vbUnicode))
                        
                        
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "コードマスタ")
                Unload Me
        End Select
    
        com = BtOpGetNext
    
    Loop

    


                                '初期画面作成
    If Init_Disp_Proc() Then
        Unload Me
    End If
    
    
    If List_Disp_Proc() Then
        Unload Me
    End If

    Text1(ptxJITU_Date).Text = Format(Now, "YYYY/MM/DD")


    Text1(ptxTanto_Code).SetFocus

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
                                            '構成マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '向け先マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先マスタ")
        End If
    End If
                                            
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "コードマスタ")
        End If
    End If
                                            
                                            '輸送箱実績ＣＬＯＳＥ
    sts = BTRV(BtOpClose, SE_USOU_HAKO_POS, SE_USOU_HAKOREC, Len(SE_USOU_HAKOREC), K0_SE_USOU_HAKO, Len(K0_SE_USOU_HAKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "輸送箱実績")
        End If
    End If
    
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    End
End Sub





Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    SEI00301.MousePointer = vbHourglass

    TDBGrid1.Enabled = False

    Call Ctrl_Lock(SEI00301)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(SEI00301)

    TDBGrid1.Enabled = True

    SEI00301.MousePointer = vbDefault

End Sub
Private Function Init_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   INIﾌｧｲﾙより入力可能箇所の設定
'----------------------------------------------------------------------------
Dim c       As String * 128

Dim i       As Integer
Dim j       As Integer

Dim sts     As Integer

Dim Row     As Integer

Dim ITEM    As Variant

Dim com     As Integer
    
    Init_Disp_Proc = True





    ReDim MUKE_TBL(0 To 0)

    i = 1
    j = ColMUKE - 1
    
    Do
        If GetIni("MUKE", "MUKE" & Format(i, "00"), "SEI_SYS", c) Then
            Exit Do
        Else
        
        
        
            If Trim(c) = "********" Then
        
        
                j = j + 1
                TDBGrid1.Columns(j).Caption = "その他"
                TDBGrid1.Columns(j).Locked = False
                TDBGrid1.Columns(j).Visible = True
                
                
                If i > 1 Then
                    ReDim Preserve MUKE_TBL(0 To i - 1)
                End If
                
                MUKE_TBL(i - 1) = Trim(c)
            
            Else
        
        
                Call UniCode_Conv(K0_MTS.MUKE_CODE, Trim(c))
                Call UniCode_Conv(K0_MTS.SS_CODE, "")
                sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                Select Case sts
                    Case BtNoErr
                        j = j + 1
                        TDBGrid1.Columns(j).Caption = Trim(StrConv(MTSREC.MUKE_DNAME, vbUnicode))
                        TDBGrid1.Columns(j).Locked = False
                        TDBGrid1.Columns(j).Visible = True
                        
                        
                        If i > 1 Then
                            ReDim Preserve MUKE_TBL(0 To i - 1)
                        End If
                        
                        MUKE_TBL(i - 1) = Trim(c)
                    
                    Case BtErrKeyNotFound
                        
                        MsgBox "向け先コード（" & Trim(c) & "）登録させていません"
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "向け先マスタ")
                        Exit Function
                End Select
            End If
        End If
        i = i + 1
    
    Loop



    Set SE_USOU_HAKO = Nothing
    
'    ReDim Hinban(0 To 0)
    
    
'    Row = 0
    
'    Do
'        If GetIni("HIN", "HIN" & Format(Row + 1, "00"), App.EXEName, c) Then
'            Exit Do
'        Else
'
'            ITEM = Split(Trim(c), ",", -1)
'
'
'            Call UniCode_Conv(K0_ITEM.JGYOBU, Format(ITEM(0)))
'            Call UniCode_Conv(K0_ITEM.NAIGAI, Format(ITEM(1)))
'            Call UniCode_Conv(K0_ITEM.HIN_GAI, Format(ITEM(2)))
'            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'            Select Case sts
'                Case BtNoErr
'                Case BtErrKeyNotFound
'                    MsgBox "品目コード（" & Trim(ITEM(2)) & "）登録させていません"
'                    Exit Function
'                Case Else
'                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
'                    Exit Function
'            End Select
'            Row = Row + 1
'            SE_USOU_HAKO.ReDim Min_Row, Row, Min_Col, Max_Col
'            SE_USOU_HAKO(Row, ColJGYOBU) = Trim(StrConv(ITEMREC.JGYOBU, vbUnicode))
'            SE_USOU_HAKO(Row, ColNAIGAI) = Trim(StrConv(ITEMREC.NAIGAI, vbUnicode))
'            SE_USOU_HAKO(Row, ColHIN_GAI) = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
'            SE_USOU_HAKO(Row, ColHIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
'            If IsNumeric(StrConv(ITEMREC.SAI_SU, vbUnicode)) Then
'                SE_USOU_HAKO(Row, ColSAI_SU) = Format(CDbl(StrConv(ITEMREC.SAI_SU, vbUnicode)), "#.0")
'            Else
'                SE_USOU_HAKO(Row, ColSAI_SU) = "0.0"
'            End If
'
'
'
'
'        End If
'
'        i = i + 1
'    Loop


    



    SE_USOU_HAKO.ReDim Min_Row, 1, Min_Col, Max_Col
    SE_USOU_HAKO(1, ColSE_USOU_F) = " "
    SE_USOU_HAKO(1, ColJGYOBU) = "*"
    SE_USOU_HAKO(1, ColNAIGAI) = "*"
    SE_USOU_HAKO(1, ColHIN_GAI) = "**********"
    SE_USOU_HAKO(1, ColHIN_NAME) = "[枚数合計]"
    SE_USOU_HAKO(1, ColSAI_SU) = "**.*"




    SE_USOU_HAKO.ReDim Min_Row, 2, Min_Col, Max_Col
    SE_USOU_HAKO(2, ColSE_USOU_F) = "!"
    SE_USOU_HAKO(2, ColJGYOBU) = "*"
    SE_USOU_HAKO(2, ColNAIGAI) = "*"
    SE_USOU_HAKO(2, ColHIN_GAI) = "**********"
    SE_USOU_HAKO(2, ColHIN_NAME) = "[才数合計]"
    SE_USOU_HAKO(2, ColSAI_SU) = "**.*"

    Row = 2
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, "")

    
    com = BtOpGetGreater

    Do
        
        DoEvents
        
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> SHIZAI Or _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) <> NAIGAI_NAI Then
                    Exit Do
                End If
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        End Select



        If Trim(StrConv(ITEMREC.SE_USOU_F, vbUnicode)) = "" Then
        Else

            Row = Row + 1
            SE_USOU_HAKO.ReDim Min_Row, Row, Min_Col, Max_Col
            
            
            If IsNumeric(Trim(StrConv(ITEMREC.SE_USOU_F, vbUnicode))) Then
                SE_USOU_HAKO(Row, ColSE_USOU_F) = Format(Trim(StrConv(ITEMREC.SE_USOU_F, vbUnicode)), "00")
            Else
                
                SE_USOU_HAKO(Row, ColSE_USOU_F) = Trim(StrConv(ITEMREC.SE_USOU_F, vbUnicode))
            End If
            
            
            
            
            SE_USOU_HAKO(Row, ColJGYOBU) = Trim(StrConv(ITEMREC.JGYOBU, vbUnicode))
            SE_USOU_HAKO(Row, ColNAIGAI) = Trim(StrConv(ITEMREC.NAIGAI, vbUnicode))
            SE_USOU_HAKO(Row, ColHIN_GAI) = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
            SE_USOU_HAKO(Row, ColHIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
            If IsNumeric(StrConv(ITEMREC.SAI_SU, vbUnicode)) Then
                 SE_USOU_HAKO(Row, ColSAI_SU) = Format(CDbl(StrConv(ITEMREC.SAI_SU, vbUnicode)), "#.0")
            Else
                 SE_USOU_HAKO(Row, ColSAI_SU) = "0.0"
            End If
    
        End If
    
        com = BtOpGetNext
    
    Loop


    SE_USOU_HAKO.QuickSort Min_Row, SE_USOU_HAKO.UpperBound(1), ColSE_USOU_F, 0, XTYPE_STRING

                                'DBテーブルリンク
    Set TDBGrid1.Array = SE_USOU_HAKO
    
    
    TDBGrid1.Bookmark = Null
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.ScrollBars = dbgAutomatic




    Init_Disp_Proc = False
End Function
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

Private Sub SHORI_Click(Index As Integer)
    Select Case Index
    
        
        Case 0      '更新
        
        
            Command1(Index).Value = True
        
        
        Case 1      '表示
        
        
            Command1(Index).Value = True
        
        Case 2      '終了
        
        
            Command1(Index).Value = True
        
        
        
        Case 3      '画面印刷
        
            Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)
                    
    
    End Select

End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)

Dim sts         As Integer
Dim Bookmark    As Variant
    
Dim i           As Integer
Dim j           As Integer
Dim Err_Flg     As Boolean
    
    
Dim GK_MAISU    As Long
Dim GK_SAISU    As Long
    
Dim MAISU()     As Long
Dim SAISU()     As Long
    
    
    If TDBGrid1.Bookmark < 0 Then
        Exit Sub
    End If
    
    Set TDBGrid1.Array = SE_USOU_HAKO
'    TDBGrid1.Refresh
    TDBGrid1.Update
                
                
                
                
                
    If SE_USOU_HAKO(TDBGrid1.Bookmark, ColHIN_GAI) = "**********" Or _
        Trim(SE_USOU_HAKO(TDBGrid1.Bookmark, ColHIN_GAI)) = "" Then
        Exit Sub
    End If
                
                
                
    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, SE_USOU_HAKO(TDBGrid1.Bookmark, ColHIN_GAI))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                
    Select Case sts
        Case BtNoErr
            
            SE_USOU_HAKO(TDBGrid1.Bookmark, ColHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
            If IsNumeric(StrConv(ITEMREC.SAI_SU, vbUnicode)) Then
                SE_USOU_HAKO(TDBGrid1.Bookmark, ColSAI_SU) = Format(CDbl(StrConv(ITEMREC.SAI_SU, vbUnicode)), "#0.0")
            Else
                SE_USOU_HAKO(TDBGrid1.Bookmark, ColSAI_SU) = "0.0"
            End If
                
                
        Case BtErrKeyNotFound
                                        
            For i = 0 To UBound(SE_JGYOBU_T)
                                   
                Call UniCode_Conv(K0_ITEM.JGYOBU, SE_JGYOBU_T(i).JGYOBU)
                Call UniCode_Conv(K0_ITEM.NAIGAI, SE_JGYOBU_T(i).NAIGAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, SE_USOU_HAKO(TDBGrid1.Bookmark, ColHIN_GAI))
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        
                Select Case sts
                    Case BtNoErr
                        
                            
                        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, SE_JGYOBU_T(i).SHIMUKE_CODE)
                        Call UniCode_Conv(K0_P_COMPO.JGYOBU, SE_JGYOBU_T(i).JGYOBU)
                        Call UniCode_Conv(K0_P_COMPO.NAIGAI, SE_JGYOBU_T(i).NAIGAI)
                        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, SE_USOU_HAKO(TDBGrid1.Bookmark, ColHIN_GAI))
                        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_GAISOU)
                        Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
                                                    
                        sts = BTRV(BtOpGetGreater, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                    
                        Select Case sts
                            Case BtNoErr
                                    
                                If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> SE_JGYOBU_T(i).SHIMUKE_CODE Or _
                                    StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                                    StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                                    Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> SE_USOU_HAKO(TDBGrid1.Bookmark, ColHIN_GAI) Or _
                                    StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_GAISOU Then
                                Else
                                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            SE_USOU_HAKO(TDBGrid1.Bookmark, ColHIN_GAI) = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                                            SE_USOU_HAKO(TDBGrid1.Bookmark, ColHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                            If IsNumeric(StrConv(ITEMREC.SAI_SU, vbUnicode)) Then
                                                SE_USOU_HAKO(TDBGrid1.Bookmark, ColSAI_SU) = Format(CDbl(StrConv(ITEMREC.SAI_SU, vbUnicode)), "#0.0")
                                            Else
                                                SE_USOU_HAKO(TDBGrid1.Bookmark, ColSAI_SU) = "0.0"
                                            End If
                            
                            
                                            SE_USOU_HAKO(TDBGrid1.Bookmark, ColSE_USOU_F) = StrConv(ITEMREC.SE_USOU_F, vbUnicode)
                                        Case BtErrKeyNotFound
                                        
                                            MsgBox "品番未登録です！！"
                                            SE_USOU_HAKO(TDBGrid1.Bookmark, ColHIN_NAME) = "品番未登録です！！"
                                        
                                        
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                            Unload Me
                                    End Select
                                End If
                            Case BtErrEOF
                                MsgBox "構成未登録です！！"
                                    
                                SE_USOU_HAKO(TDBGrid1.Bookmark, ColHIN_NAME) = "構成未登録です！！"
                    
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "構成マスタ")
                                Unload Me
                        End Select
                                                
                        Exit For
                        
                    Case BtErrKeyNotFound
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Unload Me
                        
                End Select
            
            Next i
                    
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Unload Me
            
    End Select
                
                
                
                
                
                
                
                
                
                
                
''    Err_Flg = False
''    For i = 1 To SE_USOU_HAKO.Count(1)
''
''        For j = 1 To SE_USOU_HAKO.Count(1)
''
''            If i <> j Then
''                If Trim(SE_USOU_HAKO(i, ColHIN_GAI)) = Trim(SE_USOU_HAKO(j, ColHIN_GAI)) Then
''                    Err_Flg = True
''                    Exit For
''                End If
''            End If
''        Next j
''        If Err_Flg Then
''            Exit For
''        End If
''    Next i
''
''    If Err_Flg Then
''
''        MsgBox "品番が重複しています！！"
''        SE_USOU_HAKO(TDBGrid1.Bookmark, ColHIN_NAME) = "重複品番です！！"
''    End If
        
    SE_USOU_HAKO(TDBGrid1.Bookmark, ColJGYOBU) = StrConv(ITEMREC.JGYOBU, vbUnicode)
    SE_USOU_HAKO(TDBGrid1.Bookmark, ColNAIGAI) = StrConv(ITEMREC.NAIGAI, vbUnicode)
        
        
    ReDim MAISU(ColMUKE To UBound(MUKE_TBL) + ColMUKE)
    ReDim SAISU(ColMUKE To UBound(MUKE_TBL) + ColMUKE)
        
        
    For i = 3 To SE_USOU_HAKO.UpperBound(1)
    
        For j = ColMUKE To UBound(MAISU)
    
    
    
    
            If IsNumeric(SE_USOU_HAKO(i, j)) Then
                MAISU(j) = MAISU(j) + CLng(SE_USOU_HAKO(i, j))
                GK_MAISU = GK_MAISU + CLng(SE_USOU_HAKO(i, j))
            End If
        
        Next j
    
    Next i
        
        
    For i = 3 To SE_USOU_HAKO.UpperBound(1)
    
        For j = ColMUKE To UBound(SAISU)
    
    
            If IsNumeric(SE_USOU_HAKO(i, j)) And IsNumeric(SE_USOU_HAKO(i, ColSAI_SU)) Then
                
                SAISU(j) = SAISU(j) + (CLng(SE_USOU_HAKO(i, j)) * CDbl(SE_USOU_HAKO(i, ColSAI_SU)))
                'GK_SAISU = GK_SAISU + CLng(SE_USOU_HAKO(i, j)) 2013.04.01
            End If
        
        Next j
    
    Next i
        
        
    For i = ColMUKE To UBound(MAISU)
        
        SE_USOU_HAKO(1, i) = MAISU(i)
        SE_USOU_HAKO(2, i) = SAISU(i)
    
    
    Next i
        
        
    Text1(ptxMAISU).Text = Format(GK_MAISU, "#,##0")
    'Text1(ptxSAISU).Text = Format(GK_MAISU, "#,##0.0")     '2013.04.01
    
    GK_SAISU = 0
    For j = ColMUKE To UBound(SAISU)                        '2013.04.01
        GK_SAISU = GK_SAISU + SAISU(j)                      '2013.04.01
    Next j                                                  '2013.04.01
    Text1(ptxSAISU).Text = Format(GK_SAISU, "#,##0.0")      '2013.04.01
        
        
        
    Set TDBGrid1.Array = SE_USOU_HAKO
        
    
    TDBGrid1.Refresh
    TDBGrid1.Update
'    TDBGrid1.ScrollBars = dbgAutomatic
    
    TDBGrid1.SetFocus


End Sub



Private Sub TDBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)


'    If TDBGrid1.Bookmark = 1 Or TDBGrid1.Bookmark = 2 Then
'        Cancel = True
'    End If

End Sub

Private Sub TDBGrid1_BeforeDelete(Cancel As Integer)

Dim sts As Integer
    
    If Not IsNumeric(TDBGrid1.Bookmark) Then
        Exit Sub
    End If
    
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, SE_USOU_HAKO(TDBGrid1.Bookmark, ColHIN_GAI))
        
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            
                If Trim(StrConv(ITEMREC.SE_USOU_F, vbUnicode)) <> "" Then
                    Cancel = True
                    Exit Sub
                End If
            
            
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Unload Me
        End Select


    




End Sub

Private Sub TDBGrid1_BeforeInsert(Cancel As Integer)
    
'    SE_USOU_HAKO.ReDim Min_Row, SE_USOU_HAKO.Count(1), Min_Col, Max_Col

End Sub


Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub
Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   輸送箱実績の内容を表示する
'----------------------------------------------------------------------------

Dim sts                 As Integer
Dim com                 As Integer

Dim i                   As Integer
Dim j                   As Integer


Dim Row                 As Long
    
Dim GK_MAISU            As Long
Dim GK_SAISU            As Long
    
Dim Skip_Flg            As Boolean
    
Dim wkSHIMUKE_CODE      As String * 2
    
    
Dim MAISU()             As Long
Dim SAISU()             As Long
    
    
    List_Disp_Proc = True
                                    
    Call Input_Lock
                                    
                                    'テーブルリセット
    Set SE_USOU_HAKO = Nothing
                                                    



    SE_USOU_HAKO.ReDim Min_Row, 1, Min_Col, Max_Col
    SE_USOU_HAKO(1, ColSE_USOU_F) = " "
    SE_USOU_HAKO(1, ColJGYOBU) = "*"
    SE_USOU_HAKO(1, ColNAIGAI) = "*"
    SE_USOU_HAKO(1, ColHIN_GAI) = "**********"
    SE_USOU_HAKO(1, ColHIN_NAME) = "[枚数合計]"
    SE_USOU_HAKO(1, ColSAI_SU) = "**.*"




    SE_USOU_HAKO.ReDim Min_Row, 2, Min_Col, Max_Col
    SE_USOU_HAKO(2, ColSE_USOU_F) = "!"
    SE_USOU_HAKO(2, ColJGYOBU) = "*"
    SE_USOU_HAKO(2, ColNAIGAI) = "*"
    SE_USOU_HAKO(2, ColHIN_GAI) = "**********"
    SE_USOU_HAKO(2, ColHIN_NAME) = "[才数合計]"
    SE_USOU_HAKO(2, ColSAI_SU) = "**.*"




    
                                                    
                                                    
    Call UniCode_Conv(K0_SE_USOU_HAKO.JITU_DATE, Format(Text1(ptxJITU_Date).Text, "YYYYMMDD"))
    Call UniCode_Conv(K0_SE_USOU_HAKO.JGYOBU, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.NAIGAI, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.HIN_GAI, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.MTS_CODE, "")
    
    Row = 2
        
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        
        sts = BTRV(com, SE_USOU_HAKO_POS, SE_USOU_HAKOREC, Len(SE_USOU_HAKOREC), K0_SE_USOU_HAKO, Len(K0_SE_USOU_HAKO), 0)
    
    
    
        Select Case sts
            Case BtNoErr
        
        
                If Format(Text1(ptxJITU_Date).Text, "YYYYMMDD") <> StrConv(SE_USOU_HAKOREC.JITU_DATE, vbUnicode) Then
                    Exit Do
                End If
                    
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "輸送箱実績")
                Exit Function
        End Select
            
            
            
            
                    
        If Grid_Set_Proc() Then
            Exit Function
        End If
        
        com = BtOpGetNext
        
    Loop
    
                                
    If SE_USOU_HAKO.Count(1) < 3 Then
                                '初期画面作成
        Call Init_Disp_Proc
    Else
                                    
                                    
                                    
                                    
        ReDim MAISU(ColMUKE To UBound(MUKE_TBL) + ColMUKE)
        ReDim SAISU(ColMUKE To UBound(MUKE_TBL) + ColMUKE)
            
            
        GK_MAISU = 0
        GK_SAISU = 0
            
            
        For i = 1 To SE_USOU_HAKO.UpperBound(1)
        
            For j = ColMUKE To UBound(MAISU)
        
        
                If IsNumeric(SE_USOU_HAKO(i, j)) Then
                    MAISU(j) = MAISU(j) + CLng(SE_USOU_HAKO(i, j))
                    GK_MAISU = GK_MAISU + CLng(SE_USOU_HAKO(i, j))
                End If
            
            Next j
        
        Next i
            
            
        For i = 1 To SE_USOU_HAKO.UpperBound(1)
        
            For j = ColMUKE To UBound(SAISU)
        
        
                If IsNumeric(SE_USOU_HAKO(i, j)) And IsNumeric(SE_USOU_HAKO(i, ColSAI_SU)) Then
                    
                    SAISU(j) = SAISU(j) + (CLng(SE_USOU_HAKO(i, j)) * CDbl(SE_USOU_HAKO(i, ColSAI_SU)))
                
                    'GK_SAISU = GK_SAISU + CLng(SE_USOU_HAKO(i, j)) 2013.04.01
                
                
                End If
            
            Next j
        
        Next i
            
            
        For i = ColMUKE To UBound(MAISU)
            
            SE_USOU_HAKO(1, i) = MAISU(i)
            SE_USOU_HAKO(2, i) = SAISU(i)
        
        
        Next i
                                    
                                    
                                    
                                    
                                    
                                    
                                    
        SE_USOU_HAKO.QuickSort Min_Row, SE_USOU_HAKO.UpperBound(1), ColSE_USOU_F, 0, XTYPE_STRING
                                    
                                    'DBテーブルリンク
        Set TDBGrid1.Array = SE_USOU_HAKO
        
        TDBGrid1.Bookmark = Null
        
        
        TDBGrid1.ReBind
        TDBGrid1.Update
        TDBGrid1.MoveFirst
        TDBGrid1.ScrollBars = dbgAutomatic
    
        Text1(ptxMAISU).Text = Format(GK_MAISU, "#,##0")
        'Text1(ptxSAISU).Text = Format(GK_MAISU, "#,##0.0")     '2013.04.01
        
        GK_SAISU = 0
        For j = ColMUKE To UBound(SAISU)                        '2013.04.01
            GK_SAISU = GK_SAISU + SAISU(j)                      '2013.04.01
        Next j                                                  '2013.04.01
        Text1(ptxSAISU).Text = Format(GK_SAISU, "#,##0.0")      '2013.04.01
        
   
        
    
    End If
    
    
    Call Input_UnLock
    
    TDBGrid1.SetFocus
    
    
    List_Disp_Proc = False

    
End Function
Private Function Grid_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   輸送箱データ-->Grid
'----------------------------------------------------------------------------

Dim sts As Integer
Dim i   As Long
Dim j   As Integer
    
    Grid_Set_Proc = True

    
    If SE_USOU_HAKO.Count(1) < 3 Then
        SE_USOU_HAKO.ReDim Min_Row, 3, Min_Col, Max_Col
        
        
        SE_USOU_HAKO(3, ColSE_USOU_F) = Trim(StrConv(SE_USOU_HAKOREC.SE_USOU_F, vbUnicode))
        
        
        SE_USOU_HAKO(3, ColHIN_GAI) = Trim(StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode))
            
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(SE_USOU_HAKOREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(SE_USOU_HAKOREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode))

        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Call UniCode_Conv(ITEMREC.SAI_SU, "00.0")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        End Select
        
        SE_USOU_HAKO(3, ColJGYOBU) = StrConv(SE_USOU_HAKOREC.JGYOBU, vbUnicode)
        SE_USOU_HAKO(3, ColNAIGAI) = StrConv(SE_USOU_HAKOREC.NAIGAI, vbUnicode)
        SE_USOU_HAKO(3, ColHIN_GAI) = StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode)
        
        SE_USOU_HAKO(3, ColHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        If IsNumeric(StrConv(ITEMREC.SAI_SU, vbUnicode)) Then
            SE_USOU_HAKO(3, ColSAI_SU) = Format(CDbl(StrConv(ITEMREC.SAI_SU, vbUnicode)), "0.0")
        Else
            SE_USOU_HAKO(3, ColSAI_SU) = "0.0"
        End If
            
            
    
    End If
    
    
    For i = 3 To SE_USOU_HAKO.Count(1)
        
        If SE_USOU_HAKO(i, ColJGYOBU) = StrConv(SE_USOU_HAKOREC.JGYOBU, vbUnicode) And _
            SE_USOU_HAKO(i, ColNAIGAI) = StrConv(SE_USOU_HAKOREC.NAIGAI, vbUnicode) And _
            SE_USOU_HAKO(i, ColHIN_GAI) = StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode) Then
                
            Exit For
    
        End If
    
    Next i
    
    If i > SE_USOU_HAKO.Count(1) Then
    
        SE_USOU_HAKO.ReDim Min_Row, i, Min_Col, Max_Col
        
        SE_USOU_HAKO(i, ColSE_USOU_F) = Trim(StrConv(SE_USOU_HAKOREC.SE_USOU_F, vbUnicode))
        SE_USOU_HAKO(i, ColJGYOBU) = StrConv(SE_USOU_HAKOREC.JGYOBU, vbUnicode)
        SE_USOU_HAKO(i, ColNAIGAI) = StrConv(SE_USOU_HAKOREC.NAIGAI, vbUnicode)
        SE_USOU_HAKO(i, ColHIN_GAI) = StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode)
    
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(SE_USOU_HAKOREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(SE_USOU_HAKOREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode))

        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Call UniCode_Conv(ITEMREC.SAI_SU, "00.0")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        End Select
        SE_USOU_HAKO(i, ColHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        If IsNumeric(StrConv(ITEMREC.SAI_SU, vbUnicode)) Then
            SE_USOU_HAKO(i, ColSAI_SU) = Format(CDbl(StrConv(ITEMREC.SAI_SU, vbUnicode)), "0.0")
        Else
            SE_USOU_HAKO(i, ColSAI_SU) = "0.0"
        End If
    
    
    
    
    End If
    
    For j = 0 To UBound(MUKE_TBL)
    
        If Trim(MUKE_TBL(j)) = "********" Then
            
            SE_USOU_HAKO(i, ColMUKE + j) = Format(CInt(StrConv(SE_USOU_HAKOREC.MAISU, vbUnicode)), "#")
            Exit For
        Else
            If MUKE_TBL(j) = StrConv(SE_USOU_HAKOREC.MTS_CODE, vbUnicode) Then
                
                SE_USOU_HAKO(i, ColMUKE + j) = Format(CLng(StrConv(SE_USOU_HAKOREC.MAISU, vbUnicode)), "#")
                Exit For
            End If
        End If
    
    Next j
        
    
    Grid_Set_Proc = False
End Function


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    If Error_Check_Proc(Index) Then     'エラーチェック
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts As Integer
    
    
    Error_Check_Proc = True
    
    Select Case Mode
    
    
        Case ptxTanto_Code     '担当者ｺｰﾄﾞ
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTanto_Code).Text)
            
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    Text1(ptxTanto_Name).Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    Text1(ptxTanto_Name).Text = ""
                    MsgBox "入力した項目はエラーです。(担当者)"
                    Text1(ptxTanto_Code).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Exit Function
            End Select
            
        Case ptxJITU_Date   '売上日付
            
            If Not IsDate(Text1(ptxJITU_Date).Text) Then
                MsgBox "入力した項目はエラーです。(売上日付)"
                Text1(ptxTanto_Code).SetFocus
                Exit Function
            End If
    
            If List_Disp_Proc() Then
                Exit Function
            End If
    End Select
        
        
    Error_Check_Proc = False
    

End Function


Private Function Grid_Error_Check_Proc() As Integer
'----------------------------------------------------------------------------
'                   ｸﾞﾘｯﾄﾞ入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts         As Integer
Dim i           As Integer
Dim j           As Integer
    
Dim Err_Flg     As Boolean
    
    
    Grid_Error_Check_Proc = True
    
    Set TDBGrid1.Array = SE_USOU_HAKO
'    TDBGrid1.Refresh
'    TDBGrid1.Update
    
    Err_Flg = False
    For i = 1 To SE_USOU_HAKO.Count(1)
            
        If Trim(SE_USOU_HAKO(i, ColHIN_GAI)) = "" Or Trim(SE_USOU_HAKO(i, ColHIN_GAI)) = "**********" Then
        Else
            
            Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, SE_USOU_HAKO(i, ColHIN_GAI))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    
            Select Case sts
                Case BtNoErr
                    
                    SE_USOU_HAKO(TDBGrid1.Bookmark, ColHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                    If IsNumeric(StrConv(ITEMREC.SAI_SU, vbUnicode)) Then
                        SE_USOU_HAKO(i, ColSAI_SU) = Format(CDbl(StrConv(ITEMREC.SAI_SU, vbUnicode)), "#0.0")
                    Else
                        SE_USOU_HAKO(i, ColSAI_SU) = "0.0"
                    End If
                
                
                Case BtErrKeyNotFound
                
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, SE_USOU_HAKO(i, ColHIN_GAI))
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        
                    Select Case sts
                        Case BtNoErr
                        
                            Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
                            Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                            Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                            Call UniCode_Conv(K0_P_COMPO.HIN_GAI, SE_USOU_HAKO(i, ColHIN_GAI))
                            Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_GAISOU)
                            Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
                                                    
                            sts = BTRV(BtOpGetGreater, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                    
                            Select Case sts
                                Case BtNoErr
                                    
                                    If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) Or _
                                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1) Or _
                                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1) Or _
                                        Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> SE_USOU_HAKO(i, ColHIN_GAI) Or _
                                        StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_GAISOU Then
                                    Else
                                        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                                        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                        Select Case sts
                                            Case BtNoErr
                                                SE_USOU_HAKO(i, ColHIN_GAI) = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                                                SE_USOU_HAKO(i, ColHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                                If IsNumeric(StrConv(ITEMREC.SAI_SU, vbUnicode)) Then
                                                    SE_USOU_HAKO(i, ColSAI_SU) = Format(CDbl(StrConv(ITEMREC.SAI_SU, vbUnicode)), "#0.0")
                                                Else
                                                    SE_USOU_HAKO(i, ColSAI_SU) = "0.0"
                                                End If
                                
                                
                                            
                                            Case BtErrKeyNotFound
                                            
                                                MsgBox "品番未登録です！！"
                                                SE_USOU_HAKO(i, ColHIN_NAME) = "品番未登録です！！"
                                                Err_Flg = True
                                                Exit For
                                            
                                            Case Else
                                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                                Unload Me
                                        End Select
                                    End If
                                Case BtErrEOF
                                    MsgBox "品番未登録です！！"
                                    SE_USOU_HAKO(i, ColHIN_NAME) = "品番未登録です！！"
                                    Err_Flg = True
                                    Exit For
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "構成マスタ")
                                    Unload Me
                            End Select
                    
                        Case BtErrKeyNotFound
                            MsgBox "品番未登録です！！"
                            SE_USOU_HAKO(i, ColHIN_NAME) = "品番未登録です！！"
                            Err_Flg = True
                            Exit For
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Unload Me
                    
                    End Select
                Case BtErrKeyNotFound
                    
                    MsgBox "品番未登録です！！"
                    SE_USOU_HAKO(i, ColHIN_NAME) = "品番未登録です！！"
                    Exit For
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Unload Me
            
            End Select
                
            SE_USOU_HAKO(i, ColJGYOBU) = StrConv(ITEMREC.JGYOBU, vbUnicode)
            SE_USOU_HAKO(i, ColNAIGAI) = StrConv(ITEMREC.NAIGAI, vbUnicode)
            
            
        End If
            
    Next i
            
            
    
''    If Not Err_Flg Then
''        For i = 1 To SE_USOU_HAKO.Count(1)
''
''            For j = 1 To SE_USOU_HAKO.Count(1)
''
''                If i <> j Then
''                    If Trim(SE_USOU_HAKO(i, ColHIN_GAI)) = Trim(SE_USOU_HAKO(j, ColHIN_GAI)) Then
''                        Err_Flg = True
''                        Exit For
''                    End If
''                End If
''            Next j
''            If Err_Flg Then
''                Exit For
''            End If
''        Next i
''
''        If Err_Flg Then
''
''            MsgBox "品番が重複しています！！"
''            SE_USOU_HAKO(TDBGrid1.Bookmark, ColHIN_NAME) = "重複品番です！！"
''
''        End If
''
''
''
''    End If
        
        
    If Not Err_Flg Then
            
        For i = 1 To SE_USOU_HAKO.Count(1)
            
            For j = 0 To UBound(MUKE_TBL)
            
                If Not IsNumeric(SE_USOU_HAKO(i, j + ColMUKE)) Then
                    SE_USOU_HAKO(i, j + ColMUKE) = ""
                Else
                    SE_USOU_HAKO(i, j + ColMUKE) = Format(CLng(SE_USOU_HAKO(i, j + ColMUKE)), "#")
                End If
            Next j
        
        Next i
            
    End If
        
    
    TDBGrid1.SetFocus
    
    
    
    
        
        
    Grid_Error_Check_Proc = False
    

End Function




Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   データ更新
'----------------------------------------------------------------------------
Dim sts                 As Integer
    
Dim i                   As Integer
Dim j                   As Integer
Dim k                   As Integer
    
Dim com                 As Integer
    
Dim CHANGE_Flg          As Boolean
    
Dim wkSHIMUKE_CODE      As String * 2
    
    
    Update_Proc = True
                                     
    Set TDBGrid1.Array = SE_USOU_HAKO
'    TDBGrid1.Refresh
    
    TDBGrid1.Update
                                     
    If SE_USOU_HAKO.Count(1) < 3 Then
        Update_Proc = False
        Exit Function
    End If
                                     
                                     
    For i = 1 To SE_USOU_HAKO.Count(1)
        If Trim(SE_USOU_HAKO(i, ColHIN_GAI)) = "" Or Trim(SE_USOU_HAKO(i, ColHIN_GAI)) = "**********" Then
        Else
            
            For j = 1 To SE_USOU_HAKO.Count(1)
            
            
                If i = j Then
                Else
                    If Trim(SE_USOU_HAKO(i, ColHIN_GAI)) = Trim(SE_USOU_HAKO(j, ColHIN_GAI)) Then
                        For k = 0 To UBound(MUKE_TBL)
                            If Not IsNumeric(SE_USOU_HAKO(i, k + ColMUKE)) Then
                                SE_USOU_HAKO(i, k + ColMUKE) = "0"
                            End If
                            If Not IsNumeric(SE_USOU_HAKO(j, k + ColMUKE)) Then
                                SE_USOU_HAKO(j, k + ColMUKE) = "0"
                            End If
                            SE_USOU_HAKO(i, k + ColMUKE) = CLng(SE_USOU_HAKO(i, k + ColMUKE)) + CLng(SE_USOU_HAKO(j, k + ColMUKE))
                            SE_USOU_HAKO(j, k + ColMUKE) = 0
                        Next k
                        
                        SE_USOU_HAKO(j, ColHIN_GAI) = ""
                    End If
                End If
            
            Next j
        End If
    Next i
                                     
                                     
                                     
                                     
                                     
                                     
                                     
                                     
                                     'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, SE_USOU_HAKO_POS, SE_USOU_HAKOREC, Len(SE_USOU_HAKOREC), K0_SE_USOU_HAKO, Len(K0_SE_USOU_HAKO), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
                                    
                                    
                                    
    Call UniCode_Conv(K0_SE_USOU_HAKO.JITU_DATE, Format(Text1(ptxJITU_Date).Text, "YYYYMMDD"))
    Call UniCode_Conv(K0_SE_USOU_HAKO.JGYOBU, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.NAIGAI, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.HIN_GAI, "")
    
    com = BtOpGetGreater
    
    Do
        DoEvents
    
        sts = BTRV(com, SE_USOU_HAKO_POS, SE_USOU_HAKOREC, Len(SE_USOU_HAKOREC), K0_SE_USOU_HAKO, Len(K0_SE_USOU_HAKO), 0)
        Select Case sts
            Case BtNoErr
            
                If StrConv(SE_USOU_HAKOREC.JITU_DATE, vbUnicode) <> Format(Text1(ptxJITU_Date).Text, "YYYYMMDD") Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "輸送箱実績")
                Exit Function
        End Select
    
        For i = 1 To SE_USOU_HAKO.Count(1)
        
            If StrConv(SE_USOU_HAKOREC.JGYOBU, vbUnicode) = SE_USOU_HAKO(i, ColJGYOBU) And _
                StrConv(SE_USOU_HAKOREC.NAIGAI, vbUnicode) = SE_USOU_HAKO(i, ColNAIGAI) And _
                StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode) = SE_USOU_HAKO(i, ColHIN_GAI) Then

        
                Exit For
        
            End If
        
        
        Next i
    
    
        If i > SE_USOU_HAKO.Count(1) Then
    
            sts = BTRV(BtOpDelete, SE_USOU_HAKO_POS, SE_USOU_HAKOREC, Len(SE_USOU_HAKOREC), K0_SE_USOU_HAKO, Len(K0_SE_USOU_HAKO), 0)
            Select Case sts
                Case BtNoErr
                Case Else
                    Call File_Error(sts, BtOpDelete, "輸送箱実績")
                    Exit Function
            End Select
        
        End If
    
    
        com = BtOpGetNext
    
    
    Loop
                                    
                                    
                                    
    For i = 1 To SE_USOU_HAKO.Count(1)
                                    
        
        If Trim(SE_USOU_HAKO(i, ColHIN_GAI)) = "" Or Trim(SE_USOU_HAKO(i, ColHIN_GAI)) = "**********" Then
        Else
                                    
            For j = 0 To UBound(SE_JGYOBU_T)
            
                If Trim(SE_USOU_HAKO(i, ColJGYOBU)) = SE_JGYOBU_T(j).JGYOBU And _
                    Trim(SE_USOU_HAKO(i, ColNAIGAI)) = SE_JGYOBU_T(j).NAIGAI Then
            
                    wkSHIMUKE_CODE = SE_JGYOBU_T(j).SHIMUKE_CODE
                    Exit For
                End If
            Next j
                                    
            Call UniCode_Conv(K0_SE_USOU_HAKO.JITU_DATE, Format(Text1(ptxJITU_Date).Text, "YYYYMMDD"))
            Call UniCode_Conv(K0_SE_USOU_HAKO.JGYOBU, SE_USOU_HAKO(i, ColJGYOBU))
            Call UniCode_Conv(K0_SE_USOU_HAKO.NAIGAI, SE_USOU_HAKO(i, ColNAIGAI))
            Call UniCode_Conv(K0_SE_USOU_HAKO.HIN_GAI, SE_USOU_HAKO(i, ColHIN_GAI))
                                                
                                                
            For j = 0 To UBound(MUKE_TBL)
            
                Call UniCode_Conv(K0_SE_USOU_HAKO.MTS_CODE, MUKE_TBL(j))
            
                sts = BTRV(BtOpGetEqual, SE_USOU_HAKO_POS, SE_USOU_HAKOREC, Len(SE_USOU_HAKOREC), K0_SE_USOU_HAKO, Len(K0_SE_USOU_HAKO), 0)
                Select Case sts
                    Case BtNoErr
                        com = BtOpUpdate
                    Case BtErrKeyNotFound
                        com = BtOpInsert
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "輸送箱実績")
                        Exit Function
                End Select
            
                If com = BtOpInsert Then
                
                    Call UniCode_Conv(SE_USOU_HAKOREC.SHIMUKE_CODE, wkSHIMUKE_CODE)
                    Call UniCode_Conv(SE_USOU_HAKOREC.JITU_DATE, Format(Text1(ptxJITU_Date).Text, "YYYYMMDD"))
                    Call UniCode_Conv(SE_USOU_HAKOREC.JGYOBU, SE_USOU_HAKO(i, ColJGYOBU))
                    Call UniCode_Conv(SE_USOU_HAKOREC.NAIGAI, SE_USOU_HAKO(i, ColNAIGAI))
                    Call UniCode_Conv(SE_USOU_HAKOREC.HIN_GAI, SE_USOU_HAKO(i, ColHIN_GAI))
                
                    Call UniCode_Conv(SE_USOU_HAKOREC.MTS_CODE, MUKE_TBL(j))
                
                    Call UniCode_Conv(SE_USOU_HAKOREC.CYU_KBN, "")
                    Call UniCode_Conv(SE_USOU_HAKOREC.CYOK_KBN, "")
                
                
                    If Not IsNumeric(SE_USOU_HAKO(i, ColMUKE + j)) Then
                        Call UniCode_Conv(SE_USOU_HAKOREC.MAISU, "000000")
                    Else
                        Call UniCode_Conv(SE_USOU_HAKOREC.MAISU, Format(CLng(SE_USOU_HAKO(i, ColMUKE + j)), "000000"))
                    End If
                
                    Call UniCode_Conv(SE_USOU_HAKOREC.FILLER, "")
                                
                    Call UniCode_Conv(SE_USOU_HAKOREC.UPD_TANTO, Text1(ptxTanto_Code).Text)
                    Call UniCode_Conv(SE_USOU_HAKOREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                            
                Else
                
                    If Not IsNumeric(SE_USOU_HAKO(i, ColMUKE + j)) Then
                        SE_USOU_HAKO(i, ColMUKE + j) = "0"
                    End If
                
                    
                    If CLng(StrConv(SE_USOU_HAKOREC.MAISU, vbUnicode)) <> CLng(SE_USOU_HAKO(i, ColMUKE + j)) Then
                        Call UniCode_Conv(SE_USOU_HAKOREC.UPD_TANTO, Text1(ptxTanto_Code).Text)
                        Call UniCode_Conv(SE_USOU_HAKOREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                    End If
                
                
                    If Not IsNumeric(SE_USOU_HAKO(i, ColMUKE + j)) Then
                        Call UniCode_Conv(SE_USOU_HAKOREC.MAISU, "000000")
                    Else
                        Call UniCode_Conv(SE_USOU_HAKOREC.MAISU, Format(CLng(SE_USOU_HAKO(i, ColMUKE + j)), "000000"))
                    End If
                
                
                End If
                
                                        
                Call UniCode_Conv(K0_ITEM.JGYOBU, SE_USOU_HAKO(i, ColJGYOBU))
                Call UniCode_Conv(K0_ITEM.NAIGAI, SE_USOU_HAKO(i, ColNAIGAI))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, SE_USOU_HAKO(i, ColHIN_GAI))
                                        
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        If Trim(StrConv(ITEMREC.SE_USOU_F, vbUnicode)) = "" Then
                            Call UniCode_Conv(ITEMREC.SE_USOU_F, "zz")
                        End If
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(ITEMREC.SE_USOU_F, "zz")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                End Select
                                        
                Call UniCode_Conv(SE_USOU_HAKOREC.SE_USOU_F, StrConv(ITEMREC.SE_USOU_F, vbUnicode))
                                        
                                        
            
                sts = BTRV(com, SE_USOU_HAKO_POS, SE_USOU_HAKOREC, Len(SE_USOU_HAKOREC), K0_SE_USOU_HAKO, Len(K0_SE_USOU_HAKO), 0)
                Select Case sts
                    Case BtNoErr
                    Case Else
                        Call File_Error(sts, com, "輸送箱実績")
                        Exit Function
                End Select
            
            Next j
        End If
    Next i
                                    
                                    
                                    
                                        
                                        
End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, SE_USOU_HAKO_POS, SE_USOU_HAKOREC, Len(SE_USOU_HAKOREC), K0_SE_USOU_HAKO, Len(K0_SE_USOU_HAKO), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, SE_USOU_HAKO_POS, SE_USOU_HAKOREC, Len(SE_USOU_HAKOREC), K0_SE_USOU_HAKO, Len(K0_SE_USOU_HAKO), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If


End Function

