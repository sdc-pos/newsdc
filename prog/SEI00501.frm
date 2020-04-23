VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form SEI00501 
   Caption         =   "[請求システム]輸送箱日別実績作成処理[SEI0050] 2012.09.26 16:00"
   ClientHeight    =   11145
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   16020
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
   ScaleHeight     =   11145
   ScaleWidth      =   16020
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Command1 
      Caption         =   "終  了"
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
      Left            =   4410
      TabIndex        =   4
      Top             =   120
      Width           =   1905
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   3150
      TabIndex        =   1
      Top             =   960
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1470
      TabIndex        =   0
      Top             =   960
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXCEL"
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
      Left            =   2310
      TabIndex        =   3
      Top             =   120
      Width           =   1905
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   6
      Top             =   10320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton Command1 
      Caption         =   "表  示"
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
      TabIndex        =   2
      Top             =   120
      Width           =   1905
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   7935
      Left            =   315
      TabIndex        =   5
      Top             =   1800
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   13996
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "品番"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "品名"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "才数"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "単価"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
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
      Columns(36)._VlistStyle=   0
      Columns(36)._MaxComboItems=   5
      Columns(36).DataField=   ""
      Columns(36)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(37)._VlistStyle=   0
      Columns(37)._MaxComboItems=   5
      Columns(37).DataField=   ""
      Columns(37)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   38
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=38"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2170"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2037"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=4366"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=4233"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1799"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1667"
      Splits(0)._ColumnProps(12)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=2646"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2514"
      Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=2170"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=2037"
      Splits(0)._ColumnProps(22)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(24)=   "Column(5).Width=2170"
      Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=2037"
      Splits(0)._ColumnProps(27)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(29)=   "Column(6).Width=2170"
      Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=2037"
      Splits(0)._ColumnProps(32)=   "Column(6)._ColStyle=1"
      Splits(0)._ColumnProps(33)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(34)=   "Column(7).Width=2170"
      Splits(0)._ColumnProps(35)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(7)._WidthInPix=2037"
      Splits(0)._ColumnProps(37)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(38)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(39)=   "Column(8).Width=2170"
      Splits(0)._ColumnProps(40)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(8)._WidthInPix=2037"
      Splits(0)._ColumnProps(42)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(43)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(44)=   "Column(9).Width=2170"
      Splits(0)._ColumnProps(45)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(9)._WidthInPix=2037"
      Splits(0)._ColumnProps(47)=   "Column(9)._ColStyle=2"
      Splits(0)._ColumnProps(48)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(49)=   "Column(10).Width=2170"
      Splits(0)._ColumnProps(50)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(10)._WidthInPix=2037"
      Splits(0)._ColumnProps(52)=   "Column(10)._ColStyle=2"
      Splits(0)._ColumnProps(53)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(54)=   "Column(11).Width=2170"
      Splits(0)._ColumnProps(55)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(11)._WidthInPix=2037"
      Splits(0)._ColumnProps(57)=   "Column(11)._ColStyle=2"
      Splits(0)._ColumnProps(58)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(59)=   "Column(12).Width=2170"
      Splits(0)._ColumnProps(60)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(12)._WidthInPix=2037"
      Splits(0)._ColumnProps(62)=   "Column(12)._ColStyle=2"
      Splits(0)._ColumnProps(63)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(64)=   "Column(13).Width=2170"
      Splits(0)._ColumnProps(65)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(66)=   "Column(13)._WidthInPix=2037"
      Splits(0)._ColumnProps(67)=   "Column(13)._ColStyle=2"
      Splits(0)._ColumnProps(68)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(69)=   "Column(14).Width=2170"
      Splits(0)._ColumnProps(70)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(14)._WidthInPix=2037"
      Splits(0)._ColumnProps(72)=   "Column(14)._ColStyle=2"
      Splits(0)._ColumnProps(73)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(74)=   "Column(15).Width=2170"
      Splits(0)._ColumnProps(75)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(76)=   "Column(15)._WidthInPix=2037"
      Splits(0)._ColumnProps(77)=   "Column(15)._ColStyle=2"
      Splits(0)._ColumnProps(78)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(79)=   "Column(16).Width=2170"
      Splits(0)._ColumnProps(80)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(81)=   "Column(16)._WidthInPix=2037"
      Splits(0)._ColumnProps(82)=   "Column(16)._ColStyle=2"
      Splits(0)._ColumnProps(83)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(84)=   "Column(17).Width=2170"
      Splits(0)._ColumnProps(85)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(86)=   "Column(17)._WidthInPix=2037"
      Splits(0)._ColumnProps(87)=   "Column(17)._ColStyle=2"
      Splits(0)._ColumnProps(88)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(89)=   "Column(18).Width=2170"
      Splits(0)._ColumnProps(90)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(18)._WidthInPix=2037"
      Splits(0)._ColumnProps(92)=   "Column(18)._ColStyle=2"
      Splits(0)._ColumnProps(93)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(94)=   "Column(19).Width=2170"
      Splits(0)._ColumnProps(95)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(96)=   "Column(19)._WidthInPix=2037"
      Splits(0)._ColumnProps(97)=   "Column(19)._ColStyle=2"
      Splits(0)._ColumnProps(98)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(99)=   "Column(20).Width=2170"
      Splits(0)._ColumnProps(100)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(101)=   "Column(20)._WidthInPix=2037"
      Splits(0)._ColumnProps(102)=   "Column(20)._ColStyle=2"
      Splits(0)._ColumnProps(103)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(104)=   "Column(21).Width=2170"
      Splits(0)._ColumnProps(105)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(106)=   "Column(21)._WidthInPix=2037"
      Splits(0)._ColumnProps(107)=   "Column(21)._ColStyle=2"
      Splits(0)._ColumnProps(108)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(109)=   "Column(22).Width=2170"
      Splits(0)._ColumnProps(110)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(111)=   "Column(22)._WidthInPix=2037"
      Splits(0)._ColumnProps(112)=   "Column(22)._ColStyle=2"
      Splits(0)._ColumnProps(113)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(114)=   "Column(23).Width=2170"
      Splits(0)._ColumnProps(115)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(116)=   "Column(23)._WidthInPix=2037"
      Splits(0)._ColumnProps(117)=   "Column(23)._ColStyle=2"
      Splits(0)._ColumnProps(118)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(119)=   "Column(24).Width=2170"
      Splits(0)._ColumnProps(120)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(121)=   "Column(24)._WidthInPix=2037"
      Splits(0)._ColumnProps(122)=   "Column(24)._ColStyle=2"
      Splits(0)._ColumnProps(123)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(124)=   "Column(25).Width=2170"
      Splits(0)._ColumnProps(125)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(126)=   "Column(25)._WidthInPix=2037"
      Splits(0)._ColumnProps(127)=   "Column(25)._ColStyle=2"
      Splits(0)._ColumnProps(128)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(129)=   "Column(26).Width=2170"
      Splits(0)._ColumnProps(130)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(131)=   "Column(26)._WidthInPix=2037"
      Splits(0)._ColumnProps(132)=   "Column(26)._ColStyle=2"
      Splits(0)._ColumnProps(133)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(134)=   "Column(27).Width=2170"
      Splits(0)._ColumnProps(135)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(136)=   "Column(27)._WidthInPix=2037"
      Splits(0)._ColumnProps(137)=   "Column(27)._ColStyle=2"
      Splits(0)._ColumnProps(138)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(139)=   "Column(28).Width=2170"
      Splits(0)._ColumnProps(140)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(141)=   "Column(28)._WidthInPix=2037"
      Splits(0)._ColumnProps(142)=   "Column(28)._ColStyle=2"
      Splits(0)._ColumnProps(143)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(144)=   "Column(29).Width=2170"
      Splits(0)._ColumnProps(145)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(146)=   "Column(29)._WidthInPix=2037"
      Splits(0)._ColumnProps(147)=   "Column(29)._ColStyle=2"
      Splits(0)._ColumnProps(148)=   "Column(29).Order=30"
      Splits(0)._ColumnProps(149)=   "Column(30).Width=2170"
      Splits(0)._ColumnProps(150)=   "Column(30).DividerColor=0"
      Splits(0)._ColumnProps(151)=   "Column(30)._WidthInPix=2037"
      Splits(0)._ColumnProps(152)=   "Column(30)._ColStyle=2"
      Splits(0)._ColumnProps(153)=   "Column(30).Order=31"
      Splits(0)._ColumnProps(154)=   "Column(31).Width=2170"
      Splits(0)._ColumnProps(155)=   "Column(31).DividerColor=0"
      Splits(0)._ColumnProps(156)=   "Column(31)._WidthInPix=2037"
      Splits(0)._ColumnProps(157)=   "Column(31)._ColStyle=2"
      Splits(0)._ColumnProps(158)=   "Column(31).Order=32"
      Splits(0)._ColumnProps(159)=   "Column(32).Width=2170"
      Splits(0)._ColumnProps(160)=   "Column(32).DividerColor=0"
      Splits(0)._ColumnProps(161)=   "Column(32)._WidthInPix=2037"
      Splits(0)._ColumnProps(162)=   "Column(32)._ColStyle=2"
      Splits(0)._ColumnProps(163)=   "Column(32).Order=33"
      Splits(0)._ColumnProps(164)=   "Column(33).Width=2170"
      Splits(0)._ColumnProps(165)=   "Column(33).DividerColor=0"
      Splits(0)._ColumnProps(166)=   "Column(33)._WidthInPix=2037"
      Splits(0)._ColumnProps(167)=   "Column(33)._ColStyle=2"
      Splits(0)._ColumnProps(168)=   "Column(33).Order=34"
      Splits(0)._ColumnProps(169)=   "Column(34).Width=2170"
      Splits(0)._ColumnProps(170)=   "Column(34).DividerColor=0"
      Splits(0)._ColumnProps(171)=   "Column(34)._WidthInPix=2037"
      Splits(0)._ColumnProps(172)=   "Column(34).Order=35"
      Splits(0)._ColumnProps(173)=   "Column(35).Width=4366"
      Splits(0)._ColumnProps(174)=   "Column(35).DividerColor=0"
      Splits(0)._ColumnProps(175)=   "Column(35)._WidthInPix=4233"
      Splits(0)._ColumnProps(176)=   "Column(35).Visible=0"
      Splits(0)._ColumnProps(177)=   "Column(35).Order=36"
      Splits(0)._ColumnProps(178)=   "Column(36).Width=4366"
      Splits(0)._ColumnProps(179)=   "Column(36).DividerColor=0"
      Splits(0)._ColumnProps(180)=   "Column(36)._WidthInPix=4233"
      Splits(0)._ColumnProps(181)=   "Column(36).Visible=0"
      Splits(0)._ColumnProps(182)=   "Column(36).Order=37"
      Splits(0)._ColumnProps(183)=   "Column(37).Width=4366"
      Splits(0)._ColumnProps(184)=   "Column(37).DividerColor=0"
      Splits(0)._ColumnProps(185)=   "Column(37)._WidthInPix=4233"
      Splits(0)._ColumnProps(186)=   "Column(37).Visible=0"
      Splits(0)._ColumnProps(187)=   "Column(37).Order=38"
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
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "箱数使用実績"
      AllowArrows     =   0   'False
      MultipleLines   =   0
      EmptyRows       =   -1  'True
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
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=102,.parent=87,.alignment=1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=99,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=100,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=101,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=114,.parent=87,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=111,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=112,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=113,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=16,.parent=87,.alignment=1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=13,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=14,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=15,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=20,.parent=87,.alignment=1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=17,.parent=88"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=18,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=19,.parent=91"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=24,.parent=87,.alignment=2"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=21,.parent=88"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=22,.parent=89"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=23,.parent=91"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=28,.parent=87,.alignment=1"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=88"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=89"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=91"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=32,.parent=87,.alignment=1"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=29,.parent=88"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=30,.parent=89"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=31,.parent=91"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=46,.parent=87,.alignment=1"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=43,.parent=88"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=44,.parent=89"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=45,.parent=91"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=50,.parent=87,.alignment=1"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=47,.parent=88"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=48,.parent=89"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=49,.parent=91"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=54,.parent=87,.alignment=1"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=51,.parent=88"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=52,.parent=89"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=53,.parent=91"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=58,.parent=87,.alignment=1"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=55,.parent=88"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=56,.parent=89"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=57,.parent=91"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=62,.parent=87,.alignment=1"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=59,.parent=88"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=60,.parent=89"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=61,.parent=91"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=66,.parent=87,.alignment=1"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=63,.parent=88"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=64,.parent=89"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=65,.parent=91"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=70,.parent=87,.alignment=1"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=67,.parent=88"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=68,.parent=89"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=69,.parent=91"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=74,.parent=87,.alignment=1"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=71,.parent=88"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=72,.parent=89"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=73,.parent=91"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=78,.parent=87,.alignment=1"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=75,.parent=88"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=76,.parent=89"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=77,.parent=91"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=82,.parent=87,.alignment=1"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=79,.parent=88"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=80,.parent=89"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=81,.parent=91"
      _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=86,.parent=87,.alignment=1"
      _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=83,.parent=88"
      _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=84,.parent=89"
      _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=85,.parent=91"
      _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=118,.parent=87,.alignment=1"
      _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=115,.parent=88"
      _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=116,.parent=89"
      _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=117,.parent=91"
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
      _StyleDefs(172) =   "Splits(0).Columns(34).Style:id=174,.parent=87"
      _StyleDefs(173) =   "Splits(0).Columns(34).HeadingStyle:id=171,.parent=88"
      _StyleDefs(174) =   "Splits(0).Columns(34).FooterStyle:id=172,.parent=89"
      _StyleDefs(175) =   "Splits(0).Columns(34).EditorStyle:id=173,.parent=91"
      _StyleDefs(176) =   "Splits(0).Columns(35).Style:id=246,.parent=87"
      _StyleDefs(177) =   "Splits(0).Columns(35).HeadingStyle:id=243,.parent=88"
      _StyleDefs(178) =   "Splits(0).Columns(35).FooterStyle:id=244,.parent=89"
      _StyleDefs(179) =   "Splits(0).Columns(35).EditorStyle:id=245,.parent=91"
      _StyleDefs(180) =   "Splits(0).Columns(36).Style:id=254,.parent=87"
      _StyleDefs(181) =   "Splits(0).Columns(36).HeadingStyle:id=251,.parent=88"
      _StyleDefs(182) =   "Splits(0).Columns(36).FooterStyle:id=252,.parent=89"
      _StyleDefs(183) =   "Splits(0).Columns(36).EditorStyle:id=253,.parent=91"
      _StyleDefs(184) =   "Splits(0).Columns(37).Style:id=258,.parent=87"
      _StyleDefs(185) =   "Splits(0).Columns(37).HeadingStyle:id=255,.parent=88"
      _StyleDefs(186) =   "Splits(0).Columns(37).FooterStyle:id=256,.parent=89"
      _StyleDefs(187) =   "Splits(0).Columns(37).EditorStyle:id=257,.parent=91"
      _StyleDefs(188) =   "Named:id=33:Normal"
      _StyleDefs(189) =   ":id=33,.parent=0"
      _StyleDefs(190) =   "Named:id=34:Heading"
      _StyleDefs(191) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(192) =   ":id=34,.wraptext=-1"
      _StyleDefs(193) =   "Named:id=35:Footing"
      _StyleDefs(194) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(195) =   "Named:id=36:Selected"
      _StyleDefs(196) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(197) =   "Named:id=37:Caption"
      _StyleDefs(198) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(199) =   "Named:id=38:HighlightRow"
      _StyleDefs(200) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(201) =   "Named:id=39:EvenRow"
      _StyleDefs(202) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(203) =   "Named:id=40:OddRow"
      _StyleDefs(204) =   ":id=40,.parent=33"
      _StyleDefs(205) =   "Named:id=41:RecordSelector"
      _StyleDefs(206) =   ":id=41,.parent=34"
      _StyleDefs(207) =   "Named:id=42:FilterBar"
      _StyleDefs(208) =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '実線
      Caption         =   "～"
      Height          =   375
      Index           =   8
      Left            =   2835
      TabIndex        =   8
      Top             =   960
      Width           =   330
   End
   Begin VB.Label Label1 
      Caption         =   "日付範囲"
      Height          =   375
      Index           =   7
      Left            =   315
      TabIndex        =   7
      Top             =   1080
      Width           =   1170
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "表示"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "EXCEL"
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
Attribute VB_Name = "SEI00501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Const ptxS_JITU_DATE% = 0       '日付範囲　開始
Private Const ptxE_JITU_DATE% = 1       '日付範囲　開始


Dim SE_USOU_HAKO    As New XArrayDB

Private Const Min_Row% = 1              '最小行数

Dim Max_Row         As Integer          'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 34             '最大列数

Private Const ColHIN_GAI% = 0           '品番
Private Const ColHIN_NAME% = 1          '品名
Private Const ColSAI_SU% = 2            '才数
Private Const ColTANKA% = 3             '単価


Private Const ColDay% = 4               '日別　数量／才数 4～34

'--------------------------------------- EXCEL用定数    2012.09.26
Private Const xlDot% = -4118
Private Const xlCalculationManual% = -4135
Private Const xlLeft% = -4131
Private Const xlCenter% = -4108
Private Const xlBottom% = -4107
Private Const xlNone% = -4142
Private Const xlContinuous% = 1
Private Const xlThin% = 2
Private Const xlAutomatic% = -4105
Private Const xlRight% = -4152
Private Const xlDiagonalDown% = 5
Private Const xlDiagonalUp% = 6
Private Const xlEdgeLeft% = 7
Private Const xlEdgeTop% = 8
Private Const xlEdgeBottom% = 9
Private Const xlEdgeRight% = 10
Private Const xlInsideVertical% = 11
Private Const xlInsideHorizontal% = 12
Private Const xlThick% = 4
Private Const xlCalculationAutomatic% = -4105
Private Const xlPortrait% = 1
'--------------------------------------- EXCEL用定数



'Private Const EXCEL_OBJECT_NAME As String = "Excel.Application"    2012.09.26







Private Sub Command1_Click(Index As Integer)

Dim ans As Integer
Dim i   As Integer



    Select Case Index
        Case 0                              '再表示
            
            For i = ptxS_JITU_DATE To ptxE_JITU_DATE
                If Error_Check_Proc(i) Then
                    Exit Sub
                End If
            Next i
            
            
            
            If List_Disp_Proc Then
                Unload Me
            End If
        
        Case 1                              'データ出力
        
            For i = ptxS_JITU_DATE To ptxE_JITU_DATE
                If Error_Check_Proc(i) Then
                    Exit Sub
                End If
            Next i
            ans = MsgBox("輸送箱出荷個数(日別)作成しますか？", vbYesNo + vbQuestion, "確認入力")
            
            
            If ans = vbYes Then
                If DETAIL_Proc() Then
                    Unload Me
                End If
            End If
        
        
        Case 2                              '終了
            Unload Me
    End Select
    

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
Private Sub Form_Load()
Dim i           As Integer
Dim c           As String * 128
Dim sts         As Integer
    
    
Dim S_DATE      As String
Dim E_DATE      As String
Dim S_YY        As String * 4
Dim S_MM        As String * 2
Dim S_DD        As String * 2
    
    
    
    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If


    
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]輸送箱日別実績作成処理", Me.hwnd, 0)
    
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SETPARTS, 0, -1)



    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                


    Max_Row = 9999
                                

                                '倉庫マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '管理マスタＯＰＥＮ
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '輸送箱実績ＯＰＥＮ
    If SE_USOU_HAKO_Open(BtOpenNomal) Then
        Unload Me
    End If






    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
            
        Case Else
            Unload Me
    End Select



    E_DATE = Format(Now, "YYYY/MM/DD")
    S_DATE = DateAdd("m", -1, E_DATE)
    S_DD = Right(S_DATE, 2)
    S_DD = Format(CInt(S_DD) + 1, "00")
    
    S_DATE = Left(S_DATE, 7) & "/" & S_DD
    If IsDate(S_DATE) Then
    Else
        S_MM = Mid(S_DATE, 6, 2)
        S_MM = Format(S_MM + 1, "00")

        S_DATE = Right(S_DATE, 5) & S_MM & "/01"


        If IsDate(S_DATE) Then
        Else
            S_YY = Right(S_DATE, 4)
            S_YY = Format(CInt(S_YY) + 1, "0000")

            S_DATE = S_YY & "/01/01"
        End If
    End If


    Text1(ptxS_JITU_DATE).Text = S_DATE
    Text1(ptxE_JITU_DATE).Text = E_DATE



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
Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   明細表示
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
    
Dim j           As Integer
    
Dim GK_MAISU    As Long
Dim GK_KINGAKU  As Long
Dim GK_SAISU    As Long
    
Dim Skip_Flg    As Boolean
    
Dim End_Date    As String
    
    
    
Dim Fast_Flg    As Boolean
    
    
    List_Disp_Proc = True
                                    
    Call Input_Lock
    
    Set SE_USOU_HAKO = Nothing
    
    
    
    
    
    
    'ｸﾞﾘｯﾄﾞﾍｯﾀﾞｰｾｯﾄ
        
    
    i = ColDay - 1
    
    End_Date = Text1(ptxS_JITU_DATE).Text
    Do
        
        i = i + 1
        TDBGrid1.Columns(i).Caption = Mid(End_Date, 6, 10)
        End_Date = DateAdd("d", 1, End_Date)
        If End_Date > Text1(ptxE_JITU_DATE).Text Then
            Exit Do
        End If
    
    Loop
    
    
    
    
    
    
    If IsDate(Text1(ptxS_JITU_DATE).Text) Then
        Call UniCode_Conv(K0_SE_USOU_HAKO.JITU_DATE, Format((Text1(ptxS_JITU_DATE).Text), "YYYYMMDD"))
    Else
        Call UniCode_Conv(K0_SE_USOU_HAKO.JITU_DATE, Text1(ptxS_JITU_DATE).Text)
    End If
    
    If IsDate(Text1(ptxE_JITU_DATE).Text) Then
        End_Date = Format((Text1(ptxE_JITU_DATE).Text), "YYYYMMDD")
    Else
        End_Date = Text1(ptxE_JITU_DATE).Text
    End If
    
    
    
    
    
    
    
    Fast_Flg = True
    
    
    Call UniCode_Conv(K0_SE_USOU_HAKO.JGYOBU, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.NAIGAI, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.HIN_GAI, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.MTS_CODE, "")
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        
        sts = BTRV(com, SE_USOU_HAKO_POS, SE_USOU_HAKOREC, Len(SE_USOU_HAKOREC), K0_SE_USOU_HAKO, Len(K0_SE_USOU_HAKO), 0)
    
        Select Case sts
            Case BtNoErr
        
                If StrConv(SE_USOU_HAKOREC.JITU_DATE, vbUnicode) > End_Date Then
                    Exit Do
                End If
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "輸送箱実績")
                Exit Function
        End Select
                
        Skip_Flg = False
        If CLng(StrConv(SE_USOU_HAKOREC.MAISU, vbUnicode)) = 0 Then
            Skip_Flg = True
        End If
                        
                
        If Not Skip_Flg Then
            
            
            
            If Fast_Flg Then
                SE_USOU_HAKO.ReDim Min_Row, 1, Min_Col, Max_Col
                SE_USOU_HAKO(1, ColHIN_GAI) = Trim(StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode))
                
                Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(SE_USOU_HAKOREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(SE_USOU_HAKOREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode))
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        SE_USOU_HAKO(1, ColHIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                    
                    
                        '才数
                        If IsNumeric(StrConv(ITEMREC.SAI_SU, vbUnicode)) Then
                            SE_USOU_HAKO(1, ColSAI_SU) = Format(CDbl(StrConv(ITEMREC.SAI_SU, vbUnicode)), "#0.0")
                        Else
                            SE_USOU_HAKO(1, ColSAI_SU) = "0.0"
                        End If
                    
                        '単価
                        If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                            SE_USOU_HAKO(1, ColTANKA) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.0")
                        Else
                            SE_USOU_HAKO(1, ColTANKA) = "0.0"
                        End If
                    
                    
                    Case BtErrKeyNotFound
                        SE_USOU_HAKO(1, ColHIN_NAME) = ""
                        SE_USOU_HAKO(1, ColSAI_SU) = "0.0"
                        SE_USOU_HAKO(1, ColTANKA) = "0.0"
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                End Select
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                Fast_Flg = False
            End If
            
            
            If Grid_Set_Proc() Then
                Exit Function
            End If
        End If
        
        com = BtOpGetNext
        
    Loop
    
                                
                                'DBテーブルリンク
    
    For i = 1 To SE_USOU_HAKO.UpperBound(1)
    
    
        j = ColDay - 1
    
        End_Date = Text1(ptxS_JITU_DATE).Text
        Do
            j = j + 1
            SE_USOU_HAKO(i, j) = Format(CLng(SE_USOU_HAKO(i, j)), "#,##0") & "/" & Format(CDbl(SE_USOU_HAKO(i, ColSAI_SU)) * CLng(SE_USOU_HAKO(i, j)), "#,##0")
            End_Date = DateAdd("d", 1, End_Date)
            If End_Date > Text1(ptxE_JITU_DATE).Text Then
                Exit Do
            End If
        Loop
        
    
    Next i
    
    
    SE_USOU_HAKO.QuickSort Min_Row, SE_USOU_HAKO.UpperBound(1), ColHIN_GAI, 0, XTYPE_STRING
    
    
    Set TDBGrid1.Array = SE_USOU_HAKO
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.ScrollBars = dbgAutomatic
    
    Call Input_UnLock
    
    
    Text1(ptxS_JITU_DATE).SetFocus
    
    List_Disp_Proc = False

    
End Function

Private Function DETAIL_Proc() As Integer
'----------------------------------------------------------------------------
'                   ＥＸＣＥＬ（明細）出力
'----------------------------------------------------------------------------

    '2012.09.26 excel-->objcet
'Dim excelApplication    As excel.Application
'Dim excelWorkBook       As excel.Workbook
'Dim excelSheet          As excel.Worksheet
Dim excelApplication    As Object
Dim excelWorkBook       As Object
Dim excelSheet          As Object
    '2012.09.26 excel-->objcet



Dim i                   As Integer
Dim j                   As Integer
Dim Row                 As Integer
    
    
Dim sts                 As Integer
Dim com                 As Integer
    
Dim End_Date            As String
Dim svMM                As String

Dim Fast_Flg            As Boolean

Dim ParaM               As String
 
Dim List_index          As Integer
    
    
    DETAIL_Proc = True
    
    Call Input_Lock
    
    Set SE_USOU_HAKO = Nothing
    
    
    Set excelApplication = CreateObject("Excel.Application")
'2009.02.26    excelApplication.Visible = True


    
    Set excelWorkBook = excelApplication.Workbooks.Add
    Set excelSheet = excelWorkBook.Worksheets(1)
    
    excelApplication.StandardFontSize = 11
    excelApplication.StandardFont = "ＭＳ　ゴシック"

    
    

    
    
    i = 3
    svMM = ""
    
    End_Date = Text1(ptxS_JITU_DATE).Text
    Do
        
        i = i + 1
        
        
        If svMM = "" Then
            excelSheet.Application.Cells(3, i).Value = Mid(End_Date, 6, 2)
            svMM = Mid(End_Date, 6, 2)
        End If
        
        If svMM <> Mid(End_Date, 6, 2) Then
            excelSheet.Application.Cells(3, i).Value = Mid(End_Date, 6, 2)
            svMM = Mid(End_Date, 6, 2)
        End If
        
        
        excelSheet.Application.Cells(4, i).Value = Right(End_Date, 2)
        End_Date = DateAdd("d", 1, End_Date)
        If End_Date > Text1(ptxE_JITU_DATE).Text Then
            Exit Do
        End If
    
    Loop
    i = i + 1
    excelSheet.Application.Cells(4, i).Value = "合計"
    
    
    '罫線
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(4, 3)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    excelSheet.Application.Range(excelSheet.Application.Cells(3, 4), excelSheet.Application.Cells(4, i)).Select
    With excelSheet.Application.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    
    
    
    
    
    If IsDate(Text1(ptxS_JITU_DATE).Text) Then
        Call UniCode_Conv(K0_SE_USOU_HAKO.JITU_DATE, Format((Text1(ptxS_JITU_DATE).Text), "YYYYMMDD"))
    Else
        Call UniCode_Conv(K0_SE_USOU_HAKO.JITU_DATE, Text1(ptxS_JITU_DATE).Text)
    End If
    
    If IsDate(Text1(ptxE_JITU_DATE).Text) Then
        End_Date = Format((Text1(ptxE_JITU_DATE).Text), "YYYYMMDD")
    Else
        End_Date = Text1(ptxE_JITU_DATE).Text
    End If
    
    
    
    
    
    
    
    Fast_Flg = True
    
    
        
    
    
    
    Call UniCode_Conv(K0_SE_USOU_HAKO.JGYOBU, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.NAIGAI, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.HIN_GAI, "")
    Call UniCode_Conv(K0_SE_USOU_HAKO.MTS_CODE, "")
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        
        sts = BTRV(com, SE_USOU_HAKO_POS, SE_USOU_HAKOREC, Len(SE_USOU_HAKOREC), K0_SE_USOU_HAKO, Len(K0_SE_USOU_HAKO), 0)
    
        Select Case sts
            Case BtNoErr
        
                If StrConv(SE_USOU_HAKOREC.JITU_DATE, vbUnicode) > End_Date Then
                    Exit Do
                End If
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "輸送箱実績")
                Exit Function
        End Select
                
            
            
            
        If Fast_Flg Then
            SE_USOU_HAKO.ReDim Min_Row, 1, Min_Col, Max_Col
            SE_USOU_HAKO(1, ColHIN_GAI) = Trim(StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode))
            
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(SE_USOU_HAKOREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(SE_USOU_HAKOREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    SE_USOU_HAKO(1, ColHIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                
                
                    '才数
                    If IsNumeric(StrConv(ITEMREC.SAI_SU, vbUnicode)) Then
                        SE_USOU_HAKO(1, ColSAI_SU) = Format(CDbl(StrConv(ITEMREC.SAI_SU, vbUnicode)), "#0.0")
                    Else
                        SE_USOU_HAKO(1, ColSAI_SU) = "0.0"
                    End If
                
                    '単価
                    If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                        SE_USOU_HAKO(1, ColTANKA) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.0")
                    Else
                        SE_USOU_HAKO(1, ColTANKA) = "0.0"
                    End If
                
                
                Case BtErrKeyNotFound
                    SE_USOU_HAKO(1, ColHIN_NAME) = ""
                    SE_USOU_HAKO(1, ColSAI_SU) = "0.0"
                    SE_USOU_HAKO(1, ColTANKA) = "0.0"
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
            
            Fast_Flg = False
        End If
            
            
        If Grid_Set_Proc() Then
            Exit Function
        End If
        
        com = BtOpGetNext
        
    Loop
    
    SE_USOU_HAKO.QuickSort Min_Row, SE_USOU_HAKO.UpperBound(1), ColHIN_GAI, 0, XTYPE_STRING
    
    
    '---------- EXCEL 出力
    Row = 3
    For i = 1 To SE_USOU_HAKO.UpperBound(1)
    
        Row = Row + 2
        If Excel_Set_Proc(excelApplication, excelWorkBook, excelSheet, Row, i, SHIZAI, NAIGAI_NAI, SE_USOU_HAKO(i, ColHIN_GAI)) Then
            Unload Me
        End If
    
    Next i
    
    
    '---------- EXCEL 出力（合計）
    Row = Row + 2
    'セルの結合
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, 1)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 2)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row + 1, 2)).MergeCells = True
    
    excelSheet.Application.Cells(Row, 1).Value = "合計"
    '『才数』
    excelSheet.Application.Cells(Row, 3).Value = "才数"
    '『個数』
    excelSheet.Application.Cells(Row + 1, 3).Value = "個数"
    '罫線
    
    i = 3
    
    End_Date = Text1(ptxS_JITU_DATE).Text
    Do
        
        i = i + 1
        
        ParaM = ""

        j = ((Row - 1) - 4) * -1
        Do
            
            If ParaM = "" Then
                ParaM = "=R[" & j & "]C"
            Else
                ParaM = ParaM & "+R[" & j & "]C"
            End If
        
                    
            j = j + 2
            If j > -1 Then
                Exit Do
            End If
        
        
        Loop
        
        excelSheet.Application.Range(excelSheet.Application.Cells(Row, i), excelSheet.Application.Cells(Row, i)).Select
        excelSheet.Application.ActiveCell.FormulaR1C1 = ParaM
        excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, i), excelSheet.Application.Cells(Row + 1, i)).Select
        excelSheet.Application.ActiveCell.FormulaR1C1 = ParaM
        
        
        
        End_Date = DateAdd("d", 1, End_Date)
        If End_Date > Text1(ptxE_JITU_DATE).Text Then
            Exit Do
        End If
    
    Loop
    i = i + 1
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, i), excelSheet.Application.Cells(Row, i)).Select
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(RC[" & (-i + 4) & "]:RC[-1]"
    excelSheet.Application.Range(excelSheet.Application.Cells(Row + 1, i), excelSheet.Application.Cells(Row + 1, i)).Select
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(RC[" & (-i + 4) & "]:RC[-1]"





    '罫線
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row + 1, i)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With


    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 3), excelSheet.Application.Cells(Row + 1, i)).Select
    With excelSheet.Application.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    
    
    For j = 1 To i - 1
        
        excelSheet.Application.Columns(j).Select
        excelSheet.Application.Selection.ColumnWidth = 4
        
        
    Next j
    
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 9), excelSheet.Application.Cells(1, 16)).Select
    With excelSheet.Application.Selection.Font
        .Size = 14
    End With
    
    excelSheet.Application.Cells(1, 8).Value = "輸送箱出荷個数＆才数 " & _
                                                "（" & StrConv(Text1(ptxS_JITU_DATE).Text, vbWide) & "～" & _
                                                StrConv(Text1(ptxE_JITU_DATE).Text, vbWide) & "）"
    
    
    
    excelApplication.Visible = True '2009.02.26
    
    Set excelSheet = Nothing
    Set excelWorkBook = Nothing
    Set excelApplication = Nothing


    
    Call Input_UnLock
    DETAIL_Proc = False
    

End Function
Private Function Excel_Set_Proc(excelApplication As Object, excelWorkBook As Object, excelSheet As Object, _
                                                excelRow As Integer, gridRow As Integer, svJGYOBU As String, svNAIGAI As String, svHIN_GAI As String) As Integer
'Private Function Excel_Set_Proc(excelApplication As excel.Application, excelWorkBook As excel.Workbook, excelSheet As excel.Worksheet, _
'                                                excelRow As Integer, gridRow As Integer, svJGYOBU As String, svNAIGAI As String, svHIN_GAI As String) As Integer
'----------------------------------------------------------------------------
'           輸送箱--＞EXCEL
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim i           As Integer
     
Dim End_Date    As String
    
    
    Excel_Set_Proc = True
        
    'セルの結合
    excelSheet.Application.Range(excelSheet.Application.Cells(excelRow, 1), excelSheet.Application.Cells(excelRow + 1, 1)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(excelRow, 1), excelSheet.Application.Cells(excelRow + 1, 1)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(excelRow, 2), excelSheet.Application.Cells(excelRow + 1, 2)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(excelRow, 2), excelSheet.Application.Cells(excelRow + 1, 2)).MergeCells = True
    
    '品番
    Call UniCode_Conv(K0_ITEM.JGYOBU, svJGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, svNAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, svHIN_GAI)
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
        
            Call UniCode_Conv(ITEMREC.SAI_SU, "0.0")
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    End Select
    excelSheet.Application.Cells(excelRow, 1).Value = svHIN_GAI
    
    If IsNumeric(StrConv(ITEMREC.SAI_SU, vbUnicode)) Then
        excelSheet.Application.Cells(excelRow, 2).Value = Format(CDbl(StrConv(ITEMREC.SAI_SU, vbUnicode)))
    Else
        excelSheet.Application.Cells(excelRow, 2).Value = 0
    End If
    '『才数』
    excelSheet.Application.Cells(excelRow, 3).Value = "才数"
    '『個数』
    excelSheet.Application.Cells(excelRow + 1, 3).Value = "個数"
    '罫線
    
    i = 3
    
    End_Date = Text1(ptxS_JITU_DATE).Text
    Do
        
        i = i + 1
        
        excelSheet.Application.Cells(excelRow, i).NumberFormatLocal = "#,###_ "
        excelSheet.Application.Cells(excelRow, i).Value = CDbl(SE_USOU_HAKO(gridRow, ColSAI_SU)) * CLng(SE_USOU_HAKO(gridRow, i))
        excelSheet.Application.Cells(excelRow + 1, i).NumberFormatLocal = "#,###_ "
        excelSheet.Application.Cells(excelRow + 1, i).Value = CLng(SE_USOU_HAKO(gridRow, i))
        
        
        
        End_Date = DateAdd("d", 1, End_Date)
        If End_Date > Text1(ptxE_JITU_DATE).Text Then
            Exit Do
        End If
    
    Loop
    i = i + 1
    
    excelSheet.Application.Range(excelSheet.Application.Cells(excelRow, i), excelSheet.Application.Cells(excelRow, i)).Select
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(RC[" & (-i + 4) & "]:RC[-1]"
    excelSheet.Application.Range(excelSheet.Application.Cells(excelRow + 1, i), excelSheet.Application.Cells(excelRow + 1, i)).Select
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(RC[" & (-i + 4) & "]:RC[-1]"





    '罫線
    excelSheet.Application.Range(excelSheet.Application.Cells(excelRow, 1), excelSheet.Application.Cells(excelRow + 1, i)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With


    excelSheet.Application.Range(excelSheet.Application.Cells(excelRow, 3), excelSheet.Application.Cells(excelRow + 1, i)).Select
    With excelSheet.Application.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With



    Excel_Set_Proc = False

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    SEI00501.MousePointer = vbHourglass


    Call Ctrl_Lock(SEI00501)

    TDBGrid1.Enabled = False


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(SEI00501)
    
    TDBGrid1.Enabled = True


    SEI00501.MousePointer = vbDefault

End Sub

Private Function Grid_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   輸送箱実績-->Ｇｒｉｄ
'----------------------------------------------------------------------------

Dim sts     As Integer
Dim wkDec   As Long
    
Dim i       As Long
Dim j       As Long
    
    Grid_Set_Proc = True

    
    
    
    For i = 1 To SE_USOU_HAKO.UpperBound(1)
    
    
        If Trim(StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode)) = Trim(SE_USOU_HAKO(i, ColHIN_GAI)) Then
            Exit For
        End If
    
    
    Next i
    
    
    
    If i > SE_USOU_HAKO.UpperBound(1) Then
        SE_USOU_HAKO.ReDim Min_Row, i, Min_Col, Max_Col
        SE_USOU_HAKO(i, ColHIN_GAI) = Trim(StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode))
    
    
    
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(SE_USOU_HAKOREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(SE_USOU_HAKOREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(SE_USOU_HAKOREC.HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                SE_USOU_HAKO(i, ColHIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
            
            
                '才数
                If IsNumeric(StrConv(ITEMREC.SAI_SU, vbUnicode)) Then
                    SE_USOU_HAKO(i, ColSAI_SU) = Format(CDbl(StrConv(ITEMREC.SAI_SU, vbUnicode)), "#0.0")
                Else
                    SE_USOU_HAKO(i, ColSAI_SU) = "0.0"
                End If
            
                '単価
                If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                    SE_USOU_HAKO(i, ColTANKA) = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.0")
                Else
                    SE_USOU_HAKO(i, ColTANKA) = "0.0"
                End If
            
            
            Case BtErrKeyNotFound
                SE_USOU_HAKO(i, ColHIN_NAME) = ""
                SE_USOU_HAKO(i, ColSAI_SU) = "0.0"
                SE_USOU_HAKO(i, ColTANKA) = "0.0"
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        End Select
    
    End If
    
    
    
    j = (DateDiff("D", Text1(ptxS_JITU_DATE).Text, Mid(StrConv(SE_USOU_HAKOREC.JITU_DATE, vbUnicode), 1, 4) & "/" & Mid(StrConv(SE_USOU_HAKOREC.JITU_DATE, vbUnicode), 5, 2) & "/" & Mid(StrConv(SE_USOU_HAKOREC.JITU_DATE, vbUnicode), 7, 2)) + 1) + ColTANKA
    
    
    
    
    
    
    
    
    
    
    
    SE_USOU_HAKO(i, j) = CLng(SE_USOU_HAKO(i, j)) + CLng(StrConv(SE_USOU_HAKOREC.MAISU, vbUnicode))
    
    
    
    Grid_Set_Proc = False
End Function

Private Sub SHORI_Click(Index As Integer)
    Select Case Index
    
        
        
        
        Case 0 To 2
        
        
            Command1(Index).Value = True
        
        
        
        Case 3      '画面印刷
        
            Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)
                    
    
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
    
    
        Case ptxS_JITU_DATE     '開始日付
            
            
            If Not IsDate(Text1(Mode).Text) Then
                MsgBox "入力した項目はエラーです。（日付エラー）"
                Text1(Mode).SetFocus
                Exit Function
            End If
            
        Case ptxE_JITU_DATE     '終了日付
            
            If Not IsDate(Text1(Mode).Text) Then
                MsgBox "入力した項目はエラーです。（日付エラー）"
                Text1(Mode).SetFocus
                Exit Function
            End If
            
    
            If DateDiff("d", Text1(ptxS_JITU_DATE).Text, Text1(ptxE_JITU_DATE).Text) < 1 Or _
                DateDiff("d", Text1(ptxS_JITU_DATE).Text, Text1(ptxE_JITU_DATE).Text) > 31 Then
                MsgBox "入力した項目はエラーです。（日付範囲エラー）"
                Text1(Mode).SetFocus
                Exit Function
            End If
    
    
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function

