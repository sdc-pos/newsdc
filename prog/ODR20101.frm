VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form ODR20101 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "親部品　注文情報照会 [ODR2010] 2012.04.14 08:15"
   ClientHeight    =   10110
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14835
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
   ScaleHeight     =   10110
   ScaleWidth      =   14835
   StartUpPosition =   2  '画面の中央
   Begin VB.ComboBox Combo1 
      Height          =   345
      Index           =   0
      Left            =   8760
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "検　索"
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
      Left            =   360
      TabIndex        =   4
      Top             =   60
      Width           =   1800
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
      Index           =   0
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   1
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
      Index           =   1
      Left            =   3240
      MaxLength       =   5
      TabIndex        =   2
      Top             =   780
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   6930
      ScaleHeight     =   195
      ScaleWidth      =   270
      TabIndex        =   6
      Top             =   120
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
      Index           =   2
      Left            =   2520
      TabIndex        =   5
      Top             =   60
      Width           =   1800
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   8475
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1320
      Width           =   14580
      _ExtentX        =   25718
      _ExtentY        =   14949
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "　親部品　注文№"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "親品番"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "受注数"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "合　計"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "01"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "02"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "03"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "04"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "05"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "06"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "07"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "08"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "09"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "10"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "11"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "12"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "13"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "14"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "15"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "16"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "17"
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "18"
      Columns(21).DataField=   ""
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "19"
      Columns(22).DataField=   ""
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "20"
      Columns(23).DataField=   ""
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "21"
      Columns(24).DataField=   ""
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).Caption=   "22"
      Columns(25).DataField=   ""
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(26)._VlistStyle=   0
      Columns(26)._MaxComboItems=   5
      Columns(26).Caption=   "23"
      Columns(26).DataField=   ""
      Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(27)._VlistStyle=   0
      Columns(27)._MaxComboItems=   5
      Columns(27).Caption=   "24"
      Columns(27).DataField=   ""
      Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(28)._VlistStyle=   0
      Columns(28)._MaxComboItems=   5
      Columns(28).Caption=   "25"
      Columns(28).DataField=   ""
      Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(29)._VlistStyle=   0
      Columns(29)._MaxComboItems=   5
      Columns(29).Caption=   "26"
      Columns(29).DataField=   ""
      Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(30)._VlistStyle=   0
      Columns(30)._MaxComboItems=   5
      Columns(30).Caption=   "27"
      Columns(30).DataField=   ""
      Columns(30)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(31)._VlistStyle=   0
      Columns(31)._MaxComboItems=   5
      Columns(31).Caption=   "28"
      Columns(31).DataField=   ""
      Columns(31)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(32)._VlistStyle=   0
      Columns(32)._MaxComboItems=   5
      Columns(32).Caption=   "29"
      Columns(32).DataField=   ""
      Columns(32)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(33)._VlistStyle=   0
      Columns(33)._MaxComboItems=   5
      Columns(33).Caption=   "30"
      Columns(33).DataField=   ""
      Columns(33)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(34)._VlistStyle=   0
      Columns(34)._MaxComboItems=   5
      Columns(34).Caption=   "31"
      Columns(34).DataField=   ""
      Columns(34)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(35)._VlistStyle=   0
      Columns(35)._MaxComboItems=   5
      Columns(35).Caption=   "ＫＥＹ項目"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2461"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2328"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2646"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2514"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=1402"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1270"
      Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=1402"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=1270"
      Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=1270"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=1138"
      Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=8194"
      Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(25)=   "Column(5).Width=1270"
      Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=1138"
      Splits(0)._ColumnProps(28)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(30)=   "Column(6).Width=1270"
      Splits(0)._ColumnProps(31)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(6)._WidthInPix=1138"
      Splits(0)._ColumnProps(33)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(34)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(35)=   "Column(7).Width=1270"
      Splits(0)._ColumnProps(36)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(7)._WidthInPix=1138"
      Splits(0)._ColumnProps(38)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(39)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(40)=   "Column(8).Width=1270"
      Splits(0)._ColumnProps(41)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(8)._WidthInPix=1138"
      Splits(0)._ColumnProps(43)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(44)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(45)=   "Column(9).Width=1270"
      Splits(0)._ColumnProps(46)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(9)._WidthInPix=1138"
      Splits(0)._ColumnProps(48)=   "Column(9)._ColStyle=2"
      Splits(0)._ColumnProps(49)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(50)=   "Column(10).Width=1270"
      Splits(0)._ColumnProps(51)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(10)._WidthInPix=1138"
      Splits(0)._ColumnProps(53)=   "Column(10)._ColStyle=2"
      Splits(0)._ColumnProps(54)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(55)=   "Column(11).Width=1270"
      Splits(0)._ColumnProps(56)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(11)._WidthInPix=1138"
      Splits(0)._ColumnProps(58)=   "Column(11)._ColStyle=2"
      Splits(0)._ColumnProps(59)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(60)=   "Column(12).Width=1270"
      Splits(0)._ColumnProps(61)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(12)._WidthInPix=1138"
      Splits(0)._ColumnProps(63)=   "Column(12)._ColStyle=2"
      Splits(0)._ColumnProps(64)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(65)=   "Column(13).Width=1270"
      Splits(0)._ColumnProps(66)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(13)._WidthInPix=1138"
      Splits(0)._ColumnProps(68)=   "Column(13)._ColStyle=2"
      Splits(0)._ColumnProps(69)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(70)=   "Column(14).Width=1270"
      Splits(0)._ColumnProps(71)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(72)=   "Column(14)._WidthInPix=1138"
      Splits(0)._ColumnProps(73)=   "Column(14)._ColStyle=2"
      Splits(0)._ColumnProps(74)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(75)=   "Column(15).Width=1270"
      Splits(0)._ColumnProps(76)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(77)=   "Column(15)._WidthInPix=1138"
      Splits(0)._ColumnProps(78)=   "Column(15)._ColStyle=2"
      Splits(0)._ColumnProps(79)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(80)=   "Column(16).Width=1270"
      Splits(0)._ColumnProps(81)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(82)=   "Column(16)._WidthInPix=1138"
      Splits(0)._ColumnProps(83)=   "Column(16)._ColStyle=2"
      Splits(0)._ColumnProps(84)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(85)=   "Column(17).Width=1270"
      Splits(0)._ColumnProps(86)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(87)=   "Column(17)._WidthInPix=1138"
      Splits(0)._ColumnProps(88)=   "Column(17)._ColStyle=2"
      Splits(0)._ColumnProps(89)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(90)=   "Column(18).Width=1270"
      Splits(0)._ColumnProps(91)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(92)=   "Column(18)._WidthInPix=1138"
      Splits(0)._ColumnProps(93)=   "Column(18)._ColStyle=2"
      Splits(0)._ColumnProps(94)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(95)=   "Column(19).Width=1270"
      Splits(0)._ColumnProps(96)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(97)=   "Column(19)._WidthInPix=1138"
      Splits(0)._ColumnProps(98)=   "Column(19)._ColStyle=2"
      Splits(0)._ColumnProps(99)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(100)=   "Column(20).Width=1270"
      Splits(0)._ColumnProps(101)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(102)=   "Column(20)._WidthInPix=1138"
      Splits(0)._ColumnProps(103)=   "Column(20)._ColStyle=2"
      Splits(0)._ColumnProps(104)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(105)=   "Column(21).Width=1270"
      Splits(0)._ColumnProps(106)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(107)=   "Column(21)._WidthInPix=1138"
      Splits(0)._ColumnProps(108)=   "Column(21)._ColStyle=2"
      Splits(0)._ColumnProps(109)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(110)=   "Column(22).Width=1270"
      Splits(0)._ColumnProps(111)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(112)=   "Column(22)._WidthInPix=1138"
      Splits(0)._ColumnProps(113)=   "Column(22)._ColStyle=2"
      Splits(0)._ColumnProps(114)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(115)=   "Column(23).Width=1270"
      Splits(0)._ColumnProps(116)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(117)=   "Column(23)._WidthInPix=1138"
      Splits(0)._ColumnProps(118)=   "Column(23)._ColStyle=2"
      Splits(0)._ColumnProps(119)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(120)=   "Column(24).Width=1270"
      Splits(0)._ColumnProps(121)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(122)=   "Column(24)._WidthInPix=1138"
      Splits(0)._ColumnProps(123)=   "Column(24)._ColStyle=2"
      Splits(0)._ColumnProps(124)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(125)=   "Column(25).Width=1270"
      Splits(0)._ColumnProps(126)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(127)=   "Column(25)._WidthInPix=1138"
      Splits(0)._ColumnProps(128)=   "Column(25)._ColStyle=2"
      Splits(0)._ColumnProps(129)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(130)=   "Column(26).Width=1270"
      Splits(0)._ColumnProps(131)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(132)=   "Column(26)._WidthInPix=1138"
      Splits(0)._ColumnProps(133)=   "Column(26)._ColStyle=2"
      Splits(0)._ColumnProps(134)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(135)=   "Column(27).Width=1270"
      Splits(0)._ColumnProps(136)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(137)=   "Column(27)._WidthInPix=1138"
      Splits(0)._ColumnProps(138)=   "Column(27)._ColStyle=2"
      Splits(0)._ColumnProps(139)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(140)=   "Column(28).Width=1270"
      Splits(0)._ColumnProps(141)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(142)=   "Column(28)._WidthInPix=1138"
      Splits(0)._ColumnProps(143)=   "Column(28)._ColStyle=2"
      Splits(0)._ColumnProps(144)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(145)=   "Column(29).Width=1270"
      Splits(0)._ColumnProps(146)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(147)=   "Column(29)._WidthInPix=1138"
      Splits(0)._ColumnProps(148)=   "Column(29)._ColStyle=2"
      Splits(0)._ColumnProps(149)=   "Column(29).Order=30"
      Splits(0)._ColumnProps(150)=   "Column(30).Width=1270"
      Splits(0)._ColumnProps(151)=   "Column(30).DividerColor=0"
      Splits(0)._ColumnProps(152)=   "Column(30)._WidthInPix=1138"
      Splits(0)._ColumnProps(153)=   "Column(30)._ColStyle=2"
      Splits(0)._ColumnProps(154)=   "Column(30).Order=31"
      Splits(0)._ColumnProps(155)=   "Column(31).Width=1270"
      Splits(0)._ColumnProps(156)=   "Column(31).DividerColor=0"
      Splits(0)._ColumnProps(157)=   "Column(31)._WidthInPix=1138"
      Splits(0)._ColumnProps(158)=   "Column(31)._ColStyle=2"
      Splits(0)._ColumnProps(159)=   "Column(31).Order=32"
      Splits(0)._ColumnProps(160)=   "Column(32).Width=1270"
      Splits(0)._ColumnProps(161)=   "Column(32).DividerColor=0"
      Splits(0)._ColumnProps(162)=   "Column(32)._WidthInPix=1138"
      Splits(0)._ColumnProps(163)=   "Column(32)._ColStyle=2"
      Splits(0)._ColumnProps(164)=   "Column(32).Order=33"
      Splits(0)._ColumnProps(165)=   "Column(33).Width=1270"
      Splits(0)._ColumnProps(166)=   "Column(33).DividerColor=0"
      Splits(0)._ColumnProps(167)=   "Column(33)._WidthInPix=1138"
      Splits(0)._ColumnProps(168)=   "Column(33)._ColStyle=2"
      Splits(0)._ColumnProps(169)=   "Column(33).Order=34"
      Splits(0)._ColumnProps(170)=   "Column(34).Width=1270"
      Splits(0)._ColumnProps(171)=   "Column(34).DividerColor=0"
      Splits(0)._ColumnProps(172)=   "Column(34)._WidthInPix=1138"
      Splits(0)._ColumnProps(173)=   "Column(34)._ColStyle=2"
      Splits(0)._ColumnProps(174)=   "Column(34).Order=35"
      Splits(0)._ColumnProps(175)=   "Column(35).Width=2778"
      Splits(0)._ColumnProps(176)=   "Column(35).DividerColor=0"
      Splits(0)._ColumnProps(177)=   "Column(35)._WidthInPix=2646"
      Splits(0)._ColumnProps(178)=   "Column(35).Visible=0"
      Splits(0)._ColumnProps(179)=   "Column(35).Order=36"
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
      HeadLines       =   2
      FootLines       =   1
      Caption         =   "親部品　注文情報"
      MultipleLines   =   0
      CellTipsWidth   =   0
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.bold=0,.fontsize=1125"
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
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&H80FF80&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFF80&"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=87,.parent=1,.bgcolor=&H80FF80&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=96,.parent=4,.bgcolor=&H80FF00&"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=89,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=91,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=94,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=95,.parent=10,.namedParent=40,.bgcolor=&H80FF00&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=102,.parent=87,.alignment=2"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=99,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=100,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=101,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=110,.parent=87"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=107,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=108,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=109,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=114,.parent=87,.alignment=1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=111,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=112,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=113,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=118,.parent=87,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=115,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=116,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=117,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=122,.parent=87,.namedParent=13,.alignment=1"
      _StyleDefs(53)  =   ":id=122,.locked=-1"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=119,.parent=88"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=120,.parent=89"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=121,.parent=91"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=126,.parent=87,.alignment=1"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=123,.parent=88"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=124,.parent=89"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=125,.parent=91"
      _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=21,.parent=87,.alignment=1"
      _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=18,.parent=88"
      _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=19,.parent=89"
      _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=20,.parent=91"
      _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=25,.parent=87,.alignment=1"
      _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=22,.parent=88"
      _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=23,.parent=89"
      _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=24,.parent=91"
      _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=17,.parent=87,.alignment=1"
      _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=14,.parent=88"
      _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=15,.parent=89"
      _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=16,.parent=91"
      _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=130,.parent=87,.alignment=1"
      _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=127,.parent=88"
      _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=128,.parent=89"
      _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=129,.parent=91"
      _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=29,.parent=87,.alignment=1"
      _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=26,.parent=88"
      _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=27,.parent=89"
      _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=28,.parent=91"
      _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=43,.parent=87,.alignment=1"
      _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=30,.parent=88"
      _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=31,.parent=89"
      _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=32,.parent=91"
      _StyleDefs(85)  =   "Splits(0).Columns(12).Style:id=47,.parent=87,.alignment=1"
      _StyleDefs(86)  =   "Splits(0).Columns(12).HeadingStyle:id=44,.parent=88"
      _StyleDefs(87)  =   "Splits(0).Columns(12).FooterStyle:id=45,.parent=89"
      _StyleDefs(88)  =   "Splits(0).Columns(12).EditorStyle:id=46,.parent=91"
      _StyleDefs(89)  =   "Splits(0).Columns(13).Style:id=51,.parent=87,.alignment=1"
      _StyleDefs(90)  =   "Splits(0).Columns(13).HeadingStyle:id=48,.parent=88"
      _StyleDefs(91)  =   "Splits(0).Columns(13).FooterStyle:id=49,.parent=89"
      _StyleDefs(92)  =   "Splits(0).Columns(13).EditorStyle:id=50,.parent=91"
      _StyleDefs(93)  =   "Splits(0).Columns(14).Style:id=55,.parent=87,.alignment=1"
      _StyleDefs(94)  =   "Splits(0).Columns(14).HeadingStyle:id=52,.parent=88"
      _StyleDefs(95)  =   "Splits(0).Columns(14).FooterStyle:id=53,.parent=89"
      _StyleDefs(96)  =   "Splits(0).Columns(14).EditorStyle:id=54,.parent=91"
      _StyleDefs(97)  =   "Splits(0).Columns(15).Style:id=59,.parent=87,.alignment=1"
      _StyleDefs(98)  =   "Splits(0).Columns(15).HeadingStyle:id=56,.parent=88"
      _StyleDefs(99)  =   "Splits(0).Columns(15).FooterStyle:id=57,.parent=89"
      _StyleDefs(100) =   "Splits(0).Columns(15).EditorStyle:id=58,.parent=91"
      _StyleDefs(101) =   "Splits(0).Columns(16).Style:id=63,.parent=87,.alignment=1"
      _StyleDefs(102) =   "Splits(0).Columns(16).HeadingStyle:id=60,.parent=88"
      _StyleDefs(103) =   "Splits(0).Columns(16).FooterStyle:id=61,.parent=89"
      _StyleDefs(104) =   "Splits(0).Columns(16).EditorStyle:id=62,.parent=91"
      _StyleDefs(105) =   "Splits(0).Columns(17).Style:id=67,.parent=87,.alignment=1"
      _StyleDefs(106) =   "Splits(0).Columns(17).HeadingStyle:id=64,.parent=88"
      _StyleDefs(107) =   "Splits(0).Columns(17).FooterStyle:id=65,.parent=89"
      _StyleDefs(108) =   "Splits(0).Columns(17).EditorStyle:id=66,.parent=91"
      _StyleDefs(109) =   "Splits(0).Columns(18).Style:id=71,.parent=87,.alignment=1"
      _StyleDefs(110) =   "Splits(0).Columns(18).HeadingStyle:id=68,.parent=88"
      _StyleDefs(111) =   "Splits(0).Columns(18).FooterStyle:id=69,.parent=89"
      _StyleDefs(112) =   "Splits(0).Columns(18).EditorStyle:id=70,.parent=91"
      _StyleDefs(113) =   "Splits(0).Columns(19).Style:id=75,.parent=87,.alignment=1"
      _StyleDefs(114) =   "Splits(0).Columns(19).HeadingStyle:id=72,.parent=88"
      _StyleDefs(115) =   "Splits(0).Columns(19).FooterStyle:id=73,.parent=89"
      _StyleDefs(116) =   "Splits(0).Columns(19).EditorStyle:id=74,.parent=91"
      _StyleDefs(117) =   "Splits(0).Columns(20).Style:id=79,.parent=87,.alignment=1"
      _StyleDefs(118) =   "Splits(0).Columns(20).HeadingStyle:id=76,.parent=88"
      _StyleDefs(119) =   "Splits(0).Columns(20).FooterStyle:id=77,.parent=89"
      _StyleDefs(120) =   "Splits(0).Columns(20).EditorStyle:id=78,.parent=91"
      _StyleDefs(121) =   "Splits(0).Columns(21).Style:id=83,.parent=87,.alignment=1"
      _StyleDefs(122) =   "Splits(0).Columns(21).HeadingStyle:id=80,.parent=88"
      _StyleDefs(123) =   "Splits(0).Columns(21).FooterStyle:id=81,.parent=89"
      _StyleDefs(124) =   "Splits(0).Columns(21).EditorStyle:id=82,.parent=91"
      _StyleDefs(125) =   "Splits(0).Columns(22).Style:id=103,.parent=87,.alignment=1"
      _StyleDefs(126) =   "Splits(0).Columns(22).HeadingStyle:id=84,.parent=88"
      _StyleDefs(127) =   "Splits(0).Columns(22).FooterStyle:id=85,.parent=89"
      _StyleDefs(128) =   "Splits(0).Columns(22).EditorStyle:id=86,.parent=91"
      _StyleDefs(129) =   "Splits(0).Columns(23).Style:id=139,.parent=87,.alignment=1"
      _StyleDefs(130) =   "Splits(0).Columns(23).HeadingStyle:id=104,.parent=88"
      _StyleDefs(131) =   "Splits(0).Columns(23).FooterStyle:id=105,.parent=89"
      _StyleDefs(132) =   "Splits(0).Columns(23).EditorStyle:id=106,.parent=91"
      _StyleDefs(133) =   "Splits(0).Columns(24).Style:id=143,.parent=87,.alignment=1"
      _StyleDefs(134) =   "Splits(0).Columns(24).HeadingStyle:id=140,.parent=88"
      _StyleDefs(135) =   "Splits(0).Columns(24).FooterStyle:id=141,.parent=89"
      _StyleDefs(136) =   "Splits(0).Columns(24).EditorStyle:id=142,.parent=91"
      _StyleDefs(137) =   "Splits(0).Columns(25).Style:id=147,.parent=87,.alignment=1"
      _StyleDefs(138) =   "Splits(0).Columns(25).HeadingStyle:id=144,.parent=88"
      _StyleDefs(139) =   "Splits(0).Columns(25).FooterStyle:id=145,.parent=89"
      _StyleDefs(140) =   "Splits(0).Columns(25).EditorStyle:id=146,.parent=91"
      _StyleDefs(141) =   "Splits(0).Columns(26).Style:id=151,.parent=87,.alignment=1"
      _StyleDefs(142) =   "Splits(0).Columns(26).HeadingStyle:id=148,.parent=88"
      _StyleDefs(143) =   "Splits(0).Columns(26).FooterStyle:id=149,.parent=89"
      _StyleDefs(144) =   "Splits(0).Columns(26).EditorStyle:id=150,.parent=91"
      _StyleDefs(145) =   "Splits(0).Columns(27).Style:id=155,.parent=87,.alignment=1"
      _StyleDefs(146) =   "Splits(0).Columns(27).HeadingStyle:id=152,.parent=88"
      _StyleDefs(147) =   "Splits(0).Columns(27).FooterStyle:id=153,.parent=89"
      _StyleDefs(148) =   "Splits(0).Columns(27).EditorStyle:id=154,.parent=91"
      _StyleDefs(149) =   "Splits(0).Columns(28).Style:id=159,.parent=87,.alignment=1"
      _StyleDefs(150) =   "Splits(0).Columns(28).HeadingStyle:id=156,.parent=88"
      _StyleDefs(151) =   "Splits(0).Columns(28).FooterStyle:id=157,.parent=89"
      _StyleDefs(152) =   "Splits(0).Columns(28).EditorStyle:id=158,.parent=91"
      _StyleDefs(153) =   "Splits(0).Columns(29).Style:id=163,.parent=87,.alignment=1"
      _StyleDefs(154) =   "Splits(0).Columns(29).HeadingStyle:id=160,.parent=88"
      _StyleDefs(155) =   "Splits(0).Columns(29).FooterStyle:id=161,.parent=89"
      _StyleDefs(156) =   "Splits(0).Columns(29).EditorStyle:id=162,.parent=91"
      _StyleDefs(157) =   "Splits(0).Columns(30).Style:id=167,.parent=87,.alignment=1"
      _StyleDefs(158) =   "Splits(0).Columns(30).HeadingStyle:id=164,.parent=88"
      _StyleDefs(159) =   "Splits(0).Columns(30).FooterStyle:id=165,.parent=89"
      _StyleDefs(160) =   "Splits(0).Columns(30).EditorStyle:id=166,.parent=91"
      _StyleDefs(161) =   "Splits(0).Columns(31).Style:id=171,.parent=87,.alignment=1"
      _StyleDefs(162) =   "Splits(0).Columns(31).HeadingStyle:id=168,.parent=88"
      _StyleDefs(163) =   "Splits(0).Columns(31).FooterStyle:id=169,.parent=89"
      _StyleDefs(164) =   "Splits(0).Columns(31).EditorStyle:id=170,.parent=91"
      _StyleDefs(165) =   "Splits(0).Columns(32).Style:id=175,.parent=87,.alignment=1"
      _StyleDefs(166) =   "Splits(0).Columns(32).HeadingStyle:id=172,.parent=88"
      _StyleDefs(167) =   "Splits(0).Columns(32).FooterStyle:id=173,.parent=89"
      _StyleDefs(168) =   "Splits(0).Columns(32).EditorStyle:id=174,.parent=91"
      _StyleDefs(169) =   "Splits(0).Columns(33).Style:id=179,.parent=87,.alignment=1"
      _StyleDefs(170) =   "Splits(0).Columns(33).HeadingStyle:id=176,.parent=88"
      _StyleDefs(171) =   "Splits(0).Columns(33).FooterStyle:id=177,.parent=89"
      _StyleDefs(172) =   "Splits(0).Columns(33).EditorStyle:id=178,.parent=91"
      _StyleDefs(173) =   "Splits(0).Columns(34).Style:id=134,.parent=87,.alignment=1"
      _StyleDefs(174) =   "Splits(0).Columns(34).HeadingStyle:id=131,.parent=88"
      _StyleDefs(175) =   "Splits(0).Columns(34).FooterStyle:id=132,.parent=89"
      _StyleDefs(176) =   "Splits(0).Columns(34).EditorStyle:id=133,.parent=91"
      _StyleDefs(177) =   "Splits(0).Columns(35).Style:id=138,.parent=87"
      _StyleDefs(178) =   "Splits(0).Columns(35).HeadingStyle:id=135,.parent=88"
      _StyleDefs(179) =   "Splits(0).Columns(35).FooterStyle:id=136,.parent=89"
      _StyleDefs(180) =   "Splits(0).Columns(35).EditorStyle:id=137,.parent=91"
      _StyleDefs(181) =   "Named:id=33:Normal"
      _StyleDefs(182) =   ":id=33,.parent=0"
      _StyleDefs(183) =   "Named:id=34:Heading"
      _StyleDefs(184) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(185) =   ":id=34,.wraptext=-1"
      _StyleDefs(186) =   "Named:id=35:Footing"
      _StyleDefs(187) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(188) =   "Named:id=36:Selected"
      _StyleDefs(189) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(190) =   "Named:id=37:Caption"
      _StyleDefs(191) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(192) =   "Named:id=38:HighlightRow"
      _StyleDefs(193) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(194) =   "Named:id=39:EvenRow"
      _StyleDefs(195) =   ":id=39,.parent=33"
      _StyleDefs(196) =   "Named:id=40:OddRow"
      _StyleDefs(197) =   ":id=40,.parent=33,.bgcolor=&H40FF00&"
      _StyleDefs(198) =   "Named:id=41:RecordSelector"
      _StyleDefs(199) =   ":id=41,.parent=34"
      _StyleDefs(200) =   "Named:id=42:FilterBar"
      _StyleDefs(201) =   ":id=42,.parent=33"
      _StyleDefs(202) =   "Named:id=13:LockItem"
      _StyleDefs(203) =   ":id=13,.parent=39"
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
      Left            =   7980
      TabIndex        =   10
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
      Left            =   420
      TabIndex        =   9
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
      Left            =   4080
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   3255
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
      Left            =   2460
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "検索"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "画面印刷"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   2
      End
   End
End
Attribute VB_Name = "ODR20101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NAIGAI_CODE()   As String * 1
Private NAIGAI_NAME()   As String

'コンボ用添字
Private Const pcmbSM = 0            '仕向け先

'テキスト用添字
Private Const ptxTOP% = 0
Private Const ptxLAST% = 1

Private Const ptxUSE_YY% = 0
Private Const ptxTANTO_CD% = 1

'ラベル用添字
Private Const plabTANTO_NM% = 0

'コマンドボタン用添字
Private Const FuncSRC% = 0       '検索
'Private Const FuncDEL% = 1       '削除
Private Const FuncEND% = 2       '終了

'ListBox添字
'Private Const plst_DISP% = 0     '表示用データ　Sort順＆Key


'グリッド用定義
Private ORDR_GRID   As New XArrayDB

Private Const Min_Row% = 1              '最小行数
Private Max_Row As Long                 '最大表示行数

Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 35             '最大列数

Private Const Col_ORDR_NO% = 0              '親部品　注文№
Private Const Col_OYA_ITEM% = 1             '親部品コード
Private Const Col_ORDR_QTY% = 2             '注文数量
Private Const Col_TOTAL_QTY% = 3            '合計
Private Const Col_01% = 4                   '01日

Private Const Col_31% = 34                  '31日
Private Const Col_KEY% = 35                 'データＫｅｙ情報

Dim row         As Long                 '対象　行


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

Private Function ERR_CHK(Index As Integer)
Dim sts         As Integer
Dim yn          As Integer

Dim W_Str       As String


    ERR_CHK = True
    
                        '入力文字数チェック
    If LenB(StrConv(Text1(Index), vbFromUnicode)) > Text1(Index).MaxLength Then
        MsgBox "入力した項目は（桁あふれエラー）です。", vbExclamation
        Exit Function
    End If
    
    Select Case Index
        Case ptxTANTO_CD%
            Lab_Dsp(plabTANTO_NM) = ""
            If Trim(Text1(Index)) = "" Then
                MsgBox "担当者を指定して下さい。", vbExclamation
                Exit Function
            End If
            
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(Index))
            Do
                sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound       'レコード無し
                        MsgBox "担当者　未登録！", vbExclamation
                        Exit Function
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                        Beep
                        yn = MsgBox("他で使用中です！<TANTO>" & Chr(13) & Chr(10) & _
                                    "再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                        If yn = vbNo Then Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "TANTO")
                        Exit Function
                End Select
            Loop
            Lab_Dsp(plabTANTO_NM) = Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
            
        Case ptxUSE_YY%
            If Trim(Text1(Index)) = "" Then
                MsgBox "使用年月を指定して下さい。", vbExclamation
                Exit Function
            End If
            
            W_Str = Text1(ptxUSE_YY%) & "/01"
            
            If Not IsDate(W_Str) Then
                MsgBox "使用月エラー！", vbExclamation
                Exit Function
            End If
            
            W_Str = Format(W_Str, "yyyy/mm/dd")
            Text1(ptxUSE_YY%) = Left(W_Str, 7)
            
            If Left(W_Str, 4) < "2000" Then
                MsgBox "使用月　＜　2000年エラー！", vbExclamation
                Exit Function
            End If
            If Left(W_Str, 4) > "2100" Then
                MsgBox "使用月　＞　2100年エラー！", vbExclamation
                Exit Function
            End If
            
            
    End Select
    
    
    ERR_CHK = False
End Function

Private Function Data_Disp()
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_Row       As Long

Dim X_i         As Long
Dim X_j         As Long
Dim W_Key       As String

Dim W_Date      As String
Dim W_Col       As Integer
Dim W_QTY       As Double
Dim W_Str       As String


Dim ODR_KEY_TB()       As String        'Key内容
Dim ODR_QTY_TB()   As String            'Key単位の日別数量


    Data_Disp = True
    
    row = Min_Row - 1
    Call Input_Lock                             '画面項目ロック
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "注文情報　検索・表示中。<Data_Disp>", Me.hwnd, 0)
    DoEvents
    
    Set ORDR_GRID = Nothing
    Erase ODR_KEY_TB
    Erase ODR_QTY_TB
    ReDim Preserve ODR_KEY_TB(0)
    ReDim Preserve ODR_QTY_TB(32, 0)
    Call QTY_TB_CLR(ODR_QTY_TB(), 0)
    
    W_Row = -1
    
    'GW_SIMUKE = Left(Right(Combo1(pcmbSM).Text, 4), 2)
    'GW_JIGYOBU = Right(Combo1(pcmbJI).Text, 1)
    'GW_JIGYOBU = Mid(Right(Combo1(pcmbSM).Text, 4), 3, 1)
    'GW_NAIGAI = Right(Combo1(pcmbSM).Text, 1)
    
    Call UniCode_Conv(K1_ODR_ORDER.SHIMUKE, GW_SIMUKE)
    Call UniCode_Conv(K1_ODR_ORDER.JGYOBU, GW_JIGYOBU)
    Call UniCode_Conv(K1_ODR_ORDER.NAIGAI, GW_NAIGAI)
    W_Str = Left(Text1(ptxUSE_YY), 4) & Right(Text1(ptxUSE_YY), 2)
    Call UniCode_Conv(K1_ODR_ORDER.USE_YM, W_Str)
    
    Call UniCode_Conv(K1_ODR_ORDER.INS_NO, "")
    Call UniCode_Conv(K1_ODR_ORDER.ORDER_NO, "")
    Call UniCode_Conv(K1_ODR_ORDER.BUN_NO, "")
    
    com = BtOpGetGreaterEqual
    Do
        Do
            sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K1_ODR_ORDER, Len(K1_ODR_ORDER), 1)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                    'Beep
                    'MsgBox "指定された工程がありません。"
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    yn = MsgBox("他で使用中です！<注文Ｆ>" & Chr(13) & Chr(10) & _
                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If yn = vbNo Then Exit Function
                Case Else
                    Call File_Error(sts, com, "ODR_ORDER")
                    Exit Function
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU) Then Exit Do
        If Trim(StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI) Then Exit Do
        
        'If CInt(StrConv(ODR_ORDER_REC.BUN_KB, vbUnicode)) = 0 Then
        
        W_Date = Left(StrConv(ODR_ORDER_REC.USE_YM, vbUnicode), 4) & "/" & _
                            Mid(StrConv(ODR_ORDER_REC.USE_YM, vbUnicode), 5, 2)
        If W_Date <> Trim(Text1(ptxUSE_YY)) Then Exit Do
        
        '分納の基情報は表示対象外！
        If CInt(StrConv(ODR_ORDER_REC.BUN_KB, vbUnicode)) = 0 Then
        
            W_Key = StrConv(ODR_ORDER_REC.ORDER_NO, vbUnicode) & StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode)
            W_Row = -1
            For X_i = 0 To UBound(ODR_KEY_TB)
                If Trim(ODR_KEY_TB(X_i)) = "" Then
                    ODR_KEY_TB(X_i) = W_Key
                End If
                If W_Key = ODR_KEY_TB(X_i) Then
                    W_Row = X_i
                    Exit For
                End If
            Next X_i
            
            
            If W_Row = -1 Then
                W_Row = X_i
                ReDim Preserve ODR_KEY_TB(W_Row)
                ReDim Preserve ODR_QTY_TB(32, W_Row)
                Call QTY_TB_CLR(ODR_QTY_TB(), W_Row)
            End If
            
            ODR_KEY_TB(W_Row) = W_Key
            
            DIS_ORDR_QTY = CDbl(Trim(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode)))
            
            W_QTY = CDbl(ODR_QTY_TB(0, W_Row))
            W_QTY = W_QTY + CDbl(DIS_ORDR_QTY)
            ODR_QTY_TB(0, W_Row) = W_QTY
            
            W_Date = ""
            'DIS_SUM_QTY = DIS_ORDR_QTY
            'W_Date = Left(StrConv(ODR_ORDER_REC.CYUMON_DT, vbUnicode), 4) & "/" & _
            '                Mid(StrConv(ODR_ORDER_REC.CYUMON_DT, vbUnicode), 5, 2) & "/" & _
            '                    Right(StrConv(ODR_ORDER_REC.CYUMON_DT, vbUnicode), 2)
                
            If Trim(StrConv(ODR_ORDER_REC.KAITO_DT, vbUnicode)) <> "" Then
                    
                W_Date = Left(StrConv(ODR_ORDER_REC.KAITO_DT, vbUnicode), 4) & "/" & _
                            Mid(StrConv(ODR_ORDER_REC.KAITO_DT, vbUnicode), 5, 2) & "/" & _
                                Right(StrConv(ODR_ORDER_REC.KAITO_DT, vbUnicode), 2)
                DIS_SUM_QTY = DIS_ORDR_QTY
            End If
                
            If Trim(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode)) <> "" Then
                    
                W_Date = Left(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode), 4) & "/" & _
                            Mid(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode), 5, 2) & "/" & _
                                Right(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode), 2)
                DIS_SUM_QTY = DIS_ORDR_QTY
            End If
            If Trim(W_Date) <> "" Then
                W_Col = CInt(Right(W_Date, 2))
                
                W_QTY = CDbl(ODR_QTY_TB(W_Col, W_Row))
                W_QTY = W_QTY + CDbl(DIS_SUM_QTY)
                ODR_QTY_TB(W_Col, W_Row) = W_QTY
                
                W_QTY = CDbl(ODR_QTY_TB(32, W_Row))
                W_QTY = W_QTY + CDbl(DIS_SUM_QTY)
                ODR_QTY_TB(32, W_Row) = W_QTY
            End If
        End If
        com = BtOpGetNext
        
    Loop
    
    
    For X_i = 0 To UBound(ODR_KEY_TB)
        If Trim(ODR_KEY_TB(X_i)) <> "" Then
        
            DIS_ORDR_NO = Left(ODR_KEY_TB(X_i), UBound(ODR_ORDER_REC.ORDER_NO) + 1)
            
            
            DIS_OYA_ITEM = Right(ODR_KEY_TB(X_i), UBound(ODR_ORDER_REC.HIN_GAI) + 1)
            
            DIS_ORDR_QTY = ODR_QTY_TB(0, X_i)
            
            W_QTY = 0
            For X_j = 1 To 31
                DIS_QTY(X_j) = ODR_QTY_TB(X_j, X_i)
            Next X_j
            
            DIS_SUM_QTY = ODR_QTY_TB(32, X_i)
            
            row = X_i + 1
            
            If Grid_Set_Proc() Then
                Exit Function
            End If
        
        End If
        
    Next X_i
    
    
    Set TDBGrid1.Array = ORDR_GRID
    
    TDBGrid1.style.Locked = True
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    TDBGrid1.Bookmark = 1
    
    If ORDR_GRID.Count(1) = 0 Then
        hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "注文情報　表示終了（対象データ無し）", Me.hwnd, 0)
    Else
        hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "注文情報　表示終了（全 " & ORDR_GRID.Count(1) & " 行）", Me.hwnd, 0)
    End If
    
    DoEvents
    
    Call Input_UnLock                             '画面項目ロック
    
    DoEvents
    
    Data_Disp = False
    
        
    
End Function
Private Sub QTY_TB_CLR(QTY_TB() As String, X_j As Long)
Dim X_i     As Integer
    For X_i = 0 To 32
        QTY_TB(X_i, X_j) = "0"
    Next X_i
End Sub
Private Function Grid_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   グリッド表示（移動歴データ内容）
'               Row   行数
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim W_Col       As Integer


    Grid_Set_Proc = True

    ORDR_GRID.ReDim Min_Row, row, Min_Col, Max_Col

'Private Const Col_ORDR_NO% = 0              '親部品　注文№
'Private Const Col_OYA_ITEM% = 1             '親部品コード
'Private Const Col_ORDR_QTY% = 2             '注文数量
'Private Const Col_TOTAL_QTY% = 3            '合計
'Private Const Col_01% = 4                   '01日

'Private Const Col_31% = 34                  '31日
'Private Const Col_KEY% = 35                 'データＫｅｙ情報

    ORDR_GRID(row, Col_ORDR_NO%) = DIS_ORDR_NO
    ORDR_GRID(row, Col_OYA_ITEM%) = DIS_OYA_ITEM
    ORDR_GRID(row, Col_ORDR_QTY%) = DIS_ORDR_QTY
    ORDR_GRID(row, Col_TOTAL_QTY%) = DIS_SUM_QTY
    
    For W_Col = 1 To 31
        If CDbl(DIS_QTY(W_Col)) <> 0 Then
            ORDR_GRID(row, W_Col + 3) = DIS_QTY(W_Col)
        End If
    Next W_Col
    
    Grid_Set_Proc = False

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    ODR20101.MousePointer = vbHourglass

    Call Ctrl_Lock(ODR20101)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(ODR20101)


    ODR20101.MousePointer = vbDefault

End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Tab_Ctrl(Shift)        '移動
    End If

End Sub

Private Sub Command1_Click(Index As Integer)
Dim yn      As Integer

    Select Case Index
    
        Case FuncSRC%
            
            
            '検索＆表示処理
            If Data_Disp Then
                MsgBox "指定条件の注文情報表示　失敗！", vbExclamation
                Call Text1_GotFocus(ptxTOP%)
                Text1(ptxTOP%).SetFocus
                Exit Sub
            End If
            
            If ORDR_GRID.Count(1) = 0 Then
                Text1(ptxUSE_YY).SetFocus
            Else
                TDBGrid1.SetFocus
            End If
            
            Exit Sub

        Case FuncEND%
            'yn = MsgBox("終了しますか？", vbYesNo + vbDefaultButton1 + vbQuestion, "確認入力")
            yn = vbYes
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            Unload Me
    
    End Select

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




'コモンコントロールを初期化する
cc.dwSize = Len(cc)
cc.dwICC = ICC_BAR_CLASSES

'ステータスウィンドウを作成する
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "親部品　注文情報照会", Me.hwnd, 0)
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
    
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        GW_PC_NM = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        GW_PC_NM = "???"
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
    
                                '親品番注文Ｆ　ＯＰＥＮ
    If ODR_ORDER_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                               
                               
    Max_Row = 25000
    
'テキストを設定する
    W_Date = Format(Date, "yyyy/mm/dd")
    Text1(ptxUSE_YY) = Left(W_Date, 7)
    
    '2008/10/07 最初の使用月を下記に変更。
    If GetIni("PR00030", "LAST_SHIME_DT01", "P_SYS", c) Then
        GW_TOUGETU = Left(Format(Date, "yyyymmdd"), 6)
    Else
        GW_TOUGETU = Left(Format(Trim(c), "yyyymmdd"), 6)
    End If
    
    Text1(ptxUSE_YY) = Left(GW_TOUGETU, 4) & "/" & Right(GW_TOUGETU, 2)
    
    
    
    
    
    
    
    
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
    GW_JIGYOBU = Mid(Right(Combo1(pcmbSM).Text, 4), 3, 1)
    GW_NAIGAI = Right(Combo1(pcmbSM).Text, 1)
    
    'Combo1(pcmbSM).SetFocus
    
    Text1(ptxTOP).SetFocus
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim yn      As Integer

    If UnloadMode = 1 Then Exit Sub
    
    yn = MsgBox("終了しますか？", vbYesNo + vbDefaultButton1 + vbQuestion, "確認入力")
    'yn = vbYes
    If yn = vbNo Then
        Cancel = 1
        Exit Sub
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts     As Integer


    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "TANTO")
        End If
    End If
    
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "P_CODE")
        End If
    End If
    
    'sts = BTRV(BtOpClose, ODR_TEMP1_POS, ODR_TEMP1_REC, Len(ODR_TEMP1_REC), K0_ODR_TEMP1, Len(K0_ODR_TEMP1), 0)
    
    sts = BTRV(BtOpClose, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_ORDER, Len(K0_ODR_ORDER), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_ORDER")
        End If
    End If
    
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "P_KANRI")
        End If
    End If



    End
End Sub

Private Sub SHORI_Click(Index As Integer)
Dim yn      As Integer


    Select Case Index
        Case 0      '検索
            Call Command1_Click(FuncSRC)
            
        Case 1      '画面印刷
            yn = MsgBox("画面印刷しますか？", vbYesNo + vbDefaultButton2 + vbQuestion, "確認入力")
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)
        
        
        Case 2      '終了
            Call Command1_Click(FuncEND)
    
    End Select


End Sub

Private Sub TDBGrid1_DblClick()

    If TDBGrid1.Bookmark = -1 Then
    Else
        
        'ODR20102.Show vbModal
        
        'If KENPIN_Update_Proc() Then
        '    Unload Me
        'End If
    End If
    
    '再表示
'    If List_Disp Then
'        Unload Me
'    End If


End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
Dim yn          As Integer
Dim W_Index     As Integer

    'TDBGrid1.Bookmark = -1
    W_Index = ColIndex
    
    If row <= 1 Then Exit Sub
    
    If ColIndex <= Col_TOTAL_QTY Then
        yn = MsgBox("並べ換えますか？", vbYesNo + vbExclamation, "確認入力")
        If yn <> vbYes Then Exit Sub
    End If
    'Set ORDR_GRID = TDBGrid1.Array
    
    Select Case ColIndex
        Case Col_ORDR_NO%           '親部品　注文№
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_ORDR_NO%, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM%, XORDER_ASCEND, XTYPE_STRING
                
        Case Col_OYA_ITEM%          '親部品コード
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_OYA_ITEM%, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ORDR_NO%, XORDER_ASCEND, XTYPE_STRING
        
        Case Col_ORDR_QTY%          '親部品　注文数量
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_ORDR_QTY%, XORDER_ASCEND, XTYPE_LONG, _
                        Col_ORDR_NO%, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM%, XORDER_ASCEND, XTYPE_STRING
        
        Case Col_TOTAL_QTY%         '合計
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_TOTAL_QTY%, XORDER_ASCEND, XTYPE_LONG, _
                        Col_ORDR_NO%, XORDER_ASCEND, XTYPE_STRING, _
                        Col_OYA_ITEM%, XORDER_ASCEND, XTYPE_STRING
        
            
        Case Else
            'MsgBox "並べ替指定 除外項目！", vbExclamation
            Exit Sub
        
    End Select

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
    
    
    Call Tab_Ctrl(Shift)    '移動
    
End Sub

