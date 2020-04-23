VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form ODR30201 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "子部品　在庫推移照会"
   ClientHeight    =   10020
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   19110
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
   ScaleHeight     =   10020
   ScaleWidth      =   19110
   StartUpPosition =   2  '画面の中央
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
      Index           =   3
      Left            =   8715
      MaxLength       =   20
      TabIndex        =   15
      Top             =   720
      Width           =   1305
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
      Index           =   2
      Left            =   7140
      MaxLength       =   20
      TabIndex        =   2
      Top             =   720
      Width           =   1305
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
      TabIndex        =   12
      Top             =   120
      Width           =   1800
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Index           =   0
      Left            =   11700
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Index           =   1
      Left            =   9000
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   1155
      MaxLength       =   20
      TabIndex        =   0
      Top             =   720
      Width           =   2655
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
      Left            =   4830
      MaxLength       =   20
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   4800
      ScaleHeight     =   195
      ScaleWidth      =   270
      TabIndex        =   7
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
      Index           =   1
      Left            =   2415
      TabIndex        =   3
      Top             =   120
      Width           =   1800
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   8415
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   18540
      _ExtentX        =   32703
      _ExtentY        =   14843
      _LayoutType     =   4
      _RowHeight      =   15
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "仕入先"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "子部品"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "子部品名"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
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
      Columns(38)._VlistStyle=   0
      Columns(38)._MaxComboItems=   5
      Columns(38).DataField=   ""
      Columns(38)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(39)._VlistStyle=   0
      Columns(39)._MaxComboItems=   5
      Columns(39).DataField=   ""
      Columns(39)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(40)._VlistStyle=   0
      Columns(40)._MaxComboItems=   5
      Columns(40).DataField=   ""
      Columns(40)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(41)._VlistStyle=   0
      Columns(41)._MaxComboItems=   5
      Columns(41).DataField=   ""
      Columns(41)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(42)._VlistStyle=   0
      Columns(42)._MaxComboItems=   5
      Columns(42).DataField=   ""
      Columns(42)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(43)._VlistStyle=   0
      Columns(43)._MaxComboItems=   5
      Columns(43).DataField=   ""
      Columns(43)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(44)._VlistStyle=   0
      Columns(44)._MaxComboItems=   5
      Columns(44).DataField=   ""
      Columns(44)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(45)._VlistStyle=   0
      Columns(45)._MaxComboItems=   5
      Columns(45).DataField=   ""
      Columns(45)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(46)._VlistStyle=   0
      Columns(46)._MaxComboItems=   5
      Columns(46).DataField=   ""
      Columns(46)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(47)._VlistStyle=   0
      Columns(47)._MaxComboItems=   5
      Columns(47).DataField=   ""
      Columns(47)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(48)._VlistStyle=   0
      Columns(48)._MaxComboItems=   5
      Columns(48).DataField=   ""
      Columns(48)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(49)._VlistStyle=   0
      Columns(49)._MaxComboItems=   5
      Columns(49).DataField=   ""
      Columns(49)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(50)._VlistStyle=   0
      Columns(50)._MaxComboItems=   5
      Columns(50).DataField=   ""
      Columns(50)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(51)._VlistStyle=   0
      Columns(51)._MaxComboItems=   5
      Columns(51).DataField=   ""
      Columns(51)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(52)._VlistStyle=   0
      Columns(52)._MaxComboItems=   5
      Columns(52).DataField=   ""
      Columns(52)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(53)._VlistStyle=   0
      Columns(53)._MaxComboItems=   5
      Columns(53).DataField=   ""
      Columns(53)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(54)._VlistStyle=   0
      Columns(54)._MaxComboItems=   5
      Columns(54).DataField=   ""
      Columns(54)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(55)._VlistStyle=   0
      Columns(55)._MaxComboItems=   5
      Columns(55).DataField=   ""
      Columns(55)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(56)._VlistStyle=   0
      Columns(56)._MaxComboItems=   5
      Columns(56).DataField=   ""
      Columns(56)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(57)._VlistStyle=   0
      Columns(57)._MaxComboItems=   5
      Columns(57).DataField=   ""
      Columns(57)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(58)._VlistStyle=   0
      Columns(58)._MaxComboItems=   5
      Columns(58).DataField=   ""
      Columns(58)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(59)._VlistStyle=   0
      Columns(59)._MaxComboItems=   5
      Columns(59).DataField=   ""
      Columns(59)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(60)._VlistStyle=   0
      Columns(60)._MaxComboItems=   5
      Columns(60).DataField=   ""
      Columns(60)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(61)._VlistStyle=   0
      Columns(61)._MaxComboItems=   5
      Columns(61).DataField=   ""
      Columns(61)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(62)._VlistStyle=   0
      Columns(62)._MaxComboItems=   5
      Columns(62).DataField=   ""
      Columns(62)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(63)._VlistStyle=   0
      Columns(63)._MaxComboItems=   5
      Columns(63).DataField=   ""
      Columns(63)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(64)._VlistStyle=   0
      Columns(64)._MaxComboItems=   5
      Columns(64).DataField=   ""
      Columns(64)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(65)._VlistStyle=   0
      Columns(65)._MaxComboItems=   5
      Columns(65).DataField=   ""
      Columns(65)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(66)._VlistStyle=   0
      Columns(66)._MaxComboItems=   5
      Columns(66).DataField=   ""
      Columns(66)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(67)._VlistStyle=   0
      Columns(67)._MaxComboItems=   5
      Columns(67).DataField=   ""
      Columns(67)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(68)._VlistStyle=   0
      Columns(68)._MaxComboItems=   5
      Columns(68).DataField=   ""
      Columns(68)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(69)._VlistStyle=   0
      Columns(69)._MaxComboItems=   5
      Columns(69).DataField=   ""
      Columns(69)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(70)._VlistStyle=   0
      Columns(70)._MaxComboItems=   5
      Columns(70).DataField=   ""
      Columns(70)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(71)._VlistStyle=   0
      Columns(71)._MaxComboItems=   5
      Columns(71).DataField=   ""
      Columns(71)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(72)._VlistStyle=   0
      Columns(72)._MaxComboItems=   5
      Columns(72).DataField=   ""
      Columns(72)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(73)._VlistStyle=   0
      Columns(73)._MaxComboItems=   5
      Columns(73).DataField=   ""
      Columns(73)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(74)._VlistStyle=   0
      Columns(74)._MaxComboItems=   5
      Columns(74).DataField=   ""
      Columns(74)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(75)._VlistStyle=   0
      Columns(75)._MaxComboItems=   5
      Columns(75).DataField=   ""
      Columns(75)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(76)._VlistStyle=   0
      Columns(76)._MaxComboItems=   5
      Columns(76).DataField=   ""
      Columns(76)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(77)._VlistStyle=   0
      Columns(77)._MaxComboItems=   5
      Columns(77).DataField=   ""
      Columns(77)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(78)._VlistStyle=   0
      Columns(78)._MaxComboItems=   5
      Columns(78).DataField=   ""
      Columns(78)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(79)._VlistStyle=   0
      Columns(79)._MaxComboItems=   5
      Columns(79).DataField=   ""
      Columns(79)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(80)._VlistStyle=   0
      Columns(80)._MaxComboItems=   5
      Columns(80).DataField=   ""
      Columns(80)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(81)._VlistStyle=   0
      Columns(81)._MaxComboItems=   5
      Columns(81).DataField=   ""
      Columns(81)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(82)._VlistStyle=   0
      Columns(82)._MaxComboItems=   5
      Columns(82).DataField=   ""
      Columns(82)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(83)._VlistStyle=   0
      Columns(83)._MaxComboItems=   5
      Columns(83).DataField=   ""
      Columns(83)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(84)._VlistStyle=   0
      Columns(84)._MaxComboItems=   5
      Columns(84).DataField=   ""
      Columns(84)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(85)._VlistStyle=   0
      Columns(85)._MaxComboItems=   5
      Columns(85).DataField=   ""
      Columns(85)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(86)._VlistStyle=   0
      Columns(86)._MaxComboItems=   5
      Columns(86).DataField=   ""
      Columns(86)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(87)._VlistStyle=   0
      Columns(87)._MaxComboItems=   5
      Columns(87).DataField=   ""
      Columns(87)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(88)._VlistStyle=   0
      Columns(88)._MaxComboItems=   5
      Columns(88).DataField=   ""
      Columns(88)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(89)._VlistStyle=   0
      Columns(89)._MaxComboItems=   5
      Columns(89).DataField=   ""
      Columns(89)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(90)._VlistStyle=   0
      Columns(90)._MaxComboItems=   5
      Columns(90).DataField=   ""
      Columns(90)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(91)._VlistStyle=   0
      Columns(91)._MaxComboItems=   5
      Columns(91).DataField=   ""
      Columns(91)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(92)._VlistStyle=   0
      Columns(92)._MaxComboItems=   5
      Columns(92).DataField=   ""
      Columns(92)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(93)._VlistStyle=   0
      Columns(93)._MaxComboItems=   5
      Columns(93).DataField=   ""
      Columns(93)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   94
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=94"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2143"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2011"
      Splits(0)._ColumnProps(4)=   "Column(0).FetchStyle=1"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2117"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1984"
      Splits(0)._ColumnProps(9)=   "Column(1).FetchStyle=1"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2170"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2037"
      Splits(0)._ColumnProps(14)=   "Column(2).FetchStyle=1"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=1720"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1588"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(20)=   "Column(3).WrapText=1"
      Splits(0)._ColumnProps(21)=   "Column(3).FetchStyle=1"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=1270"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=1138"
      Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=8194"
      Splits(0)._ColumnProps(27)=   "Column(4).FetchStyle=1"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(29)=   "Column(5).Width=1270"
      Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=1138"
      Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(33)=   "Column(5).FetchStyle=1"
      Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(35)=   "Column(6).Width=1270"
      Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=1138"
      Splits(0)._ColumnProps(38)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(39)=   "Column(6).FetchStyle=1"
      Splits(0)._ColumnProps(40)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(41)=   "Column(7).Width=1270"
      Splits(0)._ColumnProps(42)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(7)._WidthInPix=1138"
      Splits(0)._ColumnProps(44)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(45)=   "Column(7).FetchStyle=1"
      Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(47)=   "Column(8).Width=1270"
      Splits(0)._ColumnProps(48)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(8)._WidthInPix=1138"
      Splits(0)._ColumnProps(50)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(51)=   "Column(8).FetchStyle=1"
      Splits(0)._ColumnProps(52)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(53)=   "Column(9).Width=1270"
      Splits(0)._ColumnProps(54)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(9)._WidthInPix=1138"
      Splits(0)._ColumnProps(56)=   "Column(9)._ColStyle=2"
      Splits(0)._ColumnProps(57)=   "Column(9).FetchStyle=1"
      Splits(0)._ColumnProps(58)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(59)=   "Column(10).Width=1270"
      Splits(0)._ColumnProps(60)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(10)._WidthInPix=1138"
      Splits(0)._ColumnProps(62)=   "Column(10)._ColStyle=2"
      Splits(0)._ColumnProps(63)=   "Column(10).FetchStyle=1"
      Splits(0)._ColumnProps(64)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(65)=   "Column(11).Width=1270"
      Splits(0)._ColumnProps(66)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(11)._WidthInPix=1138"
      Splits(0)._ColumnProps(68)=   "Column(11)._ColStyle=2"
      Splits(0)._ColumnProps(69)=   "Column(11).FetchStyle=1"
      Splits(0)._ColumnProps(70)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(71)=   "Column(12).Width=1270"
      Splits(0)._ColumnProps(72)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(12)._WidthInPix=1138"
      Splits(0)._ColumnProps(74)=   "Column(12)._ColStyle=2"
      Splits(0)._ColumnProps(75)=   "Column(12).FetchStyle=1"
      Splits(0)._ColumnProps(76)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(77)=   "Column(13).Width=1270"
      Splits(0)._ColumnProps(78)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(79)=   "Column(13)._WidthInPix=1138"
      Splits(0)._ColumnProps(80)=   "Column(13)._ColStyle=2"
      Splits(0)._ColumnProps(81)=   "Column(13).FetchStyle=1"
      Splits(0)._ColumnProps(82)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(83)=   "Column(14).Width=1270"
      Splits(0)._ColumnProps(84)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(85)=   "Column(14)._WidthInPix=1138"
      Splits(0)._ColumnProps(86)=   "Column(14)._ColStyle=2"
      Splits(0)._ColumnProps(87)=   "Column(14).FetchStyle=1"
      Splits(0)._ColumnProps(88)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(89)=   "Column(15).Width=1270"
      Splits(0)._ColumnProps(90)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(15)._WidthInPix=1138"
      Splits(0)._ColumnProps(92)=   "Column(15)._ColStyle=2"
      Splits(0)._ColumnProps(93)=   "Column(15).FetchStyle=1"
      Splits(0)._ColumnProps(94)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(95)=   "Column(16).Width=1270"
      Splits(0)._ColumnProps(96)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(97)=   "Column(16)._WidthInPix=1138"
      Splits(0)._ColumnProps(98)=   "Column(16)._ColStyle=2"
      Splits(0)._ColumnProps(99)=   "Column(16).FetchStyle=1"
      Splits(0)._ColumnProps(100)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(101)=   "Column(17).Width=1270"
      Splits(0)._ColumnProps(102)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(103)=   "Column(17)._WidthInPix=1138"
      Splits(0)._ColumnProps(104)=   "Column(17)._ColStyle=2"
      Splits(0)._ColumnProps(105)=   "Column(17).FetchStyle=1"
      Splits(0)._ColumnProps(106)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(107)=   "Column(18).Width=1270"
      Splits(0)._ColumnProps(108)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(109)=   "Column(18)._WidthInPix=1138"
      Splits(0)._ColumnProps(110)=   "Column(18)._ColStyle=2"
      Splits(0)._ColumnProps(111)=   "Column(18).FetchStyle=1"
      Splits(0)._ColumnProps(112)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(113)=   "Column(19).Width=1270"
      Splits(0)._ColumnProps(114)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(115)=   "Column(19)._WidthInPix=1138"
      Splits(0)._ColumnProps(116)=   "Column(19)._ColStyle=2"
      Splits(0)._ColumnProps(117)=   "Column(19).FetchStyle=1"
      Splits(0)._ColumnProps(118)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(119)=   "Column(20).Width=1270"
      Splits(0)._ColumnProps(120)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(121)=   "Column(20)._WidthInPix=1138"
      Splits(0)._ColumnProps(122)=   "Column(20)._ColStyle=2"
      Splits(0)._ColumnProps(123)=   "Column(20).FetchStyle=1"
      Splits(0)._ColumnProps(124)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(125)=   "Column(21).Width=1270"
      Splits(0)._ColumnProps(126)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(127)=   "Column(21)._WidthInPix=1138"
      Splits(0)._ColumnProps(128)=   "Column(21)._ColStyle=2"
      Splits(0)._ColumnProps(129)=   "Column(21).FetchStyle=1"
      Splits(0)._ColumnProps(130)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(131)=   "Column(22).Width=1270"
      Splits(0)._ColumnProps(132)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(133)=   "Column(22)._WidthInPix=1138"
      Splits(0)._ColumnProps(134)=   "Column(22)._ColStyle=2"
      Splits(0)._ColumnProps(135)=   "Column(22).FetchStyle=1"
      Splits(0)._ColumnProps(136)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(137)=   "Column(23).Width=1270"
      Splits(0)._ColumnProps(138)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(139)=   "Column(23)._WidthInPix=1138"
      Splits(0)._ColumnProps(140)=   "Column(23)._ColStyle=2"
      Splits(0)._ColumnProps(141)=   "Column(23).FetchStyle=1"
      Splits(0)._ColumnProps(142)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(143)=   "Column(24).Width=1270"
      Splits(0)._ColumnProps(144)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(145)=   "Column(24)._WidthInPix=1138"
      Splits(0)._ColumnProps(146)=   "Column(24)._ColStyle=2"
      Splits(0)._ColumnProps(147)=   "Column(24).FetchStyle=1"
      Splits(0)._ColumnProps(148)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(149)=   "Column(25).Width=1270"
      Splits(0)._ColumnProps(150)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(151)=   "Column(25)._WidthInPix=1138"
      Splits(0)._ColumnProps(152)=   "Column(25)._ColStyle=2"
      Splits(0)._ColumnProps(153)=   "Column(25).FetchStyle=1"
      Splits(0)._ColumnProps(154)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(155)=   "Column(26).Width=1270"
      Splits(0)._ColumnProps(156)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(157)=   "Column(26)._WidthInPix=1138"
      Splits(0)._ColumnProps(158)=   "Column(26)._ColStyle=2"
      Splits(0)._ColumnProps(159)=   "Column(26).FetchStyle=1"
      Splits(0)._ColumnProps(160)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(161)=   "Column(27).Width=1270"
      Splits(0)._ColumnProps(162)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(163)=   "Column(27)._WidthInPix=1138"
      Splits(0)._ColumnProps(164)=   "Column(27)._ColStyle=2"
      Splits(0)._ColumnProps(165)=   "Column(27).FetchStyle=1"
      Splits(0)._ColumnProps(166)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(167)=   "Column(28).Width=1270"
      Splits(0)._ColumnProps(168)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(169)=   "Column(28)._WidthInPix=1138"
      Splits(0)._ColumnProps(170)=   "Column(28)._ColStyle=2"
      Splits(0)._ColumnProps(171)=   "Column(28).FetchStyle=1"
      Splits(0)._ColumnProps(172)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(173)=   "Column(29).Width=1270"
      Splits(0)._ColumnProps(174)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(175)=   "Column(29)._WidthInPix=1138"
      Splits(0)._ColumnProps(176)=   "Column(29)._ColStyle=2"
      Splits(0)._ColumnProps(177)=   "Column(29).FetchStyle=1"
      Splits(0)._ColumnProps(178)=   "Column(29).Order=30"
      Splits(0)._ColumnProps(179)=   "Column(30).Width=1270"
      Splits(0)._ColumnProps(180)=   "Column(30).DividerColor=0"
      Splits(0)._ColumnProps(181)=   "Column(30)._WidthInPix=1138"
      Splits(0)._ColumnProps(182)=   "Column(30)._ColStyle=2"
      Splits(0)._ColumnProps(183)=   "Column(30).FetchStyle=1"
      Splits(0)._ColumnProps(184)=   "Column(30).Order=31"
      Splits(0)._ColumnProps(185)=   "Column(31).Width=1270"
      Splits(0)._ColumnProps(186)=   "Column(31).DividerColor=0"
      Splits(0)._ColumnProps(187)=   "Column(31)._WidthInPix=1138"
      Splits(0)._ColumnProps(188)=   "Column(31)._ColStyle=2"
      Splits(0)._ColumnProps(189)=   "Column(31).FetchStyle=1"
      Splits(0)._ColumnProps(190)=   "Column(31).Order=32"
      Splits(0)._ColumnProps(191)=   "Column(32).Width=1270"
      Splits(0)._ColumnProps(192)=   "Column(32).DividerColor=0"
      Splits(0)._ColumnProps(193)=   "Column(32)._WidthInPix=1138"
      Splits(0)._ColumnProps(194)=   "Column(32)._ColStyle=2"
      Splits(0)._ColumnProps(195)=   "Column(32).FetchStyle=1"
      Splits(0)._ColumnProps(196)=   "Column(32).Order=33"
      Splits(0)._ColumnProps(197)=   "Column(33).Width=1270"
      Splits(0)._ColumnProps(198)=   "Column(33).DividerColor=0"
      Splits(0)._ColumnProps(199)=   "Column(33)._WidthInPix=1138"
      Splits(0)._ColumnProps(200)=   "Column(33)._ColStyle=2"
      Splits(0)._ColumnProps(201)=   "Column(33).FetchStyle=1"
      Splits(0)._ColumnProps(202)=   "Column(33).Order=34"
      Splits(0)._ColumnProps(203)=   "Column(34).Width=1270"
      Splits(0)._ColumnProps(204)=   "Column(34).DividerColor=0"
      Splits(0)._ColumnProps(205)=   "Column(34)._WidthInPix=1138"
      Splits(0)._ColumnProps(206)=   "Column(34)._ColStyle=2"
      Splits(0)._ColumnProps(207)=   "Column(34).FetchStyle=1"
      Splits(0)._ColumnProps(208)=   "Column(34).Order=35"
      Splits(0)._ColumnProps(209)=   "Column(35).Width=1270"
      Splits(0)._ColumnProps(210)=   "Column(35).DividerColor=0"
      Splits(0)._ColumnProps(211)=   "Column(35)._WidthInPix=1138"
      Splits(0)._ColumnProps(212)=   "Column(35)._ColStyle=2"
      Splits(0)._ColumnProps(213)=   "Column(35).FetchStyle=1"
      Splits(0)._ColumnProps(214)=   "Column(35).Order=36"
      Splits(0)._ColumnProps(215)=   "Column(36).Width=1270"
      Splits(0)._ColumnProps(216)=   "Column(36).DividerColor=0"
      Splits(0)._ColumnProps(217)=   "Column(36)._WidthInPix=1138"
      Splits(0)._ColumnProps(218)=   "Column(36)._ColStyle=2"
      Splits(0)._ColumnProps(219)=   "Column(36).FetchStyle=1"
      Splits(0)._ColumnProps(220)=   "Column(36).Order=37"
      Splits(0)._ColumnProps(221)=   "Column(37).Width=1270"
      Splits(0)._ColumnProps(222)=   "Column(37).DividerColor=0"
      Splits(0)._ColumnProps(223)=   "Column(37)._WidthInPix=1138"
      Splits(0)._ColumnProps(224)=   "Column(37)._ColStyle=2"
      Splits(0)._ColumnProps(225)=   "Column(37).FetchStyle=1"
      Splits(0)._ColumnProps(226)=   "Column(37).Order=38"
      Splits(0)._ColumnProps(227)=   "Column(38).Width=1270"
      Splits(0)._ColumnProps(228)=   "Column(38).DividerColor=0"
      Splits(0)._ColumnProps(229)=   "Column(38)._WidthInPix=1138"
      Splits(0)._ColumnProps(230)=   "Column(38)._ColStyle=2"
      Splits(0)._ColumnProps(231)=   "Column(38).FetchStyle=1"
      Splits(0)._ColumnProps(232)=   "Column(38).Order=39"
      Splits(0)._ColumnProps(233)=   "Column(39).Width=1270"
      Splits(0)._ColumnProps(234)=   "Column(39).DividerColor=0"
      Splits(0)._ColumnProps(235)=   "Column(39)._WidthInPix=1138"
      Splits(0)._ColumnProps(236)=   "Column(39)._ColStyle=2"
      Splits(0)._ColumnProps(237)=   "Column(39).FetchStyle=1"
      Splits(0)._ColumnProps(238)=   "Column(39).Order=40"
      Splits(0)._ColumnProps(239)=   "Column(40).Width=1270"
      Splits(0)._ColumnProps(240)=   "Column(40).DividerColor=0"
      Splits(0)._ColumnProps(241)=   "Column(40)._WidthInPix=1138"
      Splits(0)._ColumnProps(242)=   "Column(40)._ColStyle=2"
      Splits(0)._ColumnProps(243)=   "Column(40).FetchStyle=1"
      Splits(0)._ColumnProps(244)=   "Column(40).Order=41"
      Splits(0)._ColumnProps(245)=   "Column(41).Width=1270"
      Splits(0)._ColumnProps(246)=   "Column(41).DividerColor=0"
      Splits(0)._ColumnProps(247)=   "Column(41)._WidthInPix=1138"
      Splits(0)._ColumnProps(248)=   "Column(41)._ColStyle=2"
      Splits(0)._ColumnProps(249)=   "Column(41).FetchStyle=1"
      Splits(0)._ColumnProps(250)=   "Column(41).Order=42"
      Splits(0)._ColumnProps(251)=   "Column(42).Width=1270"
      Splits(0)._ColumnProps(252)=   "Column(42).DividerColor=0"
      Splits(0)._ColumnProps(253)=   "Column(42)._WidthInPix=1138"
      Splits(0)._ColumnProps(254)=   "Column(42)._ColStyle=2"
      Splits(0)._ColumnProps(255)=   "Column(42).FetchStyle=1"
      Splits(0)._ColumnProps(256)=   "Column(42).Order=43"
      Splits(0)._ColumnProps(257)=   "Column(43).Width=1270"
      Splits(0)._ColumnProps(258)=   "Column(43).DividerColor=0"
      Splits(0)._ColumnProps(259)=   "Column(43)._WidthInPix=1138"
      Splits(0)._ColumnProps(260)=   "Column(43)._ColStyle=2"
      Splits(0)._ColumnProps(261)=   "Column(43).FetchStyle=1"
      Splits(0)._ColumnProps(262)=   "Column(43).Order=44"
      Splits(0)._ColumnProps(263)=   "Column(44).Width=1270"
      Splits(0)._ColumnProps(264)=   "Column(44).DividerColor=0"
      Splits(0)._ColumnProps(265)=   "Column(44)._WidthInPix=1138"
      Splits(0)._ColumnProps(266)=   "Column(44)._ColStyle=2"
      Splits(0)._ColumnProps(267)=   "Column(44).FetchStyle=1"
      Splits(0)._ColumnProps(268)=   "Column(44).Order=45"
      Splits(0)._ColumnProps(269)=   "Column(45).Width=1270"
      Splits(0)._ColumnProps(270)=   "Column(45).DividerColor=0"
      Splits(0)._ColumnProps(271)=   "Column(45)._WidthInPix=1138"
      Splits(0)._ColumnProps(272)=   "Column(45)._ColStyle=2"
      Splits(0)._ColumnProps(273)=   "Column(45).FetchStyle=1"
      Splits(0)._ColumnProps(274)=   "Column(45).Order=46"
      Splits(0)._ColumnProps(275)=   "Column(46).Width=1270"
      Splits(0)._ColumnProps(276)=   "Column(46).DividerColor=0"
      Splits(0)._ColumnProps(277)=   "Column(46)._WidthInPix=1138"
      Splits(0)._ColumnProps(278)=   "Column(46)._ColStyle=2"
      Splits(0)._ColumnProps(279)=   "Column(46).FetchStyle=1"
      Splits(0)._ColumnProps(280)=   "Column(46).Order=47"
      Splits(0)._ColumnProps(281)=   "Column(47).Width=1270"
      Splits(0)._ColumnProps(282)=   "Column(47).DividerColor=0"
      Splits(0)._ColumnProps(283)=   "Column(47)._WidthInPix=1138"
      Splits(0)._ColumnProps(284)=   "Column(47)._ColStyle=2"
      Splits(0)._ColumnProps(285)=   "Column(47).FetchStyle=1"
      Splits(0)._ColumnProps(286)=   "Column(47).Order=48"
      Splits(0)._ColumnProps(287)=   "Column(48).Width=1270"
      Splits(0)._ColumnProps(288)=   "Column(48).DividerColor=0"
      Splits(0)._ColumnProps(289)=   "Column(48)._WidthInPix=1138"
      Splits(0)._ColumnProps(290)=   "Column(48)._ColStyle=2"
      Splits(0)._ColumnProps(291)=   "Column(48).FetchStyle=1"
      Splits(0)._ColumnProps(292)=   "Column(48).Order=49"
      Splits(0)._ColumnProps(293)=   "Column(49).Width=1270"
      Splits(0)._ColumnProps(294)=   "Column(49).DividerColor=0"
      Splits(0)._ColumnProps(295)=   "Column(49)._WidthInPix=1138"
      Splits(0)._ColumnProps(296)=   "Column(49)._ColStyle=2"
      Splits(0)._ColumnProps(297)=   "Column(49).FetchStyle=1"
      Splits(0)._ColumnProps(298)=   "Column(49).Order=50"
      Splits(0)._ColumnProps(299)=   "Column(50).Width=1270"
      Splits(0)._ColumnProps(300)=   "Column(50).DividerColor=0"
      Splits(0)._ColumnProps(301)=   "Column(50)._WidthInPix=1138"
      Splits(0)._ColumnProps(302)=   "Column(50)._ColStyle=2"
      Splits(0)._ColumnProps(303)=   "Column(50).FetchStyle=1"
      Splits(0)._ColumnProps(304)=   "Column(50).Order=51"
      Splits(0)._ColumnProps(305)=   "Column(51).Width=1270"
      Splits(0)._ColumnProps(306)=   "Column(51).DividerColor=0"
      Splits(0)._ColumnProps(307)=   "Column(51)._WidthInPix=1138"
      Splits(0)._ColumnProps(308)=   "Column(51)._ColStyle=2"
      Splits(0)._ColumnProps(309)=   "Column(51).FetchStyle=1"
      Splits(0)._ColumnProps(310)=   "Column(51).Order=52"
      Splits(0)._ColumnProps(311)=   "Column(52).Width=1270"
      Splits(0)._ColumnProps(312)=   "Column(52).DividerColor=0"
      Splits(0)._ColumnProps(313)=   "Column(52)._WidthInPix=1138"
      Splits(0)._ColumnProps(314)=   "Column(52)._ColStyle=2"
      Splits(0)._ColumnProps(315)=   "Column(52).FetchStyle=1"
      Splits(0)._ColumnProps(316)=   "Column(52).Order=53"
      Splits(0)._ColumnProps(317)=   "Column(53).Width=1270"
      Splits(0)._ColumnProps(318)=   "Column(53).DividerColor=0"
      Splits(0)._ColumnProps(319)=   "Column(53)._WidthInPix=1138"
      Splits(0)._ColumnProps(320)=   "Column(53)._ColStyle=2"
      Splits(0)._ColumnProps(321)=   "Column(53).FetchStyle=1"
      Splits(0)._ColumnProps(322)=   "Column(53).Order=54"
      Splits(0)._ColumnProps(323)=   "Column(54).Width=1270"
      Splits(0)._ColumnProps(324)=   "Column(54).DividerColor=0"
      Splits(0)._ColumnProps(325)=   "Column(54)._WidthInPix=1138"
      Splits(0)._ColumnProps(326)=   "Column(54)._ColStyle=2"
      Splits(0)._ColumnProps(327)=   "Column(54).FetchStyle=1"
      Splits(0)._ColumnProps(328)=   "Column(54).Order=55"
      Splits(0)._ColumnProps(329)=   "Column(55).Width=1270"
      Splits(0)._ColumnProps(330)=   "Column(55).DividerColor=0"
      Splits(0)._ColumnProps(331)=   "Column(55)._WidthInPix=1138"
      Splits(0)._ColumnProps(332)=   "Column(55)._ColStyle=2"
      Splits(0)._ColumnProps(333)=   "Column(55).FetchStyle=1"
      Splits(0)._ColumnProps(334)=   "Column(55).Order=56"
      Splits(0)._ColumnProps(335)=   "Column(56).Width=1270"
      Splits(0)._ColumnProps(336)=   "Column(56).DividerColor=0"
      Splits(0)._ColumnProps(337)=   "Column(56)._WidthInPix=1138"
      Splits(0)._ColumnProps(338)=   "Column(56)._ColStyle=2"
      Splits(0)._ColumnProps(339)=   "Column(56).FetchStyle=1"
      Splits(0)._ColumnProps(340)=   "Column(56).Order=57"
      Splits(0)._ColumnProps(341)=   "Column(57).Width=1270"
      Splits(0)._ColumnProps(342)=   "Column(57).DividerColor=0"
      Splits(0)._ColumnProps(343)=   "Column(57)._WidthInPix=1138"
      Splits(0)._ColumnProps(344)=   "Column(57)._ColStyle=2"
      Splits(0)._ColumnProps(345)=   "Column(57).FetchStyle=1"
      Splits(0)._ColumnProps(346)=   "Column(57).Order=58"
      Splits(0)._ColumnProps(347)=   "Column(58).Width=1270"
      Splits(0)._ColumnProps(348)=   "Column(58).DividerColor=0"
      Splits(0)._ColumnProps(349)=   "Column(58)._WidthInPix=1138"
      Splits(0)._ColumnProps(350)=   "Column(58)._ColStyle=2"
      Splits(0)._ColumnProps(351)=   "Column(58).FetchStyle=1"
      Splits(0)._ColumnProps(352)=   "Column(58).Order=59"
      Splits(0)._ColumnProps(353)=   "Column(59).Width=1270"
      Splits(0)._ColumnProps(354)=   "Column(59).DividerColor=0"
      Splits(0)._ColumnProps(355)=   "Column(59)._WidthInPix=1138"
      Splits(0)._ColumnProps(356)=   "Column(59)._ColStyle=2"
      Splits(0)._ColumnProps(357)=   "Column(59).FetchStyle=1"
      Splits(0)._ColumnProps(358)=   "Column(59).Order=60"
      Splits(0)._ColumnProps(359)=   "Column(60).Width=1270"
      Splits(0)._ColumnProps(360)=   "Column(60).DividerColor=0"
      Splits(0)._ColumnProps(361)=   "Column(60)._WidthInPix=1138"
      Splits(0)._ColumnProps(362)=   "Column(60)._ColStyle=2"
      Splits(0)._ColumnProps(363)=   "Column(60).FetchStyle=1"
      Splits(0)._ColumnProps(364)=   "Column(60).Order=61"
      Splits(0)._ColumnProps(365)=   "Column(61).Width=1270"
      Splits(0)._ColumnProps(366)=   "Column(61).DividerColor=0"
      Splits(0)._ColumnProps(367)=   "Column(61)._WidthInPix=1138"
      Splits(0)._ColumnProps(368)=   "Column(61)._ColStyle=2"
      Splits(0)._ColumnProps(369)=   "Column(61).FetchStyle=1"
      Splits(0)._ColumnProps(370)=   "Column(61).Order=62"
      Splits(0)._ColumnProps(371)=   "Column(62).Width=1270"
      Splits(0)._ColumnProps(372)=   "Column(62).DividerColor=0"
      Splits(0)._ColumnProps(373)=   "Column(62)._WidthInPix=1138"
      Splits(0)._ColumnProps(374)=   "Column(62)._ColStyle=2"
      Splits(0)._ColumnProps(375)=   "Column(62).FetchStyle=1"
      Splits(0)._ColumnProps(376)=   "Column(62).Order=63"
      Splits(0)._ColumnProps(377)=   "Column(63).Width=1270"
      Splits(0)._ColumnProps(378)=   "Column(63).DividerColor=0"
      Splits(0)._ColumnProps(379)=   "Column(63)._WidthInPix=1138"
      Splits(0)._ColumnProps(380)=   "Column(63)._ColStyle=2"
      Splits(0)._ColumnProps(381)=   "Column(63).FetchStyle=1"
      Splits(0)._ColumnProps(382)=   "Column(63).Order=64"
      Splits(0)._ColumnProps(383)=   "Column(64).Width=1270"
      Splits(0)._ColumnProps(384)=   "Column(64).DividerColor=0"
      Splits(0)._ColumnProps(385)=   "Column(64)._WidthInPix=1138"
      Splits(0)._ColumnProps(386)=   "Column(64)._ColStyle=2"
      Splits(0)._ColumnProps(387)=   "Column(64).FetchStyle=1"
      Splits(0)._ColumnProps(388)=   "Column(64).Order=65"
      Splits(0)._ColumnProps(389)=   "Column(65).Width=1270"
      Splits(0)._ColumnProps(390)=   "Column(65).DividerColor=0"
      Splits(0)._ColumnProps(391)=   "Column(65)._WidthInPix=1138"
      Splits(0)._ColumnProps(392)=   "Column(65)._ColStyle=2"
      Splits(0)._ColumnProps(393)=   "Column(65).FetchStyle=1"
      Splits(0)._ColumnProps(394)=   "Column(65).Order=66"
      Splits(0)._ColumnProps(395)=   "Column(66).Width=1270"
      Splits(0)._ColumnProps(396)=   "Column(66).DividerColor=0"
      Splits(0)._ColumnProps(397)=   "Column(66)._WidthInPix=1138"
      Splits(0)._ColumnProps(398)=   "Column(66)._ColStyle=2"
      Splits(0)._ColumnProps(399)=   "Column(66).FetchStyle=1"
      Splits(0)._ColumnProps(400)=   "Column(66).Order=67"
      Splits(0)._ColumnProps(401)=   "Column(67).Width=1270"
      Splits(0)._ColumnProps(402)=   "Column(67).DividerColor=0"
      Splits(0)._ColumnProps(403)=   "Column(67)._WidthInPix=1138"
      Splits(0)._ColumnProps(404)=   "Column(67)._ColStyle=2"
      Splits(0)._ColumnProps(405)=   "Column(67).FetchStyle=1"
      Splits(0)._ColumnProps(406)=   "Column(67).Order=68"
      Splits(0)._ColumnProps(407)=   "Column(68).Width=1270"
      Splits(0)._ColumnProps(408)=   "Column(68).DividerColor=0"
      Splits(0)._ColumnProps(409)=   "Column(68)._WidthInPix=1138"
      Splits(0)._ColumnProps(410)=   "Column(68)._ColStyle=2"
      Splits(0)._ColumnProps(411)=   "Column(68).FetchStyle=1"
      Splits(0)._ColumnProps(412)=   "Column(68).Order=69"
      Splits(0)._ColumnProps(413)=   "Column(69).Width=1270"
      Splits(0)._ColumnProps(414)=   "Column(69).DividerColor=0"
      Splits(0)._ColumnProps(415)=   "Column(69)._WidthInPix=1138"
      Splits(0)._ColumnProps(416)=   "Column(69)._ColStyle=2"
      Splits(0)._ColumnProps(417)=   "Column(69).FetchStyle=1"
      Splits(0)._ColumnProps(418)=   "Column(69).Order=70"
      Splits(0)._ColumnProps(419)=   "Column(70).Width=1270"
      Splits(0)._ColumnProps(420)=   "Column(70).DividerColor=0"
      Splits(0)._ColumnProps(421)=   "Column(70)._WidthInPix=1138"
      Splits(0)._ColumnProps(422)=   "Column(70)._ColStyle=2"
      Splits(0)._ColumnProps(423)=   "Column(70).FetchStyle=1"
      Splits(0)._ColumnProps(424)=   "Column(70).Order=71"
      Splits(0)._ColumnProps(425)=   "Column(71).Width=1270"
      Splits(0)._ColumnProps(426)=   "Column(71).DividerColor=0"
      Splits(0)._ColumnProps(427)=   "Column(71)._WidthInPix=1138"
      Splits(0)._ColumnProps(428)=   "Column(71)._ColStyle=2"
      Splits(0)._ColumnProps(429)=   "Column(71).FetchStyle=1"
      Splits(0)._ColumnProps(430)=   "Column(71).Order=72"
      Splits(0)._ColumnProps(431)=   "Column(72).Width=1270"
      Splits(0)._ColumnProps(432)=   "Column(72).DividerColor=0"
      Splits(0)._ColumnProps(433)=   "Column(72)._WidthInPix=1138"
      Splits(0)._ColumnProps(434)=   "Column(72)._ColStyle=2"
      Splits(0)._ColumnProps(435)=   "Column(72).FetchStyle=1"
      Splits(0)._ColumnProps(436)=   "Column(72).Order=73"
      Splits(0)._ColumnProps(437)=   "Column(73).Width=1270"
      Splits(0)._ColumnProps(438)=   "Column(73).DividerColor=0"
      Splits(0)._ColumnProps(439)=   "Column(73)._WidthInPix=1138"
      Splits(0)._ColumnProps(440)=   "Column(73)._ColStyle=2"
      Splits(0)._ColumnProps(441)=   "Column(73).FetchStyle=1"
      Splits(0)._ColumnProps(442)=   "Column(73).Order=74"
      Splits(0)._ColumnProps(443)=   "Column(74).Width=1270"
      Splits(0)._ColumnProps(444)=   "Column(74).DividerColor=0"
      Splits(0)._ColumnProps(445)=   "Column(74)._WidthInPix=1138"
      Splits(0)._ColumnProps(446)=   "Column(74)._ColStyle=2"
      Splits(0)._ColumnProps(447)=   "Column(74).FetchStyle=1"
      Splits(0)._ColumnProps(448)=   "Column(74).Order=75"
      Splits(0)._ColumnProps(449)=   "Column(75).Width=1270"
      Splits(0)._ColumnProps(450)=   "Column(75).DividerColor=0"
      Splits(0)._ColumnProps(451)=   "Column(75)._WidthInPix=1138"
      Splits(0)._ColumnProps(452)=   "Column(75)._ColStyle=2"
      Splits(0)._ColumnProps(453)=   "Column(75).FetchStyle=1"
      Splits(0)._ColumnProps(454)=   "Column(75).Order=76"
      Splits(0)._ColumnProps(455)=   "Column(76).Width=1270"
      Splits(0)._ColumnProps(456)=   "Column(76).DividerColor=0"
      Splits(0)._ColumnProps(457)=   "Column(76)._WidthInPix=1138"
      Splits(0)._ColumnProps(458)=   "Column(76)._ColStyle=2"
      Splits(0)._ColumnProps(459)=   "Column(76).FetchStyle=1"
      Splits(0)._ColumnProps(460)=   "Column(76).Order=77"
      Splits(0)._ColumnProps(461)=   "Column(77).Width=1270"
      Splits(0)._ColumnProps(462)=   "Column(77).DividerColor=0"
      Splits(0)._ColumnProps(463)=   "Column(77)._WidthInPix=1138"
      Splits(0)._ColumnProps(464)=   "Column(77)._ColStyle=2"
      Splits(0)._ColumnProps(465)=   "Column(77).FetchStyle=1"
      Splits(0)._ColumnProps(466)=   "Column(77).Order=78"
      Splits(0)._ColumnProps(467)=   "Column(78).Width=1270"
      Splits(0)._ColumnProps(468)=   "Column(78).DividerColor=0"
      Splits(0)._ColumnProps(469)=   "Column(78)._WidthInPix=1138"
      Splits(0)._ColumnProps(470)=   "Column(78)._ColStyle=2"
      Splits(0)._ColumnProps(471)=   "Column(78).FetchStyle=1"
      Splits(0)._ColumnProps(472)=   "Column(78).Order=79"
      Splits(0)._ColumnProps(473)=   "Column(79).Width=1270"
      Splits(0)._ColumnProps(474)=   "Column(79).DividerColor=0"
      Splits(0)._ColumnProps(475)=   "Column(79)._WidthInPix=1138"
      Splits(0)._ColumnProps(476)=   "Column(79)._ColStyle=2"
      Splits(0)._ColumnProps(477)=   "Column(79).FetchStyle=1"
      Splits(0)._ColumnProps(478)=   "Column(79).Order=80"
      Splits(0)._ColumnProps(479)=   "Column(80).Width=1270"
      Splits(0)._ColumnProps(480)=   "Column(80).DividerColor=0"
      Splits(0)._ColumnProps(481)=   "Column(80)._WidthInPix=1138"
      Splits(0)._ColumnProps(482)=   "Column(80)._ColStyle=2"
      Splits(0)._ColumnProps(483)=   "Column(80).FetchStyle=1"
      Splits(0)._ColumnProps(484)=   "Column(80).Order=81"
      Splits(0)._ColumnProps(485)=   "Column(81).Width=1270"
      Splits(0)._ColumnProps(486)=   "Column(81).DividerColor=0"
      Splits(0)._ColumnProps(487)=   "Column(81)._WidthInPix=1138"
      Splits(0)._ColumnProps(488)=   "Column(81)._ColStyle=2"
      Splits(0)._ColumnProps(489)=   "Column(81).Order=82"
      Splits(0)._ColumnProps(490)=   "Column(82).Width=1270"
      Splits(0)._ColumnProps(491)=   "Column(82).DividerColor=0"
      Splits(0)._ColumnProps(492)=   "Column(82)._WidthInPix=1138"
      Splits(0)._ColumnProps(493)=   "Column(82)._ColStyle=2"
      Splits(0)._ColumnProps(494)=   "Column(82).FetchStyle=1"
      Splits(0)._ColumnProps(495)=   "Column(82).Order=83"
      Splits(0)._ColumnProps(496)=   "Column(83).Width=1270"
      Splits(0)._ColumnProps(497)=   "Column(83).DividerColor=0"
      Splits(0)._ColumnProps(498)=   "Column(83)._WidthInPix=1138"
      Splits(0)._ColumnProps(499)=   "Column(83)._ColStyle=2"
      Splits(0)._ColumnProps(500)=   "Column(83).FetchStyle=1"
      Splits(0)._ColumnProps(501)=   "Column(83).Order=84"
      Splits(0)._ColumnProps(502)=   "Column(84).Width=1270"
      Splits(0)._ColumnProps(503)=   "Column(84).DividerColor=0"
      Splits(0)._ColumnProps(504)=   "Column(84)._WidthInPix=1138"
      Splits(0)._ColumnProps(505)=   "Column(84)._ColStyle=2"
      Splits(0)._ColumnProps(506)=   "Column(84).FetchStyle=1"
      Splits(0)._ColumnProps(507)=   "Column(84).Order=85"
      Splits(0)._ColumnProps(508)=   "Column(85).Width=1270"
      Splits(0)._ColumnProps(509)=   "Column(85).DividerColor=0"
      Splits(0)._ColumnProps(510)=   "Column(85)._WidthInPix=1138"
      Splits(0)._ColumnProps(511)=   "Column(85)._ColStyle=2"
      Splits(0)._ColumnProps(512)=   "Column(85).FetchStyle=1"
      Splits(0)._ColumnProps(513)=   "Column(85).Order=86"
      Splits(0)._ColumnProps(514)=   "Column(86).Width=1270"
      Splits(0)._ColumnProps(515)=   "Column(86).DividerColor=0"
      Splits(0)._ColumnProps(516)=   "Column(86)._WidthInPix=1138"
      Splits(0)._ColumnProps(517)=   "Column(86)._ColStyle=2"
      Splits(0)._ColumnProps(518)=   "Column(86).FetchStyle=1"
      Splits(0)._ColumnProps(519)=   "Column(86).Order=87"
      Splits(0)._ColumnProps(520)=   "Column(87).Width=1270"
      Splits(0)._ColumnProps(521)=   "Column(87).DividerColor=0"
      Splits(0)._ColumnProps(522)=   "Column(87)._WidthInPix=1138"
      Splits(0)._ColumnProps(523)=   "Column(87)._ColStyle=2"
      Splits(0)._ColumnProps(524)=   "Column(87).FetchStyle=1"
      Splits(0)._ColumnProps(525)=   "Column(87).Order=88"
      Splits(0)._ColumnProps(526)=   "Column(88).Width=1270"
      Splits(0)._ColumnProps(527)=   "Column(88).DividerColor=0"
      Splits(0)._ColumnProps(528)=   "Column(88)._WidthInPix=1138"
      Splits(0)._ColumnProps(529)=   "Column(88)._ColStyle=2"
      Splits(0)._ColumnProps(530)=   "Column(88).Order=89"
      Splits(0)._ColumnProps(531)=   "Column(89).Width=1270"
      Splits(0)._ColumnProps(532)=   "Column(89).DividerColor=0"
      Splits(0)._ColumnProps(533)=   "Column(89)._WidthInPix=1138"
      Splits(0)._ColumnProps(534)=   "Column(89)._ColStyle=2"
      Splits(0)._ColumnProps(535)=   "Column(89).FetchStyle=1"
      Splits(0)._ColumnProps(536)=   "Column(89).Order=90"
      Splits(0)._ColumnProps(537)=   "Column(90).Width=1270"
      Splits(0)._ColumnProps(538)=   "Column(90).DividerColor=0"
      Splits(0)._ColumnProps(539)=   "Column(90)._WidthInPix=1138"
      Splits(0)._ColumnProps(540)=   "Column(90)._ColStyle=2"
      Splits(0)._ColumnProps(541)=   "Column(90).FetchStyle=1"
      Splits(0)._ColumnProps(542)=   "Column(90).Order=91"
      Splits(0)._ColumnProps(543)=   "Column(91).Width=1270"
      Splits(0)._ColumnProps(544)=   "Column(91).DividerColor=0"
      Splits(0)._ColumnProps(545)=   "Column(91)._WidthInPix=1138"
      Splits(0)._ColumnProps(546)=   "Column(91)._ColStyle=2"
      Splits(0)._ColumnProps(547)=   "Column(91).FetchStyle=1"
      Splits(0)._ColumnProps(548)=   "Column(91).Order=92"
      Splits(0)._ColumnProps(549)=   "Column(92).Width=1270"
      Splits(0)._ColumnProps(550)=   "Column(92).DividerColor=0"
      Splits(0)._ColumnProps(551)=   "Column(92)._WidthInPix=1138"
      Splits(0)._ColumnProps(552)=   "Column(92)._ColStyle=2"
      Splits(0)._ColumnProps(553)=   "Column(92).FetchStyle=1"
      Splits(0)._ColumnProps(554)=   "Column(92).Order=93"
      Splits(0)._ColumnProps(555)=   "Column(93).Width=1270"
      Splits(0)._ColumnProps(556)=   "Column(93).DividerColor=0"
      Splits(0)._ColumnProps(557)=   "Column(93)._WidthInPix=1138"
      Splits(0)._ColumnProps(558)=   "Column(93)._ColStyle=2"
      Splits(0)._ColumnProps(559)=   "Column(93).FetchStyle=1"
      Splits(0)._ColumnProps(560)=   "Column(93).Order=94"
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
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "子部品　在庫推移"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=87,.parent=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=96,.parent=4,.bgcolor=&H80FF00&"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=89,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=91,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=94,.parent=9,.bgcolor=&HFFFFFF&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=95,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=411,.parent=87"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=408,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=409,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=410,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=102,.parent=87"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=99,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=100,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=101,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=110,.parent=87"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=107,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=108,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=109,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=114,.parent=87,.alignment=2,.wraptext=-1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=111,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=112,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=113,.parent=91"
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
      _StyleDefs(177) =   "Splits(0).Columns(35).Style:id=118,.parent=87,.alignment=1"
      _StyleDefs(178) =   "Splits(0).Columns(35).HeadingStyle:id=115,.parent=88"
      _StyleDefs(179) =   "Splits(0).Columns(35).FooterStyle:id=116,.parent=89"
      _StyleDefs(180) =   "Splits(0).Columns(35).EditorStyle:id=117,.parent=91"
      _StyleDefs(181) =   "Splits(0).Columns(36).Style:id=138,.parent=87,.alignment=1"
      _StyleDefs(182) =   "Splits(0).Columns(36).HeadingStyle:id=135,.parent=88"
      _StyleDefs(183) =   "Splits(0).Columns(36).FooterStyle:id=136,.parent=89"
      _StyleDefs(184) =   "Splits(0).Columns(36).EditorStyle:id=137,.parent=91"
      _StyleDefs(185) =   "Splits(0).Columns(37).Style:id=183,.parent=87,.alignment=1"
      _StyleDefs(186) =   "Splits(0).Columns(37).HeadingStyle:id=180,.parent=88"
      _StyleDefs(187) =   "Splits(0).Columns(37).FooterStyle:id=181,.parent=89"
      _StyleDefs(188) =   "Splits(0).Columns(37).EditorStyle:id=182,.parent=91"
      _StyleDefs(189) =   "Splits(0).Columns(38).Style:id=187,.parent=87,.alignment=1"
      _StyleDefs(190) =   "Splits(0).Columns(38).HeadingStyle:id=184,.parent=88"
      _StyleDefs(191) =   "Splits(0).Columns(38).FooterStyle:id=185,.parent=89"
      _StyleDefs(192) =   "Splits(0).Columns(38).EditorStyle:id=186,.parent=91"
      _StyleDefs(193) =   "Splits(0).Columns(39).Style:id=191,.parent=87,.alignment=1"
      _StyleDefs(194) =   "Splits(0).Columns(39).HeadingStyle:id=188,.parent=88"
      _StyleDefs(195) =   "Splits(0).Columns(39).FooterStyle:id=189,.parent=89"
      _StyleDefs(196) =   "Splits(0).Columns(39).EditorStyle:id=190,.parent=91"
      _StyleDefs(197) =   "Splits(0).Columns(40).Style:id=195,.parent=87,.alignment=1"
      _StyleDefs(198) =   "Splits(0).Columns(40).HeadingStyle:id=192,.parent=88"
      _StyleDefs(199) =   "Splits(0).Columns(40).FooterStyle:id=193,.parent=89"
      _StyleDefs(200) =   "Splits(0).Columns(40).EditorStyle:id=194,.parent=91"
      _StyleDefs(201) =   "Splits(0).Columns(41).Style:id=199,.parent=87,.alignment=1"
      _StyleDefs(202) =   "Splits(0).Columns(41).HeadingStyle:id=196,.parent=88"
      _StyleDefs(203) =   "Splits(0).Columns(41).FooterStyle:id=197,.parent=89"
      _StyleDefs(204) =   "Splits(0).Columns(41).EditorStyle:id=198,.parent=91"
      _StyleDefs(205) =   "Splits(0).Columns(42).Style:id=203,.parent=87,.alignment=1"
      _StyleDefs(206) =   "Splits(0).Columns(42).HeadingStyle:id=200,.parent=88"
      _StyleDefs(207) =   "Splits(0).Columns(42).FooterStyle:id=201,.parent=89"
      _StyleDefs(208) =   "Splits(0).Columns(42).EditorStyle:id=202,.parent=91"
      _StyleDefs(209) =   "Splits(0).Columns(43).Style:id=207,.parent=87,.alignment=1"
      _StyleDefs(210) =   "Splits(0).Columns(43).HeadingStyle:id=204,.parent=88"
      _StyleDefs(211) =   "Splits(0).Columns(43).FooterStyle:id=205,.parent=89"
      _StyleDefs(212) =   "Splits(0).Columns(43).EditorStyle:id=206,.parent=91"
      _StyleDefs(213) =   "Splits(0).Columns(44).Style:id=211,.parent=87,.alignment=1"
      _StyleDefs(214) =   "Splits(0).Columns(44).HeadingStyle:id=208,.parent=88"
      _StyleDefs(215) =   "Splits(0).Columns(44).FooterStyle:id=209,.parent=89"
      _StyleDefs(216) =   "Splits(0).Columns(44).EditorStyle:id=210,.parent=91"
      _StyleDefs(217) =   "Splits(0).Columns(45).Style:id=215,.parent=87,.alignment=1"
      _StyleDefs(218) =   "Splits(0).Columns(45).HeadingStyle:id=212,.parent=88"
      _StyleDefs(219) =   "Splits(0).Columns(45).FooterStyle:id=213,.parent=89"
      _StyleDefs(220) =   "Splits(0).Columns(45).EditorStyle:id=214,.parent=91"
      _StyleDefs(221) =   "Splits(0).Columns(46).Style:id=219,.parent=87,.alignment=1"
      _StyleDefs(222) =   "Splits(0).Columns(46).HeadingStyle:id=216,.parent=88"
      _StyleDefs(223) =   "Splits(0).Columns(46).FooterStyle:id=217,.parent=89"
      _StyleDefs(224) =   "Splits(0).Columns(46).EditorStyle:id=218,.parent=91"
      _StyleDefs(225) =   "Splits(0).Columns(47).Style:id=223,.parent=87,.alignment=1"
      _StyleDefs(226) =   "Splits(0).Columns(47).HeadingStyle:id=220,.parent=88"
      _StyleDefs(227) =   "Splits(0).Columns(47).FooterStyle:id=221,.parent=89"
      _StyleDefs(228) =   "Splits(0).Columns(47).EditorStyle:id=222,.parent=91"
      _StyleDefs(229) =   "Splits(0).Columns(48).Style:id=227,.parent=87,.alignment=1"
      _StyleDefs(230) =   "Splits(0).Columns(48).HeadingStyle:id=224,.parent=88"
      _StyleDefs(231) =   "Splits(0).Columns(48).FooterStyle:id=225,.parent=89"
      _StyleDefs(232) =   "Splits(0).Columns(48).EditorStyle:id=226,.parent=91"
      _StyleDefs(233) =   "Splits(0).Columns(49).Style:id=231,.parent=87,.alignment=1"
      _StyleDefs(234) =   "Splits(0).Columns(49).HeadingStyle:id=228,.parent=88"
      _StyleDefs(235) =   "Splits(0).Columns(49).FooterStyle:id=229,.parent=89"
      _StyleDefs(236) =   "Splits(0).Columns(49).EditorStyle:id=230,.parent=91"
      _StyleDefs(237) =   "Splits(0).Columns(50).Style:id=235,.parent=87,.alignment=1"
      _StyleDefs(238) =   "Splits(0).Columns(50).HeadingStyle:id=232,.parent=88"
      _StyleDefs(239) =   "Splits(0).Columns(50).FooterStyle:id=233,.parent=89"
      _StyleDefs(240) =   "Splits(0).Columns(50).EditorStyle:id=234,.parent=91"
      _StyleDefs(241) =   "Splits(0).Columns(51).Style:id=239,.parent=87,.alignment=1"
      _StyleDefs(242) =   "Splits(0).Columns(51).HeadingStyle:id=236,.parent=88"
      _StyleDefs(243) =   "Splits(0).Columns(51).FooterStyle:id=237,.parent=89"
      _StyleDefs(244) =   "Splits(0).Columns(51).EditorStyle:id=238,.parent=91"
      _StyleDefs(245) =   "Splits(0).Columns(52).Style:id=243,.parent=87,.alignment=1"
      _StyleDefs(246) =   "Splits(0).Columns(52).HeadingStyle:id=240,.parent=88"
      _StyleDefs(247) =   "Splits(0).Columns(52).FooterStyle:id=241,.parent=89"
      _StyleDefs(248) =   "Splits(0).Columns(52).EditorStyle:id=242,.parent=91"
      _StyleDefs(249) =   "Splits(0).Columns(53).Style:id=247,.parent=87,.alignment=1"
      _StyleDefs(250) =   "Splits(0).Columns(53).HeadingStyle:id=244,.parent=88"
      _StyleDefs(251) =   "Splits(0).Columns(53).FooterStyle:id=245,.parent=89"
      _StyleDefs(252) =   "Splits(0).Columns(53).EditorStyle:id=246,.parent=91"
      _StyleDefs(253) =   "Splits(0).Columns(54).Style:id=251,.parent=87,.alignment=1"
      _StyleDefs(254) =   "Splits(0).Columns(54).HeadingStyle:id=248,.parent=88"
      _StyleDefs(255) =   "Splits(0).Columns(54).FooterStyle:id=249,.parent=89"
      _StyleDefs(256) =   "Splits(0).Columns(54).EditorStyle:id=250,.parent=91"
      _StyleDefs(257) =   "Splits(0).Columns(55).Style:id=255,.parent=87,.alignment=1"
      _StyleDefs(258) =   "Splits(0).Columns(55).HeadingStyle:id=252,.parent=88"
      _StyleDefs(259) =   "Splits(0).Columns(55).FooterStyle:id=253,.parent=89"
      _StyleDefs(260) =   "Splits(0).Columns(55).EditorStyle:id=254,.parent=91"
      _StyleDefs(261) =   "Splits(0).Columns(56).Style:id=259,.parent=87,.alignment=1"
      _StyleDefs(262) =   "Splits(0).Columns(56).HeadingStyle:id=256,.parent=88"
      _StyleDefs(263) =   "Splits(0).Columns(56).FooterStyle:id=257,.parent=89"
      _StyleDefs(264) =   "Splits(0).Columns(56).EditorStyle:id=258,.parent=91"
      _StyleDefs(265) =   "Splits(0).Columns(57).Style:id=263,.parent=87,.alignment=1"
      _StyleDefs(266) =   "Splits(0).Columns(57).HeadingStyle:id=260,.parent=88"
      _StyleDefs(267) =   "Splits(0).Columns(57).FooterStyle:id=261,.parent=89"
      _StyleDefs(268) =   "Splits(0).Columns(57).EditorStyle:id=262,.parent=91"
      _StyleDefs(269) =   "Splits(0).Columns(58).Style:id=271,.parent=87,.alignment=1"
      _StyleDefs(270) =   "Splits(0).Columns(58).HeadingStyle:id=268,.parent=88"
      _StyleDefs(271) =   "Splits(0).Columns(58).FooterStyle:id=269,.parent=89"
      _StyleDefs(272) =   "Splits(0).Columns(58).EditorStyle:id=270,.parent=91"
      _StyleDefs(273) =   "Splits(0).Columns(59).Style:id=275,.parent=87,.alignment=1"
      _StyleDefs(274) =   "Splits(0).Columns(59).HeadingStyle:id=272,.parent=88"
      _StyleDefs(275) =   "Splits(0).Columns(59).FooterStyle:id=273,.parent=89"
      _StyleDefs(276) =   "Splits(0).Columns(59).EditorStyle:id=274,.parent=91"
      _StyleDefs(277) =   "Splits(0).Columns(60).Style:id=279,.parent=87,.alignment=1"
      _StyleDefs(278) =   "Splits(0).Columns(60).HeadingStyle:id=276,.parent=88"
      _StyleDefs(279) =   "Splits(0).Columns(60).FooterStyle:id=277,.parent=89"
      _StyleDefs(280) =   "Splits(0).Columns(60).EditorStyle:id=278,.parent=91"
      _StyleDefs(281) =   "Splits(0).Columns(61).Style:id=283,.parent=87,.alignment=1"
      _StyleDefs(282) =   "Splits(0).Columns(61).HeadingStyle:id=280,.parent=88"
      _StyleDefs(283) =   "Splits(0).Columns(61).FooterStyle:id=281,.parent=89"
      _StyleDefs(284) =   "Splits(0).Columns(61).EditorStyle:id=282,.parent=91"
      _StyleDefs(285) =   "Splits(0).Columns(62).Style:id=287,.parent=87,.alignment=1"
      _StyleDefs(286) =   "Splits(0).Columns(62).HeadingStyle:id=284,.parent=88"
      _StyleDefs(287) =   "Splits(0).Columns(62).FooterStyle:id=285,.parent=89"
      _StyleDefs(288) =   "Splits(0).Columns(62).EditorStyle:id=286,.parent=91"
      _StyleDefs(289) =   "Splits(0).Columns(63).Style:id=291,.parent=87,.alignment=1"
      _StyleDefs(290) =   "Splits(0).Columns(63).HeadingStyle:id=288,.parent=88"
      _StyleDefs(291) =   "Splits(0).Columns(63).FooterStyle:id=289,.parent=89"
      _StyleDefs(292) =   "Splits(0).Columns(63).EditorStyle:id=290,.parent=91"
      _StyleDefs(293) =   "Splits(0).Columns(64).Style:id=295,.parent=87,.alignment=1"
      _StyleDefs(294) =   "Splits(0).Columns(64).HeadingStyle:id=292,.parent=88"
      _StyleDefs(295) =   "Splits(0).Columns(64).FooterStyle:id=293,.parent=89"
      _StyleDefs(296) =   "Splits(0).Columns(64).EditorStyle:id=294,.parent=91"
      _StyleDefs(297) =   "Splits(0).Columns(65).Style:id=299,.parent=87,.alignment=1"
      _StyleDefs(298) =   "Splits(0).Columns(65).HeadingStyle:id=296,.parent=88"
      _StyleDefs(299) =   "Splits(0).Columns(65).FooterStyle:id=297,.parent=89"
      _StyleDefs(300) =   "Splits(0).Columns(65).EditorStyle:id=298,.parent=91"
      _StyleDefs(301) =   "Splits(0).Columns(66).Style:id=303,.parent=87,.alignment=1"
      _StyleDefs(302) =   "Splits(0).Columns(66).HeadingStyle:id=300,.parent=88"
      _StyleDefs(303) =   "Splits(0).Columns(66).FooterStyle:id=301,.parent=89"
      _StyleDefs(304) =   "Splits(0).Columns(66).EditorStyle:id=302,.parent=91"
      _StyleDefs(305) =   "Splits(0).Columns(67).Style:id=307,.parent=87,.alignment=1"
      _StyleDefs(306) =   "Splits(0).Columns(67).HeadingStyle:id=304,.parent=88"
      _StyleDefs(307) =   "Splits(0).Columns(67).FooterStyle:id=305,.parent=89"
      _StyleDefs(308) =   "Splits(0).Columns(67).EditorStyle:id=306,.parent=91"
      _StyleDefs(309) =   "Splits(0).Columns(68).Style:id=311,.parent=87,.alignment=1"
      _StyleDefs(310) =   "Splits(0).Columns(68).HeadingStyle:id=308,.parent=88"
      _StyleDefs(311) =   "Splits(0).Columns(68).FooterStyle:id=309,.parent=89"
      _StyleDefs(312) =   "Splits(0).Columns(68).EditorStyle:id=310,.parent=91"
      _StyleDefs(313) =   "Splits(0).Columns(69).Style:id=315,.parent=87,.alignment=1"
      _StyleDefs(314) =   "Splits(0).Columns(69).HeadingStyle:id=312,.parent=88"
      _StyleDefs(315) =   "Splits(0).Columns(69).FooterStyle:id=313,.parent=89"
      _StyleDefs(316) =   "Splits(0).Columns(69).EditorStyle:id=314,.parent=91"
      _StyleDefs(317) =   "Splits(0).Columns(70).Style:id=319,.parent=87,.alignment=1"
      _StyleDefs(318) =   "Splits(0).Columns(70).HeadingStyle:id=316,.parent=88"
      _StyleDefs(319) =   "Splits(0).Columns(70).FooterStyle:id=317,.parent=89"
      _StyleDefs(320) =   "Splits(0).Columns(70).EditorStyle:id=318,.parent=91"
      _StyleDefs(321) =   "Splits(0).Columns(71).Style:id=323,.parent=87,.alignment=1"
      _StyleDefs(322) =   "Splits(0).Columns(71).HeadingStyle:id=320,.parent=88"
      _StyleDefs(323) =   "Splits(0).Columns(71).FooterStyle:id=321,.parent=89"
      _StyleDefs(324) =   "Splits(0).Columns(71).EditorStyle:id=322,.parent=91"
      _StyleDefs(325) =   "Splits(0).Columns(72).Style:id=327,.parent=87,.alignment=1"
      _StyleDefs(326) =   "Splits(0).Columns(72).HeadingStyle:id=324,.parent=88"
      _StyleDefs(327) =   "Splits(0).Columns(72).FooterStyle:id=325,.parent=89"
      _StyleDefs(328) =   "Splits(0).Columns(72).EditorStyle:id=326,.parent=91"
      _StyleDefs(329) =   "Splits(0).Columns(73).Style:id=331,.parent=87,.alignment=1"
      _StyleDefs(330) =   "Splits(0).Columns(73).HeadingStyle:id=328,.parent=88"
      _StyleDefs(331) =   "Splits(0).Columns(73).FooterStyle:id=329,.parent=89"
      _StyleDefs(332) =   "Splits(0).Columns(73).EditorStyle:id=330,.parent=91"
      _StyleDefs(333) =   "Splits(0).Columns(74).Style:id=335,.parent=87,.alignment=1"
      _StyleDefs(334) =   "Splits(0).Columns(74).HeadingStyle:id=332,.parent=88"
      _StyleDefs(335) =   "Splits(0).Columns(74).FooterStyle:id=333,.parent=89"
      _StyleDefs(336) =   "Splits(0).Columns(74).EditorStyle:id=334,.parent=91"
      _StyleDefs(337) =   "Splits(0).Columns(75).Style:id=339,.parent=87,.alignment=1"
      _StyleDefs(338) =   "Splits(0).Columns(75).HeadingStyle:id=336,.parent=88"
      _StyleDefs(339) =   "Splits(0).Columns(75).FooterStyle:id=337,.parent=89"
      _StyleDefs(340) =   "Splits(0).Columns(75).EditorStyle:id=338,.parent=91"
      _StyleDefs(341) =   "Splits(0).Columns(76).Style:id=343,.parent=87,.alignment=1"
      _StyleDefs(342) =   "Splits(0).Columns(76).HeadingStyle:id=340,.parent=88"
      _StyleDefs(343) =   "Splits(0).Columns(76).FooterStyle:id=341,.parent=89"
      _StyleDefs(344) =   "Splits(0).Columns(76).EditorStyle:id=342,.parent=91"
      _StyleDefs(345) =   "Splits(0).Columns(77).Style:id=347,.parent=87,.alignment=1"
      _StyleDefs(346) =   "Splits(0).Columns(77).HeadingStyle:id=344,.parent=88"
      _StyleDefs(347) =   "Splits(0).Columns(77).FooterStyle:id=345,.parent=89"
      _StyleDefs(348) =   "Splits(0).Columns(77).EditorStyle:id=346,.parent=91"
      _StyleDefs(349) =   "Splits(0).Columns(78).Style:id=351,.parent=87,.alignment=1"
      _StyleDefs(350) =   "Splits(0).Columns(78).HeadingStyle:id=348,.parent=88"
      _StyleDefs(351) =   "Splits(0).Columns(78).FooterStyle:id=349,.parent=89"
      _StyleDefs(352) =   "Splits(0).Columns(78).EditorStyle:id=350,.parent=91"
      _StyleDefs(353) =   "Splits(0).Columns(79).Style:id=355,.parent=87,.alignment=1"
      _StyleDefs(354) =   "Splits(0).Columns(79).HeadingStyle:id=352,.parent=88"
      _StyleDefs(355) =   "Splits(0).Columns(79).FooterStyle:id=353,.parent=89"
      _StyleDefs(356) =   "Splits(0).Columns(79).EditorStyle:id=354,.parent=91"
      _StyleDefs(357) =   "Splits(0).Columns(80).Style:id=359,.parent=87,.alignment=1"
      _StyleDefs(358) =   "Splits(0).Columns(80).HeadingStyle:id=356,.parent=88"
      _StyleDefs(359) =   "Splits(0).Columns(80).FooterStyle:id=357,.parent=89"
      _StyleDefs(360) =   "Splits(0).Columns(80).EditorStyle:id=358,.parent=91"
      _StyleDefs(361) =   "Splits(0).Columns(81).Style:id=363,.parent=87,.alignment=1"
      _StyleDefs(362) =   "Splits(0).Columns(81).HeadingStyle:id=360,.parent=88"
      _StyleDefs(363) =   "Splits(0).Columns(81).FooterStyle:id=361,.parent=89"
      _StyleDefs(364) =   "Splits(0).Columns(81).EditorStyle:id=362,.parent=91"
      _StyleDefs(365) =   "Splits(0).Columns(82).Style:id=367,.parent=87,.alignment=1"
      _StyleDefs(366) =   "Splits(0).Columns(82).HeadingStyle:id=364,.parent=88"
      _StyleDefs(367) =   "Splits(0).Columns(82).FooterStyle:id=365,.parent=89"
      _StyleDefs(368) =   "Splits(0).Columns(82).EditorStyle:id=366,.parent=91"
      _StyleDefs(369) =   "Splits(0).Columns(83).Style:id=371,.parent=87,.alignment=1"
      _StyleDefs(370) =   "Splits(0).Columns(83).HeadingStyle:id=368,.parent=88"
      _StyleDefs(371) =   "Splits(0).Columns(83).FooterStyle:id=369,.parent=89"
      _StyleDefs(372) =   "Splits(0).Columns(83).EditorStyle:id=370,.parent=91"
      _StyleDefs(373) =   "Splits(0).Columns(84).Style:id=375,.parent=87,.alignment=1"
      _StyleDefs(374) =   "Splits(0).Columns(84).HeadingStyle:id=372,.parent=88"
      _StyleDefs(375) =   "Splits(0).Columns(84).FooterStyle:id=373,.parent=89"
      _StyleDefs(376) =   "Splits(0).Columns(84).EditorStyle:id=374,.parent=91"
      _StyleDefs(377) =   "Splits(0).Columns(85).Style:id=379,.parent=87,.alignment=1"
      _StyleDefs(378) =   "Splits(0).Columns(85).HeadingStyle:id=376,.parent=88"
      _StyleDefs(379) =   "Splits(0).Columns(85).FooterStyle:id=377,.parent=89"
      _StyleDefs(380) =   "Splits(0).Columns(85).EditorStyle:id=378,.parent=91"
      _StyleDefs(381) =   "Splits(0).Columns(86).Style:id=383,.parent=87,.alignment=1"
      _StyleDefs(382) =   "Splits(0).Columns(86).HeadingStyle:id=380,.parent=88"
      _StyleDefs(383) =   "Splits(0).Columns(86).FooterStyle:id=381,.parent=89"
      _StyleDefs(384) =   "Splits(0).Columns(86).EditorStyle:id=382,.parent=91"
      _StyleDefs(385) =   "Splits(0).Columns(87).Style:id=387,.parent=87,.alignment=1"
      _StyleDefs(386) =   "Splits(0).Columns(87).HeadingStyle:id=384,.parent=88"
      _StyleDefs(387) =   "Splits(0).Columns(87).FooterStyle:id=385,.parent=89"
      _StyleDefs(388) =   "Splits(0).Columns(87).EditorStyle:id=386,.parent=91"
      _StyleDefs(389) =   "Splits(0).Columns(88).Style:id=391,.parent=87,.alignment=1"
      _StyleDefs(390) =   "Splits(0).Columns(88).HeadingStyle:id=388,.parent=88"
      _StyleDefs(391) =   "Splits(0).Columns(88).FooterStyle:id=389,.parent=89"
      _StyleDefs(392) =   "Splits(0).Columns(88).EditorStyle:id=390,.parent=91"
      _StyleDefs(393) =   "Splits(0).Columns(89).Style:id=395,.parent=87,.alignment=1"
      _StyleDefs(394) =   "Splits(0).Columns(89).HeadingStyle:id=392,.parent=88"
      _StyleDefs(395) =   "Splits(0).Columns(89).FooterStyle:id=393,.parent=89"
      _StyleDefs(396) =   "Splits(0).Columns(89).EditorStyle:id=394,.parent=91"
      _StyleDefs(397) =   "Splits(0).Columns(90).Style:id=399,.parent=87,.alignment=1"
      _StyleDefs(398) =   "Splits(0).Columns(90).HeadingStyle:id=396,.parent=88"
      _StyleDefs(399) =   "Splits(0).Columns(90).FooterStyle:id=397,.parent=89"
      _StyleDefs(400) =   "Splits(0).Columns(90).EditorStyle:id=398,.parent=91"
      _StyleDefs(401) =   "Splits(0).Columns(91).Style:id=403,.parent=87,.alignment=1"
      _StyleDefs(402) =   "Splits(0).Columns(91).HeadingStyle:id=400,.parent=88"
      _StyleDefs(403) =   "Splits(0).Columns(91).FooterStyle:id=401,.parent=89"
      _StyleDefs(404) =   "Splits(0).Columns(91).EditorStyle:id=402,.parent=91"
      _StyleDefs(405) =   "Splits(0).Columns(92).Style:id=407,.parent=87,.alignment=1"
      _StyleDefs(406) =   "Splits(0).Columns(92).HeadingStyle:id=404,.parent=88"
      _StyleDefs(407) =   "Splits(0).Columns(92).FooterStyle:id=405,.parent=89"
      _StyleDefs(408) =   "Splits(0).Columns(92).EditorStyle:id=406,.parent=91"
      _StyleDefs(409) =   "Splits(0).Columns(93).Style:id=267,.parent=87,.alignment=1"
      _StyleDefs(410) =   "Splits(0).Columns(93).HeadingStyle:id=264,.parent=88"
      _StyleDefs(411) =   "Splits(0).Columns(93).FooterStyle:id=265,.parent=89"
      _StyleDefs(412) =   "Splits(0).Columns(93).EditorStyle:id=266,.parent=91"
      _StyleDefs(413) =   "Named:id=33:Normal"
      _StyleDefs(414) =   ":id=33,.parent=0"
      _StyleDefs(415) =   "Named:id=34:Heading"
      _StyleDefs(416) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(417) =   ":id=34,.wraptext=-1"
      _StyleDefs(418) =   "Named:id=35:Footing"
      _StyleDefs(419) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(420) =   "Named:id=36:Selected"
      _StyleDefs(421) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(422) =   "Named:id=37:Caption"
      _StyleDefs(423) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(424) =   "Named:id=38:HighlightRow"
      _StyleDefs(425) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(426) =   "Named:id=39:EvenRow"
      _StyleDefs(427) =   ":id=39,.parent=33"
      _StyleDefs(428) =   "Named:id=40:OddRow"
      _StyleDefs(429) =   ":id=40,.parent=33,.bgcolor=&H80FF00&"
      _StyleDefs(430) =   "Named:id=41:RecordSelector"
      _StyleDefs(431) =   ":id=41,.parent=34"
      _StyleDefs(432) =   "Named:id=42:FilterBar"
      _StyleDefs(433) =   ":id=42,.parent=33"
      _StyleDefs(434) =   "Named:id=13:LockItem"
      _StyleDefs(435) =   ":id=13,.parent=39"
      _StyleDefs(436) =   "Named:id=412:orgStyles"
      _StyleDefs(437) =   ":id=412,.parent=33,.bgcolor=&HC0C0C0&"
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "～"
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
      Index           =   5
      Left            =   8505
      TabIndex        =   14
      Top             =   840
      Width           =   240
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "表示範囲"
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
      Index           =   4
      Left            =   6090
      TabIndex        =   13
      Top             =   840
      Width           =   960
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "事業部"
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
      Left            =   10920
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
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
      Left            =   4050
      TabIndex        =   9
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "親品番"
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
      Left            =   300
      TabIndex        =   8
      Top             =   840
      Width           =   720
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "画面印刷"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   1
      End
   End
End
Attribute VB_Name = "ODR30201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NAIGAI_CODE()   As String * 1
Private NAIGAI_NAME()   As String

'コンボ用添字
Private Const pcmbJI = 0            '事業部
Private Const pcmbSM = 1            '仕向け先

'テキスト用添字
Private Const ptxTOP% = 0
Private Const ptxLAST% = 1

Private Const ptxOYA_CD% = 0
Private Const ptxUSE_YY% = 1

Private Const ptxS_Date% = 2
Private Const ptxE_Date% = 3

'ラベル用添字

'コマンドボタン用添字


'グリッド用定義
Private ORDR_GRID   As New XArrayDB

Private Const Min_Row% = 1              '最小行数
Private Max_Row As Long                 '最大表示行数

Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 94             '最大列数


Private Const colORDER% = 0             '仕入先
Private Const colHIN_GAI% = 1           '品番
Private Const colHIN_NAME% = 2          '品名
Private Const colTITLE% = 3             '

Private Const colQTY% = 4               '有効在庫／引当可能数／入庫数／出庫数

Private Const colStyle% = 94            '行のスタイル



Dim Mode        As Boolean
Dim Row         As Long                 '対象　行



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
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim yn          As Integer

Dim W_STR       As String


    ERR_CHK = True
    
    
    Select Case Index
        Case ptxOYA_CD
        
            
        Case ptxUSE_YY
            If Trim(Text1(Index)) = "" Then
            Else
                W_STR = Text1(ptxUSE_YY) & "/01"
                
                If Not IsDate(W_STR) Then
                    MsgBox "使用月エラー！", vbExclamation
                    Exit Function
                End If
                
                W_STR = Format(W_STR, "yyyy/mm/dd")
                Text1(ptxUSE_YY) = Left(W_STR, 7)
            End If
            
    
        Case ptxS_Date
                        
            If Not IsDate(Text1(Index).Text) Then
                MsgBox "日付エラー！", vbExclamation
                Exit Function
            Else
                Text1(Index).Text = Format(Text1(Index).Text, "YYYY/MM/DD")
            End If
        
        
        
        Case ptxE_Date
    
            If Not IsDate(Text1(Index).Text) Then
                MsgBox "日付エラー！", vbExclamation
                Exit Function
            Else
                Text1(Index).Text = Format(Text1(Index).Text, "YYYY/MM/DD")
            End If
    
    
            If DateDiff("d", Text1(ptxS_Date).Text, Text1(ptxE_Date).Text) > 90 Then
                MsgBox "日付範囲エラー！", vbExclamation
                Exit Function
            End If
            
    
    End Select
    
    ERR_CHK = False
    
End Function

Private Function Data_Disp() As Integer
'----------------------------------------------------------------------------
'                   画面表示
'----------------------------------------------------------------------------
Dim com                 As Integer
Dim sts                 As Integer
Dim yn                  As Integer


Dim W_Key               As String

Dim W_STR               As String

Dim cnt                 As Integer

Dim wkDate              As String
Dim i                   As Integer

Dim Skip_F              As Boolean
Dim Fast_F              As Boolean

Dim svJGYOBU            As String * 1
Dim svNAIGAI            As String * 1
Dim svHIN_GAI           As String * 20

Dim svORDER_CODE        As String * 5

Dim sumY_ZAIKO_QTY()    As Double
Dim sumHIKIATE_QTY()    As Double
Dim sumNYUKO_QTY()      As Double
Dim sumSYUKO_QTY()      As Double

Dim Style_Flg           As Integer


    Data_Disp = True
    
        
    
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "注文情報　検索中。<Data_Disp>", Me.hwnd, 0)
    DoEvents
    
    Call Input_Lock                             '画面項目ロック
    
    
    For i = 0 To 89
        wkDate = Format(DateAdd("d", i, Text1(ptxS_Date).Text), "YYYY/MM/DD")
        
        If wkDate > Text1(ptxE_Date).Text Then
            TDBGrid1.Columns(i + colQTY).Visible = False
        Else
            TDBGrid1.Columns(i + colQTY).Visible = True
            If i = 0 Or _
                Right(wkDate, 2) = "01" Then
                TDBGrid1.Columns(i + colQTY).Caption = Right(wkDate, 5)
            Else
                TDBGrid1.Columns(i + colQTY).Caption = Right(wkDate, 2)
            End If
        End If
        
        ReDim sumY_ZAIKO_QTY(0 To i)
        ReDim sumHIKIATE_QTY(0 To i)
        ReDim sumNYUKO_QTY(0 To i)
        ReDim sumSYUKO_QTY(0 To i)
        
        sumY_ZAIKO_QTY(i) = 0
        sumHIKIATE_QTY(i) = 0
        sumNYUKO_QTY(i) = 0
        sumSYUKO_QTY(i) = 0
    
    Next i
        
        
    Set ORDR_GRID = Nothing
    
    
    
    Row = Min_Row - 1
    Fast_F = True
    
    com = BtOpGetFirst
    
    Do
        
        DoEvents
    
        sts = BTRV(com, ODR_BUHIN_SUII_POS, ODR_BUHIN_SUII_REC, Len(ODR_BUHIN_SUII_REC), K1_ODR_BUHIN_SUII, Len(K1_ODR_BUHIN_SUII), 1)
            
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "ODR_BUHIN_SUII")
                Exit Function
        
        End Select


        Skip_F = False
        If Trim(Text1(ptxUSE_YY).Text) <> "" Then
            If Left(Format(Text1(ptxUSE_YY).Text & "01", "YYYYMMDD"), 6) <> StrConv(ODR_BUHIN_SUII_REC.USE_YM, vbUnicode) Then
                Skip_F = True
            End If
        End If


        If StrConv(ODR_BUHIN_SUII_REC.SEL_DATE, vbUnicode) < Format(Text1(ptxS_Date).Text, "YYYYMMDD") Or _
            StrConv(ODR_BUHIN_SUII_REC.SEL_DATE, vbUnicode) > Format(Text1(ptxE_Date).Text, "YYYYMMDD") Then
            
            Skip_F = True
        
        End If


        If Not Skip_F Then
            
            
            If Fast_F Then
                            
                svJGYOBU = StrConv(ODR_BUHIN_SUII_REC.KO_JGYOBU, vbUnicode)
                svNAIGAI = StrConv(ODR_BUHIN_SUII_REC.KO_NAIGAI, vbUnicode)
                svHIN_GAI = StrConv(ODR_BUHIN_SUII_REC.KO_HIN_GAI, vbUnicode)
            
                svORDER_CODE = StrConv(ODR_BUHIN_SUII_REC.ORDER_CODE, vbUnicode)
            
                Fast_F = False
            
                Style_Flg = 0
            
            End If
            
            If svJGYOBU <> StrConv(ODR_BUHIN_SUII_REC.KO_JGYOBU, vbUnicode) Or _
                svNAIGAI <> StrConv(ODR_BUHIN_SUII_REC.KO_NAIGAI, vbUnicode) Or _
                svHIN_GAI <> StrConv(ODR_BUHIN_SUII_REC.KO_HIN_GAI, vbUnicode) Then
            
                Row = Row + 4
                
                If Style_Flg = 0 Then
                    Style_Flg = 1
                Else
                    Style_Flg = 0
                End If
                
                If Grid_Set_Proc(Row, svJGYOBU, svNAIGAI, svHIN_GAI, svORDER_CODE, sumY_ZAIKO_QTY, sumHIKIATE_QTY, sumNYUKO_QTY, sumSYUKO_QTY, Style_Flg) Then
                    GoTo Err_Exit
                End If
        
                svJGYOBU = StrConv(ODR_BUHIN_SUII_REC.KO_JGYOBU, vbUnicode)
                svNAIGAI = StrConv(ODR_BUHIN_SUII_REC.KO_NAIGAI, vbUnicode)
                svHIN_GAI = StrConv(ODR_BUHIN_SUII_REC.KO_HIN_GAI, vbUnicode)
        
                svORDER_CODE = StrConv(ODR_BUHIN_SUII_REC.ORDER_CODE, vbUnicode)
        
        
                For i = 0 To UBound(sumY_ZAIKO_QTY)
        
                    sumY_ZAIKO_QTY(i) = 0
                    sumHIKIATE_QTY(i) = 0
                    sumNYUKO_QTY(i) = 0
                    sumSYUKO_QTY(i) = 0
        
                Next i
            End If
        
        
            i = DateDiff("d", Text1(ptxS_Date).Text, Mid(StrConv(ODR_BUHIN_SUII_REC.SEL_DATE, vbUnicode), 1, 4) & "/" & _
                                                    Mid(StrConv(ODR_BUHIN_SUII_REC.SEL_DATE, vbUnicode), 5, 2) & "/" & _
                                                    Mid(StrConv(ODR_BUHIN_SUII_REC.SEL_DATE, vbUnicode), 7, 2))
        
        
            sumY_ZAIKO_QTY(i) = sumY_ZAIKO_QTY(i) + CDbl(StrConv(ODR_BUHIN_SUII_REC.Y_ZAIKO_QTY, vbUnicode))
            sumHIKIATE_QTY(i) = sumHIKIATE_QTY(i) + CDbl(StrConv(ODR_BUHIN_SUII_REC.HIKIATE_QTY, vbUnicode))
            sumNYUKO_QTY(i) = sumNYUKO_QTY(i) + CDbl(StrConv(ODR_BUHIN_SUII_REC.NYUKO_QTY, vbUnicode))
            sumSYUKO_QTY(i) = sumSYUKO_QTY(i) + CDbl(StrConv(ODR_BUHIN_SUII_REC.SYUKO_QTY, vbUnicode))
        
        
        End If

        com = BtOpGetNext
    
    Loop
    
    If Not Fast_F Then
                
        Row = Row + 4
        
        If Style_Flg = 0 Then
            Style_Flg = 1
        Else
            Style_Flg = 0
        End If
        
        
        If Grid_Set_Proc(Row, svJGYOBU, svNAIGAI, svHIN_GAI, svORDER_CODE, sumY_ZAIKO_QTY, sumHIKIATE_QTY, sumNYUKO_QTY, sumSYUKO_QTY, Style_Flg) Then
            GoTo Err_Exit
        End If
    End If
    
    Set TDBGrid1.Array = ORDR_GRID
    
    TDBGrid1.style.Locked = True
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    TDBGrid1.Bookmark = 1
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "注文情報　表示中。     <Data_Disp>", Me.hwnd, 0)
    DoEvents
    
    
    Data_Disp = False
    
Err_Exit:
    
    Call Input_UnLock                             '画面項目ロック
End Function

Private Function Grid_Set_Proc(Row As Long, _
                                svJGYOBU As String, _
                                svNAIGAI As String, _
                                svHIN_GAI As String, _
                                svORDER As String, _
                                sumY_ZAIKO_QTY() As Double, _
                                sumHIKIATE_QTY() As Double, _
                                sumNYUKO_QTY() As Double, _
                                sumSYUKO_QTY() As Double, _
                                Style_Flg As Integer) As Integer
'----------------------------------------------------------------------------
'                   グリッド表示
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim i           As Integer

    Grid_Set_Proc = True

    ORDR_GRID.ReDim Min_Row, Row, Min_Col, Max_Col

    
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, svORDER)
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
        
        
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_NAME, StrConv(ODR_BUHIN_SUII_REC.ORDER_CODE, vbUnicode))
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "P_UKEHARAI")
            Exit Function
    
    End Select
    ORDR_GRID(Row - 3, colORDER) = Trim(StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode))
    
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, svJGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, svNAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, svHIN_GAI)
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
        
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "ITEM")
            Exit Function
    
    End Select
    
    
    ORDR_GRID(Row - 3, colHIN_GAI) = Trim(svHIN_GAI)
    ORDR_GRID(Row - 3, colHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    ORDR_GRID(Row - 3, colTITLE) = "在庫数"
    
    
    
    For i = 0 To UBound(sumY_ZAIKO_QTY)
    
        ORDR_GRID(Row - 3, colQTY + i) = sumY_ZAIKO_QTY(i)
    
    
    Next i
    
    ORDR_GRID(Row - 3, colStyle) = Style_Flg
    
    
    
    ORDR_GRID(Row - 2, colTITLE) = "引当残"
    For i = 0 To UBound(sumHIKIATE_QTY)
    
        ORDR_GRID(Row - 2, colQTY + i) = sumHIKIATE_QTY(i)
    
    
    Next i
    ORDR_GRID(Row - 2, colStyle) = Style_Flg
    
    
    ORDR_GRID(Row - 1, colTITLE) = "入　庫"
    For i = 0 To UBound(sumHIKIATE_QTY)
    
        ORDR_GRID(Row - 1, colQTY + i) = sumNYUKO_QTY(i)
    
    
    Next i
    ORDR_GRID(Row - 1, colStyle) = Style_Flg
    
    ORDR_GRID(Row, colTITLE) = "出　庫"
    For i = 0 To UBound(sumHIKIATE_QTY)
    
        ORDR_GRID(Row, colQTY + i) = sumSYUKO_QTY(i)
    
    
    Next i
    ORDR_GRID(Row, colStyle) = Style_Flg
    
    
        

    Grid_Set_Proc = False

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    ODR30201.MousePointer = vbHourglass

    Call Ctrl_Lock(ODR30201)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(ODR30201)


    ODR30201.MousePointer = vbDefault

End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Tab_Ctrl(Shift)        '移動
    End If

End Sub

Private Sub Command1_Click(Index As Integer)


Dim yn      As Integer
Dim i       As Integer

    Select Case Index
    
        Case 0
            
            For i = ptxOYA_CD To ptxE_Date
                If ERR_CHK(i) Then
                    Exit Sub
                End If
            Next i
            
            
            yn = MsgBox("表示しますか？", vbYesNo + vbDefaultButton1 + vbQuestion, "確認入力")
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            
            If Data_Make_Proc() Then
                Unload Me
            End If
            
            If Data_Disp() Then
                Unload Me
            End If
            
            

            
        Case 1
            
            Unload Me
    
    End Select

End Sub

Private Sub Form_Load()

Dim i       As Integer
Dim c       As String * 128
Dim sts     As Integer

Dim sBuffer As String * 255
Dim com     As String

Dim W_Date  As String

'ステータスウィンドウを作成する
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "子部品　在庫推移照会", Me.hwnd, 0)
'最後の要素を-1にすると
'親ウィンドウの全体の幅の残りの幅を
'自動的に割り当てる
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
    
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                '子部品注文データＯＰＥＮ
    If P_SHORDER_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '構成マスタＯＰＥＮ
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '受払先マスタＯＰＥＮ
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '子部品受入履歴ＯＰＥＮ
    If P_SHUKEIRE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '親部品注文データＯＰＥＮ
    If ODR_ORDER_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '事業部の獲得
    If JGYOB_TB_Set() Then
        MsgBox "事業部の獲得に失敗しました。"
        End
    End If
    Combo1(pcmbJI).Clear
    
    For i = 0 To UBound(JGYOBU_T) - 1
        Combo1(pcmbJI).AddItem JGYOBU_T(i).NAME & Space(5) & JGYOBU_T(i).CODE
    
    Next i
    Combo1(pcmbJI).ListIndex = 0

                                '国内外管理の獲得
    i = 0
    Do
        i = i + 1
        If GetIni(App.EXEName, "NAIGAI" & Format(i, "0"), "SYS_ODR3020", c) Then
            Exit Do
        End If
        ReDim Preserve NAIGAI_CODE(i - 1)
        NAIGAI_CODE(i - 1) = Trim(c)
    
    Loop
    If i = 1 Then
        MsgBox "国内外の獲得に失敗しました。"
        End
    End If
    
    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc
    
    '仕向け先のセット
    If Code_Set_Proc(pcmbSM, P_KBN04_CD, 0) Then
        Unload Me
    End If
    Combo1(pcmbSM).ListIndex = 0
    
    GW_SIMUKE = Left(Right(Combo1(pcmbSM).Text, 4), 2)
    GW_JIGYOBU = Right(Combo1(pcmbJI).Text, 1)
    
'テキストを設定する
    Text1(ptxUSE_YY) = ""
    
    
    Text1(ptxS_Date).Text = Mid(Format(Now, "YYYY/MM/DD"), 1, 8) & "01"
    Text1(ptxE_Date).Text = DateAdd("d", -1, DateAdd("m", 1, Text1(ptxS_Date).Text))
    
    
    
    
    
    Text1(ptxTOP).SetFocus
    Call Text1_GotFocus(ptxTOP)
    'Combo1(pcmbSM).SetFocus
    
    'ORDR_GRID.RowHeight = ORDR_GRID.RowHeight * 3
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim yn      As Integer

    If UnloadMode = 1 Then Exit Sub
    
    yn = MsgBox("終了しますか？", vbYesNo + vbDefaultButton1 + vbQuestion, "確認入力")
    If yn = vbNo Then
        Cancel = 1
        Exit Sub
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts     As Integer
    
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ITEM")
        End If
    End If
    
    sts = BTRV(BtOpClose, ODR_BUHIN_SUII_POS, ODR_BUHIN_SUII_REC, Len(ODR_BUHIN_SUII_REC), K0_ODR_BUHIN_SUII, Len(K0_ODR_BUHIN_SUII), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_BUHIN_SUII")
        End If
    End If

    
    sts = BTRV(BtOpClose, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_ORDER, Len(K0_ODR_ORDER), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_BUHIN_SUII")
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
            Call File_Error(sts, BtOpClose, "P_SHUKEIRE")
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


    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "")
        End If
    End If




    End
End Sub

Private Sub SHORI_Click(Index As Integer)
Dim yn      As Integer


    Select Case Index
            
        Case 0      '画面印刷
            yn = MsgBox("画面印刷しますか？", vbYesNo + vbDefaultButton2 + vbQuestion, "確認入力")
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)
        
        
        Case 1      '終了
            Call Command1_Click(1)
    
    End Select


End Sub

Private Sub TDBGrid1_DblClick()

    If TDBGrid1.Bookmark = -1 Then
    Else
        
        'ORD20102.Show vbModal
        
        'If KENPIN_Update_Proc() Then
        '    Unload Me
        'End If
    End If
    
    '再表示
'    If List_Disp Then
'        Unload Me
'    End If


End Sub


Private Sub TDBGrid1_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid80.StyleDisp)

    If ORDR_GRID(Bookmark, colStyle) = 0 Then
        CellStyle = TDBGrid1.Styles("Normal")
    Else
        CellStyle = TDBGrid1.Styles("orgStyles")
    End If

End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    TDBGrid1.Bookmark = -1
    
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
    
    If Index = ptxOYA_CD% Then
        If Data_Disp Then
            MsgBox "指定条件の注文情報で、表示失敗！", vbExclamation
            Call Text1_GotFocus(ptxTOP%)
            Text1(ptxTOP%).SetFocus
            Exit Sub
        End If
        
        
        TDBGrid1.SetFocus
        
        Exit Sub
    End If
    
    Call Tab_Ctrl(Shift)    '移動
    
End Sub


Private Function Data_Make_Proc() As Integer
'----------------------------------------------------------------------------
'                   表示用データの作成
'----------------------------------------------------------------------------
Dim sts                 As Integer
Dim com                 As Integer
Dim ans                 As Integer


Dim Sumi_Zaiko_Qty      As Long
Dim Mi_Zaiko_Qty        As Long

    
    
Dim FullPath            As String
    
Dim sBuffer             As String * 255
Dim WsNo                As String
Dim c                   As String * 128

Dim Ret                 As Integer
    
Dim i                   As Integer
    
    
Dim svJGYOBU            As String * 1
Dim svNAIGAI            As String * 1
Dim svHIN_GAI           As String * 1
    
Dim wkYM                As String * 7
    
Dim S_DATE              As String
Dim E_DATE              As String
    
    
    Data_Make_Proc = True
    
    Call Input_Lock
    
    
    
    
    
    S_DATE = Text1(ptxS_Date).Text
    
    E_DATE = Text1(ptxE_Date).Text
    
    
    
    
    sts = BTRV(BtOpClose, ODR_BUHIN_SUII_POS, ODR_BUHIN_SUII_REC, Len(ODR_BUHIN_SUII_REC), K0_ODR_BUHIN_SUII, Len(K0_ODR_BUHIN_SUII), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_BUHIN_SUII")
        End If
    End If
    
    sts = BTRV(BtOpClose, wODR_BUHIN_SUII_POS, wODR_BUHIN_SUII_REC, Len(wODR_BUHIN_SUII_REC), K0_wODR_BUHIN_SUII, Len(K0_wODR_BUHIN_SUII), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_BUHIN_SUII")
        End If
    End If
    
    
                                                        '子部品推移データ　フルパス取込み
    sts = GetIni("FILE", ODR_BUHIN_SUII_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_BUHIN_SUII_]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        WsNo = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        WsNo = "???"
    End If


    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & WsNo & Right(Trim(c), Len(Trim(c)) - Ret)
    
   On Error Resume Next
   Kill (FullPath)
   On Error GoTo 0
    
    
    If ODR_BUHIN_SUII_Open(BtOpenNomal) Then
        Exit Function
    End If
    
    If wODR_BUHIN_SUII_Open(BtOpenNomal) Then
        Exit Function
    End If
    
    
    
    
    
    sts = BTRV(BtOpClose, ODR_BUHIN_ORDER_POS, ODR_BUHIN_ORDER_REC, Len(ODR_BUHIN_ORDER_REC), K0_ODR_BUHIN_ORDER, Len(K0_ODR_BUHIN_ORDER), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_BUHIN_ORDER")
        End If
    End If
    
    
                                                        '子部品注文（一時）データ　フルパス取込み
    sts = GetIni("FILE", ODR_BUHIN_ORDER_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [ODR_BUHIN_ORDER]読み込みエラー")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        WsNo = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        WsNo = "???"
    End If


    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & WsNo & Right(Trim(c), Len(Trim(c)) - Ret)
    
    On Error Resume Next
    Kill (FullPath)
    On Error GoTo 0
    
    
    If ODR_BUHIN_ORDER_Open(BtOpenNomal) Then
        Exit Function
    End If
    
    
    
    If tmpBUHIN_ORDER_Make_Proc(S_DATE, E_DATE) Then
        Exit Function
    End If
    
    
    If tmpZaiko_Suii_Make_Proc(S_DATE, E_DATE) Then
    
    
        Exit Function
    
    
    End If




    Call Input_UnLock



    Data_Make_Proc = False


End Function

Private Function tmpZaiko_Suii_Make_Proc(S_DATE As String, E_DATE As String) As Integer
'----------------------------------------------------------------------------
'                   在庫推移データ作成
'----------------------------------------------------------------------------
Dim sts                 As Integer
Dim com                 As Integer


Dim wkDate              As String
Dim i                   As Integer

Dim yn                  As Integer


Dim Sumi_Zaiko_Qty      As Long
Dim Mi_Zaiko_Qty        As Long

Dim wkDouble            As Double

Dim tblY_ZAIKO_QTY()    As Double
Dim tblHIKIATE_QTY()    As Double
Dim tblNYUKO_QTY()      As Double
Dim tblSYUKO_QTY()      As Double
Dim Day_Max             As Integer


Dim svJGYOBU            As String * 1
Dim svNAIGAI            As String * 1
Dim svHIN_GAI           As String * 20



    tmpZaiko_Suii_Make_Proc = True


    If Text1(ptxS_Date).Text > Format(Now, "YYYY/MM/DD") Then
        S_DATE = Format(Now, "YYYY/MM/DD")
    End If
    
    E_DATE = Text1(ptxE_Date).Text
    

    Day_Max = DateDiff("d", S_DATE, E_DATE)

'---------------------------------------    完了日で作成(親部品受注分)
    Call UniCode_Conv(K4_ODR_ORDER.FIN_DT, Format(S_DATE, "YYYYMMDD"))

    com = BtOpGetGreaterEqual


    Do
        DoEvents
    
    
        sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K4_ODR_ORDER, Len(K4_ODR_ORDER), 4)

        Select Case sts
            Case BtNoErr

            
                If StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode) > Format(E_DATE, "YYYYMMDD") Then
                    Exit Do
                End If
            
            
            Case BtErrEOF
            
                Exit Do




            Case Else
                Call File_Error(sts, BtOpGetEqual, "ODR_ORDER")
                Exit Function

        End Select
    
    
            
            
            
        '構成マスタ読み込み＜子部品展開＞
        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode))
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "")
        Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
        
        com = BtOpGetGreaterEqual

        Do
            DoEvents
        
        
            sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    
            Select Case sts
                Case BtNoErr
    
                    If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode) Or _
                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode) Or _
                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode) Or _
                        StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode) <> StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode) Then
                        Exit Do
                    End If
                
                
                
                
                Case BtErrEOF
                
                    Exit Do
    
    
    
    
                Case Else
                    Call File_Error(sts, com, "ODR_ORDER")
                    Exit Function
    
            End Select
        
        
        
            If StrConv(P_COMPO_K_REC.SEQNO, vbUnicode) = "000" Then
            Else
                '子部品推移読み込み＜存在チェック＞
                Call UniCode_Conv(K0_ODR_BUHIN_SUII.SEL_DATE, Format(S_DATE, "YYYYMMDD"))
'2008.04.04                Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
'2008.04.04                Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                
                Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_JGYOBU, SHIZAI)      '2008.04.04
                Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_NAIGAI, NAIGAI_NAI)  '2008.04.04
                
                
                Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
        
                sts = BTRV(BtOpGetEqual, ODR_BUHIN_SUII_POS, ODR_BUHIN_SUII_REC, Len(ODR_BUHIN_SUII_REC), K0_ODR_BUHIN_SUII, Len(K0_ODR_BUHIN_SUII), 0)
        
                Select Case sts
                    Case BtNoErr
                    
                    
                    Case BtErrKeyNotFound
        
        
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "ODR_BUHIN_SUII")
                        Exit Function
                End Select
                If sts = BtErrKeyNotFound Then
                '子部品推移作成＜日付範囲内全レコード＞
                    wkDate = S_DATE
                            
                
                    i = 0
                    Do
                    
                        wkDate = Format(DateAdd("d", i, wkDate))
                                            
                    
                        If wkDate > E_DATE Then
                            Exit Do
                        End If
                    
                    
                    
                        Call UniCode_Conv(ODR_BUHIN_SUII_REC.SEL_DATE, Format(wkDate, "YYYYMMDD"))

'2008.04.04                        Call UniCode_Conv(ODR_BUHIN_SUII_REC.KO_JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
'2008.04.04                        Call UniCode_Conv(ODR_BUHIN_SUII_REC.KO_NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                        Call UniCode_Conv(ODR_BUHIN_SUII_REC.KO_JGYOBU, SHIZAI)         '2008.04.04.
                        Call UniCode_Conv(ODR_BUHIN_SUII_REC.KO_NAIGAI, NAIGAI_NAI)     '2008.04.04
                        
                        Call UniCode_Conv(ODR_BUHIN_SUII_REC.KO_HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))

                        If Trim(Text1(ptxUSE_YY).Text) = "" Then
                           Call UniCode_Conv(ODR_BUHIN_SUII_REC.USE_YM, "")
                        Else
                            Call UniCode_Conv(ODR_BUHIN_SUII_REC.USE_YM, Mid(Format(Text1(ptxUSE_YY).Text & "/01", "YYYYMMDD"), 1, 6))
                        End If
                    
'2008.04.04                        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
'2008.04.04                        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                    
                        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)           '2008.04.04
                        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)       '2008.04.04
                        
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    
                                                
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                
                        Select Case sts
                            Case BtNoErr
                
                            Case BtErrKeyNotFound
                
                                Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).CODE, "")
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "ITEM")
                                Exit Function
                        End Select
                    
                    
                    
                        Call UniCode_Conv(ODR_BUHIN_SUII_REC.ORDER_CODE, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
                    
                        Call UniCode_Conv(ODR_BUHIN_SUII_REC.Y_ZAIKO_QTY, "00000.00")
                        Call UniCode_Conv(ODR_BUHIN_SUII_REC.HIKIATE_QTY, "00000.00")
                        Call UniCode_Conv(ODR_BUHIN_SUII_REC.NYUKO_QTY, "00000.00")
                        Call UniCode_Conv(ODR_BUHIN_SUII_REC.SYUKO_QTY, "00000.00")








                        
                        Do
                        
                            sts = BTRV(BtOpInsert, ODR_BUHIN_SUII_POS, ODR_BUHIN_SUII_REC, Len(ODR_BUHIN_SUII_REC), K0_ODR_BUHIN_SUII, Len(K0_ODR_BUHIN_SUII), 0)
    
                            Select Case sts
                                Case BtNoErr
                                            
                                    Exit Do
                                            
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                                    yn = MsgBox("他で使用中です！<ODR_TEMP3>" & Chr(13) & Chr(10) & _
                                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                                    If yn = vbNo Then
                                        Exit Do
                                    End If
                    
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "ODR_BUHIN_SUII")
                                    Exit Function
                            End Select
                        
                        Loop
                    
                    
                        i = i + 1
                    Loop
                End If
                
                com = BtOpGetNext
    
    
            End If
    
    
        Loop
    
    
        
        com = BtOpGetNext
    
    
    
    Loop



'---------------------------------------    回答納期で作成(親部品受注分)
    Call UniCode_Conv(K3_ODR_ORDER.KAITO_DT, Format(S_DATE, "YYYYMMDD"))

    com = BtOpGetGreaterEqual


    Do
        DoEvents
    
    
        sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K3_ODR_ORDER, Len(K3_ODR_ORDER), 3)

        Select Case sts
            Case BtNoErr

            
                If StrConv(ODR_ORDER_REC.KAITO_DT, vbUnicode) > Format(E_DATE, "YYYYMMDD") Then
                    Exit Do
                End If
            
            
            Case BtErrEOF
            
                Exit Do




            Case Else
                Call File_Error(sts, BtOpGetEqual, "ODR_ORDER")
                Exit Function

        End Select
    
    
        If Trim(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode)) <> "" Then
        Else
            
            '構成マスタ読み込み＜子部品展開＞
            Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode))
            Call UniCode_Conv(K0_P_COMPO.JGYOBU, StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_P_COMPO.NAIGAI, StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_P_COMPO.HIN_GAI, StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode))
            Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "")
            Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
            
            com = BtOpGetGreaterEqual
    
            Do
                DoEvents
            
            
                sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        
                Select Case sts
                    Case BtNoErr
        
                        If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode) Or _
                            StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode) Or _
                            StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode) Or _
                            StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode) <> StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode) Then
                            Exit Do
                        End If
                    
                    
                    
                    
                    Case BtErrEOF
                    
                        Exit Do
        
        
        
        
                    Case Else
                        Call File_Error(sts, com, "ODR_ORDER")
                        Exit Function
        
                End Select
            
            
            
                If StrConv(P_COMPO_K_REC.SEQNO, vbUnicode) = "000" Then
                Else
                    '子部品推移読み込み＜存在チェック＞
                    Call UniCode_Conv(K0_ODR_BUHIN_SUII.SEL_DATE, Format(S_DATE, "YYYYMMDD"))
                    
'2008.04.04                    Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
'2008.04.04                    Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_JGYOBU, SHIZAI)      '2008.04.04
                    Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_NAIGAI, NAIGAI_NAI)  '2008.04.04
                    
                    Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
            
                    sts = BTRV(BtOpGetEqual, ODR_BUHIN_SUII_POS, ODR_BUHIN_SUII_REC, Len(ODR_BUHIN_SUII_REC), K0_ODR_BUHIN_SUII, Len(K0_ODR_BUHIN_SUII), 0)
            
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound
            
            
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "ODR_BUHIN_SUII")
                            Exit Function
                    End Select
                    If sts = BtErrKeyNotFound Then
                    '子部品推移作成＜日付範囲内全レコード＞
                        wkDate = S_DATE
                                
                    
                        i = 0
                        Do
                        
                            wkDate = DateAdd("d", i, S_DATE)
                                                
                        
                            If wkDate > E_DATE Then
                                Exit Do
                            End If
                        
                        
                        
                            Call UniCode_Conv(ODR_BUHIN_SUII_REC.SEL_DATE, Format(wkDate, "YYYYMMDD"))
'2008.04.04                            Call UniCode_Conv(ODR_BUHIN_SUII_REC.KO_JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
'2008.04.04                            Call UniCode_Conv(ODR_BUHIN_SUII_REC.KO_NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                            Call UniCode_Conv(ODR_BUHIN_SUII_REC.KO_JGYOBU, SHIZAI)         '2008.04.04
                            Call UniCode_Conv(ODR_BUHIN_SUII_REC.KO_NAIGAI, NAIGAI_NAI)     '2008.04.04
                            
                            
                            
                            Call UniCode_Conv(ODR_BUHIN_SUII_REC.KO_HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))

                            If Trim(Text1(ptxUSE_YY).Text) = "" Then
                               Call UniCode_Conv(ODR_BUHIN_SUII_REC.USE_YM, "")
                            Else
                                Call UniCode_Conv(ODR_BUHIN_SUII_REC.USE_YM, Mid(Format(Text1(ptxUSE_YY).Text & "/01", "YYYYMMDD"), 1, 6))
                            End If
                        
'2008.04.04                            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
'2008.04.04                            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                            
                            Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)       '2008.04.04
                            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)   '2008.04.04
                            
                            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                        
                                                    
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    
                            Select Case sts
                                Case BtNoErr
                    
                                Case BtErrKeyNotFound
                    
                                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).CODE, "")
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "ITEM")
                                    Exit Function
                            End Select
                        
                        
                        
                            Call UniCode_Conv(ODR_BUHIN_SUII_REC.ORDER_CODE, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
                        
                            Call UniCode_Conv(ODR_BUHIN_SUII_REC.Y_ZAIKO_QTY, "00000.00")
                            Call UniCode_Conv(ODR_BUHIN_SUII_REC.HIKIATE_QTY, "00000.00")
                            Call UniCode_Conv(ODR_BUHIN_SUII_REC.NYUKO_QTY, "00000.00")
                            Call UniCode_Conv(ODR_BUHIN_SUII_REC.SYUKO_QTY, "00000.00")

                            
                            Do
                            
                                sts = BTRV(BtOpInsert, ODR_BUHIN_SUII_POS, ODR_BUHIN_SUII_REC, Len(ODR_BUHIN_SUII_REC), K0_ODR_BUHIN_SUII, Len(K0_ODR_BUHIN_SUII), 0)
        
                                Select Case sts
                                    Case BtNoErr
                                                
                                        Exit Do
                                                
                                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                                        yn = MsgBox("他で使用中です！<ODR_TEMP3>" & Chr(13) & Chr(10) & _
                                                    "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                                        If yn = vbNo Then
                                            Exit Do
                                        End If
                        
                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual, "ODR_BUHIN_SUII")
                                        Exit Function
                                End Select
                            
                            Loop
                        
                        
                            i = i + 1
                        Loop
                    End If
                
                End If
                
                
                
                
                
                
                
                
                
                
                
                com = BtOpGetNext
    
    
            Loop
    
    
    
    
    
        End If
        
        com = BtOpGetNext
    
    
    
    Loop


'---------------------------------------    今日日付に現在庫を設定
    Call UniCode_Conv(K0_ODR_BUHIN_SUII.SEL_DATE, Format(Now, "YYYYMMDD"))
    Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_JGYOBU, "")
    Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_NAIGAI, "")
    Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_HIN_GAI, "")

    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, ODR_BUHIN_SUII_POS, ODR_BUHIN_SUII_REC, Len(ODR_BUHIN_SUII_REC), K0_ODR_BUHIN_SUII, Len(K0_ODR_BUHIN_SUII), 0)

        Select Case sts
            Case BtNoErr

                If StrConv(ODR_BUHIN_SUII_REC.SEL_DATE, vbUnicode) <> Format(Now, "YYYYMMDD") Then
                    Exit Do
                End If


            Case BtErrEOF

                Exit Do
            
            Case Else
                Call File_Error(sts, com, "ODR_BUHIN_SUII")
                Exit Function
        
        End Select
    
    
    
    
        If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, Mi_Zaiko_Qty, StrConv(ODR_BUHIN_SUII_REC.KO_JGYOBU, vbUnicode), _
                                                            StrConv(ODR_BUHIN_SUII_REC.KO_NAIGAI, vbUnicode), _
                                                            StrConv(ODR_BUHIN_SUII_REC.KO_HIN_GAI, vbUnicode)) Then
                Exit Function
        End If
    
    
        Call UniCode_Conv(ODR_BUHIN_SUII_REC.Y_ZAIKO_QTY, Format(Sumi_Zaiko_Qty + Mi_Zaiko_Qty, "00000.00"))
    
        Do
        
            sts = BTRV(BtOpUpdate, ODR_BUHIN_SUII_POS, ODR_BUHIN_SUII_REC, Len(ODR_BUHIN_SUII_REC), K0_ODR_BUHIN_SUII, Len(K0_ODR_BUHIN_SUII), 0)

            Select Case sts
                Case BtNoErr
                            
                    Exit Do
                            
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                    yn = MsgBox("他で使用中です！<ODR_TEMP3>" & Chr(13) & Chr(10) & _
                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                    If yn = vbNo Then
                        Exit Do
                    End If
    
                Case Else
                    Call File_Error(sts, BtOpUpdate, "ODR_BUHIN_SUII")
                    Exit Function
            End Select
        
        Loop
    
    
        com = BtOpGetNext
    
    
    Loop

'---------------------------------------    出庫予定OR 実績を設定（完了日）
    Call UniCode_Conv(K4_ODR_ORDER.FIN_DT, Format(S_DATE, "YYYYMMDD"))

    com = BtOpGetGreaterEqual


    Do
        DoEvents


        sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K4_ODR_ORDER, Len(K4_ODR_ORDER), 4)

        Select Case sts
            Case BtNoErr

            
                If StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode) > Format(E_DATE, "YYYYMMDD") Then
                    Exit Do
                End If
            
            
            Case BtErrEOF
            
                Exit Do




            Case Else
                Call File_Error(sts, BtOpGetEqual, "ODR_ORDER")
                Exit Function

        End Select
    
    
            
            
            
        '構成マスタ読み込み＜子部品展開＞
        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode))
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "")
        Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
        
        com = BtOpGetGreaterEqual

        Do
            DoEvents
        
        
            sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    
            Select Case sts
                Case BtNoErr
    
                    If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode) Or _
                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode) Or _
                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode) Or _
                        StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode) <> StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode) Then
                        Exit Do
                    End If
                
                
                
                
                Case BtErrEOF
                
                    Exit Do
    
    
    
    
                Case Else
                    Call File_Error(sts, com, "ODR_ORDER")
                    Exit Function
    
            End Select
        
        
        
            If StrConv(P_COMPO_K_REC.SEQNO, vbUnicode) = "000" Then
            Else
                '子部品推移読み込み＜存在チェック＞
                Call UniCode_Conv(K0_ODR_BUHIN_SUII.SEL_DATE, StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode))
                
'2008.04.04                Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
'2008.04.04                Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                
                Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_JGYOBU, SHIZAI)      '2008.04.04
                Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_NAIGAI, NAIGAI_NAI)  '2008.04.04
                
                
                Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
        
                sts = BTRV(BtOpGetEqual, ODR_BUHIN_SUII_POS, ODR_BUHIN_SUII_REC, Len(ODR_BUHIN_SUII_REC), K0_ODR_BUHIN_SUII, Len(K0_ODR_BUHIN_SUII), 0)
        
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                        MsgBox "データ作成中断！！ 再起動してください"
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "ODR_BUHIN_SUII")
                        Exit Function
                End Select
                    
                wkDouble = CDbl(StrConv(ODR_BUHIN_SUII_REC.SYUKO_QTY, vbUnicode))
                wkDouble = wkDouble + CDbl(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode)) * CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))
                    
                    
                Call UniCode_Conv(ODR_BUHIN_SUII_REC.SYUKO_QTY, Format(wkDouble, "00000.00"))


                Do
                
                    sts = BTRV(BtOpUpdate, ODR_BUHIN_SUII_POS, ODR_BUHIN_SUII_REC, Len(ODR_BUHIN_SUII_REC), K0_ODR_BUHIN_SUII, Len(K0_ODR_BUHIN_SUII), 0)
        
                    Select Case sts
                        Case BtNoErr
                                    
                            Exit Do
                                    
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                            yn = MsgBox("他で使用中です！<ODR_TEMP3>" & Chr(13) & Chr(10) & _
                                        "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                            If yn = vbNo Then
                                Exit Do
                            End If
            
                        Case Else
                            Call File_Error(sts, BtOpUpdate, "ODR_BUHIN_SUII")
                            Exit Function
                    End Select
                
                Loop




            End If





            com = BtOpGetNext
                        
    
    
        Loop
    
    
        
        com = BtOpGetNext
    
    
    
    Loop


'---------------------------------------    出庫予定OR 実績を設定（回答納期日）
    Call UniCode_Conv(K3_ODR_ORDER.KAITO_DT, Format(S_DATE, "YYYYMMDD"))

    com = BtOpGetGreaterEqual


    Do
        DoEvents
    
    
        sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K3_ODR_ORDER, Len(K3_ODR_ORDER), 3)

        Select Case sts
            Case BtNoErr

            
                If StrConv(ODR_ORDER_REC.KAITO_DT, vbUnicode) > Format(E_DATE, "YYYYMMDD") Then
                    Exit Do
                End If
            
            
            Case BtErrEOF
            
                Exit Do




            Case Else
                Call File_Error(sts, BtOpGetEqual, "ODR_ORDER")
                Exit Function

        End Select
    
    
        '構成マスタ読み込み＜子部品展開＞
        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode))
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "")
        Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
        
        com = BtOpGetGreaterEqual

        Do
            DoEvents
        
        
            sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    
            Select Case sts
                Case BtNoErr
    
                    If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode) Or _
                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode) Or _
                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode) Or _
                        StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode) <> StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode) Then
                        Exit Do
                    End If
                
                
                
                
                Case BtErrEOF
                
                    Exit Do
    
    
    
    
                Case Else
                    Call File_Error(sts, com, "P_COMPO")
                    Exit Function
    
            End Select
        
        
        
            If StrConv(P_COMPO_K_REC.SEQNO, vbUnicode) = "000" Then
            Else
                '子部品推移読み込み＜存在チェック＞
                Call UniCode_Conv(K0_ODR_BUHIN_SUII.SEL_DATE, StrConv(ODR_ORDER_REC.KAITO_DT, vbUnicode))
'2008.04.04                Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
'2008.04.04                Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                
                Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_JGYOBU, SHIZAI)      '2008.04.04
                Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_NAIGAI, NAIGAI_NAI)  '2008.04.04
                
                Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
        
                sts = BTRV(BtOpGetEqual, ODR_BUHIN_SUII_POS, ODR_BUHIN_SUII_REC, Len(ODR_BUHIN_SUII_REC), K0_ODR_BUHIN_SUII, Len(K0_ODR_BUHIN_SUII), 0)
        
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                        MsgBox "データ作成中断！！ 再起動してください"
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "ODR_BUHIN_SUII")
                        Exit Function
                End Select
                    
                wkDouble = CDbl(StrConv(ODR_BUHIN_SUII_REC.SYUKO_QTY, vbUnicode))
                wkDouble = wkDouble + CDbl(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode)) * CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))
                    
                    
                Call UniCode_Conv(ODR_BUHIN_SUII_REC.SYUKO_QTY, Format(wkDouble, "00000.00"))

                Do
                
                    sts = BTRV(BtOpUpdate, ODR_BUHIN_SUII_POS, ODR_BUHIN_SUII_REC, Len(ODR_BUHIN_SUII_REC), K0_ODR_BUHIN_SUII, Len(K0_ODR_BUHIN_SUII), 0)
        
                    Select Case sts
                        Case BtNoErr
                                    
                            Exit Do
                                    
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                            yn = MsgBox("他で使用中です！<ODR_TEMP3>" & Chr(13) & Chr(10) & _
                                        "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                            If yn = vbNo Then
                                Exit Do
                            End If
            
                        Case Else
                            Call File_Error(sts, BtOpUpdate, "ODR_BUHIN_SUII")
                            Exit Function
                    End Select
                
                Loop
    
    
            End If
            
            com = BtOpGetNext
        Loop

            


    
        
        com = BtOpGetNext
    
    
    
    Loop




'---------------------------------------    入庫予定OR 実績を設定
    Call UniCode_Conv(K3_ODR_ORDER.KAITO_DT, Format(S_DATE, "YYYYMMDD"))

    com = BtOpGetGreaterEqual


    Do
        DoEvents
    
    
        sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K3_ODR_ORDER, Len(K3_ODR_ORDER), 3)

        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
            
                Exit Do




            Case Else
                Call File_Error(sts, BtOpGetEqual, "ODR_ORDER")
                Exit Function

        End Select
    
    
        '構成マスタ読み込み＜子部品展開＞
        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode))
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "")
        Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
        
        com = BtOpGetGreaterEqual

        Do
            DoEvents
        
        
            sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    
            Select Case sts
                Case BtNoErr
    
                    If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode) Or _
                        StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode) Or _
                        StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode) Or _
                        StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode) <> StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode) Then
                        Exit Do
                    End If
                
                
                
                
                Case BtErrEOF
                
                    Exit Do
    
    
    
    
                Case Else
                    Call File_Error(sts, com, "P_COMPO")
                    Exit Function
    
            End Select
        
        
        
            If StrConv(P_COMPO_K_REC.SEQNO, vbUnicode) = "000" Then
            Else
                
                
                
                
'2008.04.04                Call UniCode_Conv(K1_ODR_BUHIN_ORDER.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
'2008.04.04                Call UniCode_Conv(K1_ODR_BUHIN_ORDER.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                Call UniCode_Conv(K1_ODR_BUHIN_ORDER.JGYOBU, SHIZAI)        '2008.04.04
                Call UniCode_Conv(K1_ODR_BUHIN_ORDER.NAIGAI, NAIGAI_NAI)    '2008.04.04
                
                
                Call UniCode_Conv(K1_ODR_BUHIN_ORDER.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                
                Call UniCode_Conv(K1_ODR_BUHIN_ORDER.SEL_DATE, "")
                Call UniCode_Conv(K1_ODR_BUHIN_ORDER.DATA_KBN, "")


                com = BtOpGetGreater
                
                Do
                
                    DoEvents
                    
                    sts = BTRV(com, ODR_BUHIN_ORDER_POS, ODR_BUHIN_ORDER_REC, Len(ODR_BUHIN_ORDER_REC), K1_ODR_BUHIN_ORDER, Len(K1_ODR_BUHIN_ORDER), 1)
            
                    Select Case sts
                        Case BtNoErr
'2008.04.04                            If StrConv(ODR_BUHIN_ORDER_REC.JGYOBU, vbUnicode) <> StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode) Or _
                                            StrConv(ODR_BUHIN_ORDER_REC.NAIGAI, vbUnicode) <> StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode) Or _
                                            StrConv(ODR_BUHIN_ORDER_REC.HIN_GAI, vbUnicode) <> StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode) Then
                        
                            If StrConv(ODR_BUHIN_ORDER_REC.JGYOBU, vbUnicode) <> SHIZAI Or _
                                StrConv(ODR_BUHIN_ORDER_REC.NAIGAI, vbUnicode) <> NAIGAI_NAI Or _
                                StrConv(ODR_BUHIN_ORDER_REC.HIN_GAI, vbUnicode) <> StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode) Then
                        
                        
                                Exit Do
                        
                            End If
                        
                        Case BtErrEOF
                        
                            Exit Do
            
                        Case Else
                            Call File_Error(sts, com, "ODR_BUHIN_ORDER")
                            Exit Function
            
                    End Select
                
                
                
                
                
                '子部品推移読み込み＜存在チェック＞
                    Call UniCode_Conv(K0_ODR_BUHIN_SUII.SEL_DATE, StrConv(ODR_BUHIN_ORDER_REC.SEL_DATE, vbUnicode))
                    Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_JGYOBU, StrConv(ODR_BUHIN_ORDER_REC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_NAIGAI, StrConv(ODR_BUHIN_ORDER_REC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_ODR_BUHIN_SUII.KO_HIN_GAI, StrConv(ODR_BUHIN_ORDER_REC.HIN_GAI, vbUnicode))
            
                    sts = BTRV(BtOpGetEqual, ODR_BUHIN_SUII_POS, ODR_BUHIN_SUII_REC, Len(ODR_BUHIN_SUII_REC), K0_ODR_BUHIN_SUII, Len(K0_ODR_BUHIN_SUII), 0)
            
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound
                            
                            MsgBox "データ作成中断！！ 再起動してください"
                            Exit Function
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "ODR_BUHIN_SUII")
                            Exit Function
                    End Select
                        
                    wkDouble = CDbl(StrConv(ODR_BUHIN_SUII_REC.NYUKO_QTY, vbUnicode))
                    wkDouble = wkDouble + CDbl(StrConv(ODR_BUHIN_ORDER_REC.NYUKO_QTY, vbUnicode))
                        
                        
                    Call UniCode_Conv(ODR_BUHIN_SUII_REC.NYUKO_QTY, Format(wkDouble, "00000.00"))
    
                    Do
                    
                        sts = BTRV(BtOpUpdate, ODR_BUHIN_SUII_POS, ODR_BUHIN_SUII_REC, Len(ODR_BUHIN_SUII_REC), K0_ODR_BUHIN_SUII, Len(K0_ODR_BUHIN_SUII), 0)
            
                        Select Case sts
                            Case BtNoErr
                                        
                                Exit Do
                                        
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                                yn = MsgBox("他で使用中です！<ODR_TEMP3>" & Chr(13) & Chr(10) & _
                                            "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                                If yn = vbNo Then
                                    Exit Do
                                End If
                
                            Case Else
                                Call File_Error(sts, BtOpUpdate, "ODR_BUHIN_SUII")
                                Exit Function
                        End Select
                    
                    Loop
    
                com = BtOpGetNext
    
    
                Loop
            End If
    
            
            com = BtOpGetNext
            
        Loop

            


    
        
        com = BtOpGetNext
    
    
    
    Loop





'---------------------------------------    在庫数／引当数を集計
    i = -1


    svJGYOBU = ""
    svNAIGAI = ""
    svHIN_GAI = ""

    ReDim tblY_ZAIKO_QTY(0 To Day_Max)
    ReDim tblHIKIATE_QTY(0 To Day_Max)
    ReDim tblNYUKO_QTY(0 To Day_Max)
    ReDim tblSYUKO_QTY(0 To Day_Max)


    com = BtOpGetFirst


    
    Do
        DoEvents
    
        sts = BTRV(com, ODR_BUHIN_SUII_POS, ODR_BUHIN_SUII_REC, Len(ODR_BUHIN_SUII_REC), K1_ODR_BUHIN_SUII, Len(K1_ODR_BUHIN_SUII), 1)

        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
            
                Exit Do

            Case Else
                Call File_Error(sts, com, "ODR_BUHIN_SUII")
                Exit Function

        End Select
    
    
    
        If Trim(svJGYOBU) = "" Then
            svJGYOBU = StrConv(ODR_BUHIN_SUII_REC.KO_JGYOBU, vbUnicode)
            svNAIGAI = StrConv(ODR_BUHIN_SUII_REC.KO_NAIGAI, vbUnicode)
            svHIN_GAI = StrConv(ODR_BUHIN_SUII_REC.KO_HIN_GAI, vbUnicode)
        
            For i = 0 To Day_Max
                tblY_ZAIKO_QTY(i) = 0
                tblHIKIATE_QTY(i) = 0
                tblNYUKO_QTY(i) = 0
                tblSYUKO_QTY(i) = 0
            Next i
        
        End If
    
    
    
    
    
        If svJGYOBU <> StrConv(ODR_BUHIN_SUII_REC.KO_JGYOBU, vbUnicode) Or _
            svNAIGAI <> StrConv(ODR_BUHIN_SUII_REC.KO_NAIGAI, vbUnicode) Or _
            svHIN_GAI <> StrConv(ODR_BUHIN_SUII_REC.KO_HIN_GAI, vbUnicode) Then
        
        
            If tmpZaiko_Suii_Put_Proc(S_DATE, E_DATE, svJGYOBU, svNAIGAI, svHIN_GAI, tblY_ZAIKO_QTY(), tblHIKIATE_QTY(), tblNYUKO_QTY(), tblSYUKO_QTY()) Then
               Unload Me
            End If
        
            svJGYOBU = StrConv(ODR_BUHIN_SUII_REC.KO_JGYOBU, vbUnicode)
            svNAIGAI = StrConv(ODR_BUHIN_SUII_REC.KO_NAIGAI, vbUnicode)
            svHIN_GAI = StrConv(ODR_BUHIN_SUII_REC.KO_HIN_GAI, vbUnicode)
        
        
            For i = 0 To Day_Max
                tblY_ZAIKO_QTY(i) = 0
                tblHIKIATE_QTY(i) = 0
                tblNYUKO_QTY(i) = 0
                tblSYUKO_QTY(i) = 0
            Next i
        
        End If
                        
                        
        i = DateDiff("d", S_DATE, Mid(StrConv(ODR_BUHIN_SUII_REC.SEL_DATE, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(ODR_BUHIN_SUII_REC.SEL_DATE, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(ODR_BUHIN_SUII_REC.SEL_DATE, vbUnicode), 7, 2))
                        
                        
        
        tblY_ZAIKO_QTY(i) = tblY_ZAIKO_QTY(i) + CDbl(StrConv(ODR_BUHIN_SUII_REC.Y_ZAIKO_QTY, vbUnicode))
        tblHIKIATE_QTY(i) = tblHIKIATE_QTY(i) + CDbl(StrConv(ODR_BUHIN_SUII_REC.HIKIATE_QTY, vbUnicode))
        tblNYUKO_QTY(i) = tblNYUKO_QTY(i) + CDbl(StrConv(ODR_BUHIN_SUII_REC.NYUKO_QTY, vbUnicode))
        tblSYUKO_QTY(i) = tblSYUKO_QTY(i) + CDbl(StrConv(ODR_BUHIN_SUII_REC.SYUKO_QTY, vbUnicode))
        
                
        
        
        
        
        
        
        
        com = BtOpGetNext
    
    
    Loop
    
    
    If Trim(svJGYOBU) <> "" Then
        If tmpZaiko_Suii_Put_Proc(S_DATE, E_DATE, svJGYOBU, svNAIGAI, svHIN_GAI, tblY_ZAIKO_QTY(), tblHIKIATE_QTY(), tblNYUKO_QTY(), tblSYUKO_QTY()) Then
           Unload Me
        End If
    End If
    
    
    




    tmpZaiko_Suii_Make_Proc = False






End Function



Private Function tmpBUHIN_ORDER_Make_Proc(S_DATE As String, E_DATE As String, Optional USE_YM As String = " ") As Integer
'----------------------------------------------------------------------------
'                   子部品注文データ作成
'----------------------------------------------------------------------------
Dim sts                 As Integer
Dim com                 As Integer


Dim wkDate              As String
Dim i                   As Integer

Dim yn                  As Integer

Dim wkDouble            As Double
Dim wkLONG              As Long


Dim Skip_F              As Boolean



    tmpBUHIN_ORDER_Make_Proc = True






'---------------------------------------    回答納期で集計
    Call UniCode_Conv(K6_P_SHORDER.ANS_NOUKI_DT, Format(S_DATE, "YYYYMMDD"))

    com = BtOpGetGreaterEqual


    Do
        DoEvents
    
    
        sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K6_P_SHORDER, Len(K6_P_SHORDER), 6)

        Skip_F = False

        Select Case sts
            Case BtNoErr

            
                If StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode) > Format(E_DATE, "YYYYMMDD") Then
                    Exit Do
                End If
            
            
                If Trim(USE_YM) <> "" Then
                    If StrConv(P_SHORDER_REC.USE_YM, vbUnicode) <> Left(Format(USE_YM & "/01", "YYYYMMDD"), 6) Then
                        Skip_F = True
                    End If
                End If
            
            
            
            
            Case BtErrEOF
            
                Exit Do

            Case Else
                Call File_Error(sts, BtOpGetEqual, "P_SHORDER")
                Exit Function

        End Select
    
    
            
        If Not Skip_F Then
        
            If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
            Else
                
                Call UniCode_Conv(K0_P_SHUKEIRE.ORDER_NO, StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode))
                Call UniCode_Conv(K0_P_SHUKEIRE.SEQNO, "")
                
                com = BtOpGetGreater
                
                wkLONG = 0
                
                Do
                    
                    DoEvents
                    
                    sts = BTRV(com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K0_P_SHUKEIRE, Len(K0_P_SHUKEIRE), 0)
            
                    Select Case sts
                        Case BtNoErr
            
                            If StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode) <> StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode) Then
                            
                                Exit Do
                            
                            End If
                        
                        Case BtErrEOF
                        
                            Exit Do
            
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "P_SHORDER")
                            Exit Function
            
                    End Select
                
                
                    Call UniCode_Conv(K0_ODR_BUHIN_ORDER.SEL_DATE, StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode))
                    Call UniCode_Conv(K0_ODR_BUHIN_ORDER.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ODR_BUHIN_ORDER.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_ODR_BUHIN_ORDER.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
        
                    
                    sts = BTRV(BtOpGetEqual, ODR_BUHIN_ORDER_POS, ODR_BUHIN_ORDER_REC, Len(ODR_BUHIN_ORDER_REC), K0_ODR_BUHIN_ORDER, Len(K0_ODR_BUHIN_ORDER), 0)
            
                    Select Case sts
                        Case BtNoErr
                        
                            com = BtOpUpdate
                        
                        
                        Case BtErrKeyNotFound
                        
                            com = BtOpInsert
            
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "ODR_BUHIN_ORDER")
                            Exit Function
            
                    End Select
                    
                    
                    If CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode)) < 0 Then
                    Else
                    
                    
                        If com = BtOpInsert Then
                        
                            Call UniCode_Conv(ODR_BUHIN_ORDER_REC.SEL_DATE, StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode))
                            Call UniCode_Conv(ODR_BUHIN_ORDER_REC.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
                            Call UniCode_Conv(ODR_BUHIN_ORDER_REC.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
                            Call UniCode_Conv(ODR_BUHIN_ORDER_REC.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
                        
                            Call UniCode_Conv(ODR_BUHIN_ORDER_REC.DATA_KBN, "2")
                            Call UniCode_Conv(ODR_BUHIN_ORDER_REC.USE_YM, StrConv(P_SHORDER_REC.USE_YM, vbUnicode))
                                                    
                            Call UniCode_Conv(ODR_BUHIN_ORDER_REC.NYUKO_QTY, "00000.00")
                        
                        
                        End If
                        wkDouble = wkDouble + CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode))
                        Call UniCode_Conv(ODR_BUHIN_ORDER_REC.NYUKO_QTY, Format(wkDouble, "00000.00"))
                        wkLONG = wkLONG + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode))
                    
                    
                        Do
                        
                        
                            sts = BTRV(com, ODR_BUHIN_ORDER_POS, ODR_BUHIN_ORDER_REC, Len(ODR_BUHIN_ORDER_REC), K0_ODR_BUHIN_ORDER, Len(K0_ODR_BUHIN_ORDER), 0)
                
                            Select Case sts
                                Case BtNoErr
                                            
                                    Exit Do
                                            
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                                    yn = MsgBox("他で使用中です！<ODR_BUHIN_ORDER>" & Chr(13) & Chr(10) & _
                                                "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                                    If yn = vbNo Then
                                        Exit Do
                                    End If
                    
                                Case Else
                                    Call File_Error(sts, BtOpUpdate, "ODR_BUHIN_ORDER")
                                    Exit Function
                            End Select
                        
                        
                        
                        
                        Loop
                    
                    End If
                
                
                    com = BtOpGetNext
                
                Loop
                
                
                If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = P_KAN_ON Or _
                    wkLONG >= CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) Then
                
                Else
                
                    wkLONG = CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) - wkLONG
            
                    Call UniCode_Conv(K0_ODR_BUHIN_ORDER.SEL_DATE, StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode))
                    Call UniCode_Conv(K0_ODR_BUHIN_ORDER.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ODR_BUHIN_ORDER.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_ODR_BUHIN_ORDER.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
        
                    
                    sts = BTRV(BtOpGetEqual, ODR_BUHIN_ORDER_POS, ODR_BUHIN_ORDER_REC, Len(ODR_BUHIN_ORDER_REC), K0_ODR_BUHIN_ORDER, Len(K0_ODR_BUHIN_ORDER), 0)
            
                    Select Case sts
                        Case BtNoErr
                        
                            com = BtOpUpdate
                        
                        
                        Case BtErrKeyNotFound
                        
                            com = BtOpInsert
            
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "ODR_BUHIN_ORDER")
                            Exit Function
            
                    End Select
                
                
                
                
                    If com = BtOpInsert Then
                    
                        Call UniCode_Conv(ODR_BUHIN_ORDER_REC.SEL_DATE, StrConv(P_SHORDER_REC.ANS_NOUKI_DT, vbUnicode))
                        Call UniCode_Conv(ODR_BUHIN_ORDER_REC.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
                        Call UniCode_Conv(ODR_BUHIN_ORDER_REC.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
                        Call UniCode_Conv(ODR_BUHIN_ORDER_REC.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
                    
                        Call UniCode_Conv(ODR_BUHIN_ORDER_REC.DATA_KBN, "1")
                        Call UniCode_Conv(ODR_BUHIN_ORDER_REC.USE_YM, StrConv(P_SHORDER_REC.USE_YM, vbUnicode))
                                                
                        Call UniCode_Conv(ODR_BUHIN_ORDER_REC.NYUKO_QTY, "00000.00")
                    
                    
                    End If
                    wkDouble = wkDouble + CDbl(wkLONG)
                    Call UniCode_Conv(ODR_BUHIN_ORDER_REC.NYUKO_QTY, Format(wkDouble, "00000.00"))
                
                
                    Do
                    
                    
                        sts = BTRV(com, ODR_BUHIN_ORDER_POS, ODR_BUHIN_ORDER_REC, Len(ODR_BUHIN_ORDER_REC), K0_ODR_BUHIN_ORDER, Len(K0_ODR_BUHIN_ORDER), 0)
            
                        Select Case sts
                            Case BtNoErr
                                        
                                Exit Do
                                        
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                                yn = MsgBox("他で使用中です！<ODR_BUHIN_ORDER>" & Chr(13) & Chr(10) & _
                                            "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                                If yn = vbNo Then
                                    Exit Do
                                End If
                
                            Case Else
                                Call File_Error(sts, BtOpUpdate, "ODR_BUHIN_ORDER")
                                Exit Function
                        End Select
                    
                    
                    
                    
                    Loop
                
                
                End If
                    
            End If
        
        
        End If
        
        com = BtOpGetNext
    Loop


    







    tmpBUHIN_ORDER_Make_Proc = False






End Function



Private Function tmpZaiko_Suii_Put_Proc(S_DATE As String _
                                        , E_DATE As String _
                                        , JGYOBU As String _
                                        , NAIGAI As String _
                                        , HIN_GAI As String _
                                        , tblY_ZAIKO_QTY() As Double _
                                        , tblHIKIATE_QTY() As Double _
                                        , tblNYUKO_QTY() As Double _
                                        , tblSYUKO_QTY() As Double) As Integer
'----------------------------------------------------------------------------
'                   在庫推移算出
'----------------------------------------------------------------------------
Dim i           As Integer

Dim Today_i     As Integer
Dim Max_Day     As Integer

Dim wkQTY       As Double
Dim wkDate      As String
Dim wkIN_QTY    As Double
Dim wkOUT_QTY   As Double

Dim sts         As Integer
Dim com         As Integer
    
Dim yn          As Integer

Dim Skip_F      As Boolean

Dim c           As String * 128
Dim sBuffer     As String * 255
Dim WsNo        As String
Dim Ret         As Integer
Dim FullPath    As String

    tmpZaiko_Suii_Put_Proc = True
    
    
    
    
'    sts = BTRV(BtOpClose, ODR_BUHIN_ORDER_POS, ODR_BUHIN_ORDER_REC, Len(ODR_BUHIN_ORDER_REC), K0_ODR_BUHIN_ORDER, Len(K0_ODR_BUHIN_ORDER), 0)
'    If sts Then
'        If sts <> BtErrNoOpen Then
'            Call File_Error(sts, BtOpClose, "ODR_BUHIN_ORDER")
'        End If
'    End If
'
'
'                                                        '子部品注文（一時）データ　フルパス取込み
'    sts = GetIni("FILE", ODR_BUHIN_ORDER_ID, "SYS", c)
'    If sts <> False Then
'        Call Log_Out(LOG_F, "SYS.INI [ODR_BUHIN_ORDER]読み込みエラー")
'        Exit Function
'    End If
'    FullPath = RTrim(c)
'
'    sBuffer = Space(255)
'    If GetComputerNameA(sBuffer, 255) <> 0 Then
'        WsNo = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
'    Else
'        WsNo = "???"
'    End If
'
'
'    Ret = InStr(1, Trim(c), ".") - 1
'    FullPath = Left(Trim(c), Ret) & WsNo & Right(Trim(c), Len(Trim(c)) - Ret)
'
'    On Error Resume Next
'    Kill (FullPath)
'    On Error GoTo 0
'
'
'    If ODR_BUHIN_ORDER_Open(BtOpenNomal) Then
'        Exit Function
'    End If
    
    
    Today_i = DateDiff("d", S_DATE, Format(Now, "YYYY/MM/DD"))
    Max_Day = DateDiff("d", S_DATE, E_DATE)
    
    '在庫計算     --------------------------------------------------------
    wkQTY = tblY_ZAIKO_QTY(Today_i)
    If Today_i > 0 Then
    
        For i = Today_i - 1 To 0 Step -1
            If i = 0 Then
                tblY_ZAIKO_QTY(i) = wkQTY + tblSYUKO_QTY(i) - tblNYUKO_QTY(i)
            Else
                
                wkQTY = wkQTY + tblSYUKO_QTY(i) - tblNYUKO_QTY(i)
                tblY_ZAIKO_QTY(i) = wkQTY
            End If
        Next i
    
    
    
    End If
    
    
    wkQTY = tblY_ZAIKO_QTY(Today_i)
    For i = Today_i + 1 To Max_Day
        If i = Max_Day Then
            tblY_ZAIKO_QTY(i) = wkQTY
        Else
            
            wkQTY = wkQTY - tblSYUKO_QTY(i - 1) + tblNYUKO_QTY(i - 1)
            tblY_ZAIKO_QTY(i) = wkQTY
        End If
    Next i
    
    
    '引当残計算     --------------------------------------------------------
    wkQTY = tblY_ZAIKO_QTY(0)
    For i = 0 To UBound(tblY_ZAIKO_QTY)
        '出庫（完了日）
        wkDate = DateAdd("d", i, S_DATE)
        Call UniCode_Conv(K4_ODR_ORDER.FIN_DT, wkDate)
        
        
        com = BtOpGetGreater
        
        
        wkIN_QTY = 0
        wkOUT_QTY = 0
        
        Do
            DoEvents
            
            Skip_F = False
            
            sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K4_ODR_ORDER, Len(K4_ODR_ORDER), 4)
            Select Case sts
                Case BtNoErr

            
                    If Trim(Text1(ptxUSE_YY).Text) <> "" Then
                        If Left(Format(Text1(ptxUSE_YY).Text & "01", "YYYYMMDD"), 4) <> StrConv(ODR_ORDER_REC.USE_YM, vbUnicode) Then
                            Skip_F = True
                        End If
                    
                    End If
            
            
            
                Case BtErrEOF
            
                    Exit Do




                Case Else
                    Call File_Error(sts, BtOpGetEqual, "ODR_ORDER")
                    Exit Function

            End Select
        
        
            If Not Skip_F Then
                '構成マスタ読み込み＜子部品展開＞
                Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode))
                Call UniCode_Conv(K0_P_COMPO.JGYOBU, StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_P_COMPO.NAIGAI, StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_P_COMPO.HIN_GAI, StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "")
                Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
                
                com = BtOpGetGreaterEqual
        
                Do
                    DoEvents
                
                
                    sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            
                    Select Case sts
                        Case BtNoErr
            
                            If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode) Or _
                                StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode) Or _
                                StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode) Or _
                                StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode) <> StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode) Then
                                Exit Do
                            End If
                        
                        
                        
                        
                        Case BtErrEOF
                        
                            Exit Do
            
            
            
            
                        Case Else
                            Call File_Error(sts, com, "ODR_ORDER")
                            Exit Function
            
                    End Select
                
                
                
                    If StrConv(P_COMPO_K_REC.SEQNO, vbUnicode) = "000" Then
                    Else
                            
'2008.04.04                        If JGYOBU <> StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode) Or _
'2008.04.04                            NAIGAI <> StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode) Or _
'2008.04.04                            HIN_GAI <> StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode) Then
                        
                        
                        If HIN_GAI <> StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode) Then
                        
                        
                        
                        
                        Else
                            wkOUT_QTY = wkOUT_QTY + CDbl(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode)) * CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))
                        End If
        
                    End If
                
                    com = BtOpGetNext
                
                Loop

            End If

            com = BtOpGetNext
    
        Loop
                
        '出庫（納期回答日）
        wkDate = DateAdd("d", i, S_DATE)
        Call UniCode_Conv(K3_ODR_ORDER.KAITO_DT, wkDate)
        com = BtOpGetGreater
        
        Do
            DoEvents
            
            Skip_F = False
            
            sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K3_ODR_ORDER, Len(K3_ODR_ORDER), 3)
            Select Case sts
                Case BtNoErr

            
                    If Trim(Text1(ptxUSE_YY).Text) <> "" Then
                        If Left(Format(Text1(ptxUSE_YY).Text & "01", "YYYYMMDD"), 4) <> StrConv(ODR_ORDER_REC.USE_YM, vbUnicode) Then
                            Skip_F = True
                        End If
                    
                    End If
            
            
            
                    If Trim(StrConv(ODR_ORDER_REC.FIN_DT, vbUnicode)) <> "" Then
                        Skip_F = True
                    End If
                        
            
                Case BtErrEOF
            
                    Exit Do




                Case Else
                    Call File_Error(sts, BtOpGetEqual, "ODR_ORDER")
                    Exit Function

            End Select
        
        
            If Not Skip_F Then
                '構成マスタ読み込み＜子部品展開＞
                Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode))
                Call UniCode_Conv(K0_P_COMPO.JGYOBU, StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_P_COMPO.NAIGAI, StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_P_COMPO.HIN_GAI, StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "")
                Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
                
                com = BtOpGetGreaterEqual
        
                Do
                    DoEvents
                
                
                    sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            
                    Select Case sts
                        Case BtNoErr
            
                            If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode) Or _
                                StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode) Or _
                                StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode) Or _
                                StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode) <> StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode) Then
                                Exit Do
                            End If
                        
                        
                        
                        
                        Case BtErrEOF
                        
                            Exit Do
            
            
            
            
                        Case Else
                            Call File_Error(sts, com, "ODR_ORDER")
                            Exit Function
            
                    End Select
                
                
                
                    
                    If Skip_F Then
                    Else
                        If StrConv(P_COMPO_K_REC.SEQNO, vbUnicode) = "000" Then
                        Else
                                
'2008.04.04                        If JGYOBU <> StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode) Or _
'2008.04.04                            NAIGAI <> StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode) Or _
'2008.04.04                            HIN_GAI <> StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode) Then
                        
                        
                            If HIN_GAI <> StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode) Then
                            Else
                                wkOUT_QTY = wkOUT_QTY + CDbl(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode)) * CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))
            
                            End If
            
                        End If
                
                    End If
                    
                    com = BtOpGetNext
                
                Loop

            End If

            com = BtOpGetNext
    
        Loop
                
                
                
                
'        If tmpBUHIN_ORDER_Make_Proc(S_DATE, "9999/12/31", Trim(Text1(ptxUSE_YY).Text)) Then
'            Exit Function
'        End If
'
'
'        '---------------------------------------    入庫予定OR 実績を設定
'        Call UniCode_Conv(K3_ODR_ORDER.KAITO_DT, Format(S_DATE, "YYYYMMDD"))
'
'        com = BtOpGetGreaterEqual
'
'
'        Do
'            DoEvents
'
'
'            sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K3_ODR_ORDER, Len(K3_ODR_ORDER), 3)
'
'            Skip_F = False
'
'            Select Case sts
'                Case BtNoErr
'
'                    If Trim(Text1(ptxUSE_YY).Text) <> "" Then
'                        If Left(Format(Text1(ptxUSE_YY).Text & "01", "YYYYMMDD"), 4) <> StrConv(ODR_ORDER_REC.USE_YM, vbUnicode) Then
'                            Skip_F = True
'                        End If
'
'                    End If
'
'                Case BtErrEOF
'
'                    Exit Do
'
'
'
'
'                Case Else
'                    Call File_Error(sts, BtOpGetEqual, "ODR_ORDER")
'                    Exit Function
'
'            End Select
'
'
'            '構成マスタ読み込み＜子部品展開＞
'            If Not Skip_F Then
'                Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode))
'                Call UniCode_Conv(K0_P_COMPO.JGYOBU, StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode))
'                Call UniCode_Conv(K0_P_COMPO.NAIGAI, StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode))
'                Call UniCode_Conv(K0_P_COMPO.HIN_GAI, StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode))
'                Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "")
'                Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
'
'                com = BtOpGetGreaterEqual
'
'                Do
'                    DoEvents
'
'
'                    sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
'
'                    Select Case sts
'                        Case BtNoErr
'
'                            If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> StrConv(ODR_ORDER_REC.SHIMUKE, vbUnicode) Or _
'                                StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> StrConv(ODR_ORDER_REC.JGYOBU, vbUnicode) Or _
'                                StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> StrConv(ODR_ORDER_REC.NAIGAI, vbUnicode) Or _
'                                StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode) <> StrConv(ODR_ORDER_REC.HIN_GAI, vbUnicode) Then
'                                Exit Do
'                            End If
'
'
'
'
'                        Case BtErrEOF
'
'                            Exit Do
'
'
'
'
'                        Case Else
'                            Call File_Error(sts, com, "P_COMPO")
'                            Exit Function
'
'                    End Select
'
'
'
'                    If StrConv(P_COMPO_K_REC.SEQNO, vbUnicode) = "000" Then
'                    Else
'
'                        Call UniCode_Conv(K1_ODR_BUHIN_ORDER.JGYOBU, SHIZAI)        '2008.04.04
'                        Call UniCode_Conv(K1_ODR_BUHIN_ORDER.NAIGAI, NAIGAI_NAI)    '2008.04.04
'
'
'                        Call UniCode_Conv(K1_ODR_BUHIN_ORDER.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
'
'                        Call UniCode_Conv(K1_ODR_BUHIN_ORDER.SEL_DATE, "")
'                        Call UniCode_Conv(K1_ODR_BUHIN_ORDER.DATA_KBN, "")
'
'
'                        com = BtOpGetGreater
'
'                        Do
'
'                            DoEvents
'
'                            sts = BTRV(com, ODR_BUHIN_ORDER_POS, ODR_BUHIN_ORDER_REC, Len(ODR_BUHIN_ORDER_REC), K1_ODR_BUHIN_ORDER, Len(K1_ODR_BUHIN_ORDER), 1)
'
'                            Skip_F = False
'
'                            Select Case sts
'                                Case BtNoErr
'
'                                    If StrConv(ODR_BUHIN_ORDER_REC.JGYOBU, vbUnicode) <> SHIZAI Or _
'                                        StrConv(ODR_BUHIN_ORDER_REC.NAIGAI, vbUnicode) <> NAIGAI_NAI Or _
'                                        StrConv(ODR_BUHIN_ORDER_REC.HIN_GAI, vbUnicode) <> StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode) Then
'
'
'                                        Exit Do
'
'                                    End If
'
'                                    If Trim(Text1(ptxUSE_YY).Text) <> "" Then
'                                        If Left(Format(Text1(ptxUSE_YY).Text & "01", "YYYYMMDD"), 4) <> StrConv(ODR_BUHIN_ORDER_REC.USE_YM, vbUnicode) Then
'                                            Skip_F = True
'                                        End If
'
'                                    End If
'
'
'
'                                Case BtErrEOF
'
'                                    Exit Do
'
'                                Case Else
'                                    Call File_Error(sts, com, "ODR_BUHIN_ORDER")
'                                    Exit Function
'
'                            End Select
'
'
'
'
'
'
'                            If Not Skip_F Then
'                                wkIN_QTY = wkIN_QTY + CDbl(StrConv(ODR_BUHIN_ORDER_REC.NYUKO_QTY, vbUnicode))
'
'
'
'                            End If
'
'                            com = BtOpGetNext
'
'
'                        Loop
'
'                    End If
'                    com = BtOpGetNext
'
'                Loop
'            End If
'
'            com = BtOpGetNext
'
'
'
'        Loop
    
    
'        tblHIKIATE_QTY(i) = tblY_ZAIKO_QTY(i) + wkIN_QTY - wkOUT_QTY
        tblHIKIATE_QTY(i) = tblY_ZAIKO_QTY(i) - wkOUT_QTY
    
    
    Next i
    
    
    
    
    
    
    
    
    
    
    'データ出力     --------------------------------------------------------
    Call UniCode_Conv(K1_wODR_BUHIN_SUII.KO_JGYOBU, JGYOBU)
    Call UniCode_Conv(K1_wODR_BUHIN_SUII.KO_NAIGAI, NAIGAI)
    Call UniCode_Conv(K1_wODR_BUHIN_SUII.KO_HIN_GAI, HIN_GAI)
        
            
    For i = 0 To Max_Day
    
    
        wkDate = DateAdd("d", i, S_DATE)
    
        Call UniCode_Conv(K1_wODR_BUHIN_SUII.SEL_DATE, Format(wkDate, "YYYYMMDD"))
    
    
        sts = BTRV(BtOpGetEqual, wODR_BUHIN_SUII_POS, wODR_BUHIN_SUII_REC, Len(wODR_BUHIN_SUII_REC), K1_wODR_BUHIN_SUII, Len(K1_wODR_BUHIN_SUII), 1)

        Select Case sts
            Case BtNoErr
                        

            Case BtErrKeyNotFound

            Case Else
                Call File_Error(sts, BtOpGetEqual, "wODR_BUHIN_SUII")
                Exit Function
        End Select
    
    
        If sts = BtNoErr Then
        
            If tblY_ZAIKO_QTY(i) >= 0 Then
                Call UniCode_Conv(wODR_BUHIN_SUII_REC.Y_ZAIKO_QTY, Format(tblY_ZAIKO_QTY(i), "00000.00"))
            Else
                Call UniCode_Conv(wODR_BUHIN_SUII_REC.Y_ZAIKO_QTY, Format(tblY_ZAIKO_QTY(i), "0000.00"))
            End If
        
            If tblHIKIATE_QTY(i) >= 0 Then
                Call UniCode_Conv(wODR_BUHIN_SUII_REC.HIKIATE_QTY, Format(tblHIKIATE_QTY(i), "00000.00"))
            Else
                Call UniCode_Conv(wODR_BUHIN_SUII_REC.HIKIATE_QTY, Format(tblHIKIATE_QTY(i), "0000.00"))
            End If
        
    
            Do
                sts = BTRV(BtOpUpdate, wODR_BUHIN_SUII_POS, wODR_BUHIN_SUII_REC, Len(wODR_BUHIN_SUII_REC), K1_wODR_BUHIN_SUII, Len(K1_wODR_BUHIN_SUII), 1)
        
                Select Case sts
                    Case BtNoErr
                                
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'レコード使用中
                        yn = MsgBox("他で使用中です！<ODR_BUHIN_SUII>" & Chr(13) & Chr(10) & _
                                    "　再試行しますか？", vbYesNo + vbExclamation, "確認入力")
                        If yn = vbNo Then
                            Exit Do
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "wODR_BUHIN_SUII")
                        Exit Function
                End Select
            Loop
    
        End If
    
    Next i
    
    
    
    tmpZaiko_Suii_Put_Proc = False
End Function


