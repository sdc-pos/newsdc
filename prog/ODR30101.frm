VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form ODR30101 
   BorderStyle     =   1  '�Œ�(����)
   ClientHeight    =   10140
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   15270
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   6240
      TabIndex        =   18
      Top             =   660
      Width           =   3060
      Begin VB.OptionButton Option1 
         Caption         =   "�e�i��"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   20
         Top             =   180
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�d���c"
         Height          =   375
         Index           =   0
         Left            =   105
         TabIndex        =   19
         Top             =   180
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���X�g"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   4950
      TabIndex        =   5
      Top             =   120
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3375
      TabIndex        =   4
      Top             =   120
      Width           =   1440
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Index           =   0
      Left            =   11025
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   15
      Top             =   75
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Index           =   1
      Left            =   12675
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   13
      Top             =   105
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�\ ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��]�[��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   8100
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�X �V"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   1440
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '��������
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   5040
      MaxLength       =   7
      TabIndex        =   1
      Top             =   780
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   0
      Top             =   780
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   9825
      ScaleHeight     =   195
      ScaleWidth      =   270
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�I�@��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   6525
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1440
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   8475
      Left            =   60
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1350
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   14949
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "�q���i"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "�q���i��"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "�g�p��"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "�K�v��"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "�����݌�"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "�s����"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "������"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "�d���c"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "�����i"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "�ݒ��}"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "���b�g"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "�񓚔[��"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "��]�[��"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "�d����"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "�d���於"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "�d���P��"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "�g�p��"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "���ƕ�"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "�����O"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "�[����"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "�[���於"
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "�݌ɐ�"
      Columns(21).DataField=   ""
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "�e��0"
      Columns(22).DataField=   ""
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "�d����"
      Columns(23).DataField=   ""
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   24
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=24"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1640"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2831"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2699"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(11)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1508"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1376"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=8194"
      Splits(0)._ColumnProps(17)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=1773"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1640"
      Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=8194"
      Splits(0)._ColumnProps(23)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=1773"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1640"
      Splits(0)._ColumnProps(28)=   "Column(4)._ColStyle=8194"
      Splits(0)._ColumnProps(29)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=1588"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1455"
      Splits(0)._ColumnProps(34)=   "Column(5)._ColStyle=8194"
      Splits(0)._ColumnProps(35)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=1588"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1455"
      Splits(0)._ColumnProps(40)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(42)=   "Column(7).Width=1588"
      Splits(0)._ColumnProps(43)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(7)._WidthInPix=1455"
      Splits(0)._ColumnProps(45)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(46)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(47)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(48)=   "Column(8).Width=1588"
      Splits(0)._ColumnProps(49)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(8)._WidthInPix=1455"
      Splits(0)._ColumnProps(51)=   "Column(8)._ColStyle=8194"
      Splits(0)._ColumnProps(52)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(53)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(54)=   "Column(9).Width=1588"
      Splits(0)._ColumnProps(55)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(9)._WidthInPix=1455"
      Splits(0)._ColumnProps(57)=   "Column(9)._ColStyle=8194"
      Splits(0)._ColumnProps(58)=   "Column(9).AllowFocus=0"
      Splits(0)._ColumnProps(59)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(60)=   "Column(10).Width=1588"
      Splits(0)._ColumnProps(61)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(10)._WidthInPix=1455"
      Splits(0)._ColumnProps(63)=   "Column(10)._ColStyle=8194"
      Splits(0)._ColumnProps(64)=   "Column(10).AllowFocus=0"
      Splits(0)._ColumnProps(65)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(66)=   "Column(11).Width=2143"
      Splits(0)._ColumnProps(67)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(11)._WidthInPix=2011"
      Splits(0)._ColumnProps(69)=   "Column(11)._ColStyle=8196"
      Splits(0)._ColumnProps(70)=   "Column(11).AllowFocus=0"
      Splits(0)._ColumnProps(71)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(72)=   "Column(12).Width=2302"
      Splits(0)._ColumnProps(73)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(74)=   "Column(12)._WidthInPix=2170"
      Splits(0)._ColumnProps(75)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(76)=   "Column(13).Width=1402"
      Splits(0)._ColumnProps(77)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(78)=   "Column(13)._WidthInPix=1270"
      Splits(0)._ColumnProps(79)=   "Column(13)._ColStyle=8196"
      Splits(0)._ColumnProps(80)=   "Column(13).AllowFocus=0"
      Splits(0)._ColumnProps(81)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(82)=   "Column(14).Width=2117"
      Splits(0)._ColumnProps(83)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(84)=   "Column(14)._WidthInPix=1984"
      Splits(0)._ColumnProps(85)=   "Column(14)._ColStyle=8196"
      Splits(0)._ColumnProps(86)=   "Column(14).AllowFocus=0"
      Splits(0)._ColumnProps(87)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(88)=   "Column(15).Width=1773"
      Splits(0)._ColumnProps(89)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(90)=   "Column(15)._WidthInPix=1640"
      Splits(0)._ColumnProps(91)=   "Column(15)._ColStyle=8194"
      Splits(0)._ColumnProps(92)=   "Column(15).AllowFocus=0"
      Splits(0)._ColumnProps(93)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(94)=   "Column(16).Width=873"
      Splits(0)._ColumnProps(95)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(96)=   "Column(16)._WidthInPix=741"
      Splits(0)._ColumnProps(97)=   "Column(16)._ColStyle=8196"
      Splits(0)._ColumnProps(98)=   "Column(16).Visible=0"
      Splits(0)._ColumnProps(99)=   "Column(16).AllowFocus=0"
      Splits(0)._ColumnProps(100)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(101)=   "Column(17).Width=529"
      Splits(0)._ColumnProps(102)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(103)=   "Column(17)._WidthInPix=397"
      Splits(0)._ColumnProps(104)=   "Column(17)._ColStyle=8196"
      Splits(0)._ColumnProps(105)=   "Column(17).Visible=0"
      Splits(0)._ColumnProps(106)=   "Column(17).AllowFocus=0"
      Splits(0)._ColumnProps(107)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(108)=   "Column(18).Width=529"
      Splits(0)._ColumnProps(109)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(110)=   "Column(18)._WidthInPix=397"
      Splits(0)._ColumnProps(111)=   "Column(18)._ColStyle=8196"
      Splits(0)._ColumnProps(112)=   "Column(18).Visible=0"
      Splits(0)._ColumnProps(113)=   "Column(18).AllowFocus=0"
      Splits(0)._ColumnProps(114)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(115)=   "Column(19).Width=2117"
      Splits(0)._ColumnProps(116)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(117)=   "Column(19)._WidthInPix=1984"
      Splits(0)._ColumnProps(118)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(119)=   "Column(20).Width=2117"
      Splits(0)._ColumnProps(120)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(121)=   "Column(20)._WidthInPix=1984"
      Splits(0)._ColumnProps(122)=   "Column(20).AllowFocus=0"
      Splits(0)._ColumnProps(123)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(124)=   "Column(21).Width=1588"
      Splits(0)._ColumnProps(125)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(126)=   "Column(21)._WidthInPix=1455"
      Splits(0)._ColumnProps(127)=   "Column(21)._ColStyle=2"
      Splits(0)._ColumnProps(128)=   "Column(21).AllowFocus=0"
      Splits(0)._ColumnProps(129)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(130)=   "Column(22).Width=1588"
      Splits(0)._ColumnProps(131)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(132)=   "Column(22)._WidthInPix=1455"
      Splits(0)._ColumnProps(133)=   "Column(22)._ColStyle=2"
      Splits(0)._ColumnProps(134)=   "Column(22).AllowFocus=0"
      Splits(0)._ColumnProps(135)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(136)=   "Column(23).Width=1588"
      Splits(0)._ColumnProps(137)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(138)=   "Column(23)._WidthInPix=1455"
      Splits(0)._ColumnProps(139)=   "Column(23)._ColStyle=2"
      Splits(0)._ColumnProps(140)=   "Column(23).AllowFocus=0"
      Splits(0)._ColumnProps(141)=   "Column(23).Order=24"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �o�S�V�b�N"
      PrintInfos(0).PageFooterFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �o�S�V�b�N"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "�q���i�@���v�E�������"
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      CellTipsWidth   =   0
      InsertMode      =   0   'False
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
      _StyleDefs(5)   =   ":id=0,.fontname=�l�r �o�S�V�b�N"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H40FF00&,.bold=0,.fontsize=1125"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=�l�r �S�V�b�N"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=�l�r �S�V�b�N"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=�l�r �S�V�b�N"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H80000005&"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&H80FF00&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFF80&"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=87,.parent=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=96,.parent=4,.namedParent=37,.bgcolor=&H80FF00&"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=89,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=91,.parent=7,.bgcolor=&H80000005&"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=94,.parent=9,.namedParent=39"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=95,.parent=10,.namedParent=40,.bgcolor=&H80FF00&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=102,.parent=87,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=99,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=100,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=101,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=110,.parent=87,.locked=-1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=107,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=108,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=109,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=114,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=111,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=112,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=113,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=118,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=115,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=116,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=117,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=55,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=52,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=53,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=54,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=126,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=123,.parent=88"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=124,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=125,.parent=91"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=21,.parent=87,.alignment=1,.bgcolor=&H80000005&"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=18,.parent=88"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=19,.parent=89"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=20,.parent=91"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=67,.parent=87,.alignment=1"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=64,.parent=88"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=65,.parent=89"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=66,.parent=91"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=25,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=22,.parent=88"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=23,.parent=89"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=24,.parent=91"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=47,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=44,.parent=88"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=45,.parent=89"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=46,.parent=91"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=17,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=14,.parent=88"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=15,.parent=89"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=16,.parent=91"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=103,.parent=87,.locked=-1"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=84,.parent=88"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=85,.parent=89"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=86,.parent=91"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=51,.parent=87,.bgcolor=&H80000005&"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=48,.parent=88"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=49,.parent=89"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=50,.parent=91"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=130,.parent=87,.alignment=3,.locked=-1"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=127,.parent=88"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=128,.parent=89"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=129,.parent=91"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=29,.parent=87,.alignment=3,.locked=-1"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=26,.parent=88"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=27,.parent=89"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=28,.parent=91"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=43,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=30,.parent=88"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=31,.parent=89"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=32,.parent=91"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=59,.parent=87,.locked=-1"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=56,.parent=88"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=57,.parent=89"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=58,.parent=91"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=63,.parent=87,.locked=-1"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=60,.parent=88"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=61,.parent=89"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=62,.parent=91"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=138,.parent=87,.locked=-1"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=135,.parent=88"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=136,.parent=89"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=137,.parent=91"
      _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=71,.parent=87,.bgcolor=&H80000005&"
      _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=68,.parent=88"
      _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=69,.parent=89"
      _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=70,.parent=91"
      _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=75,.parent=87"
      _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=72,.parent=88"
      _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=73,.parent=89"
      _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=74,.parent=91"
      _StyleDefs(120) =   "Splits(0).Columns(21).Style:id=79,.parent=87,.alignment=1"
      _StyleDefs(121) =   "Splits(0).Columns(21).HeadingStyle:id=76,.parent=88"
      _StyleDefs(122) =   "Splits(0).Columns(21).FooterStyle:id=77,.parent=89"
      _StyleDefs(123) =   "Splits(0).Columns(21).EditorStyle:id=78,.parent=91"
      _StyleDefs(124) =   "Splits(0).Columns(22).Style:id=83,.parent=87,.alignment=1"
      _StyleDefs(125) =   "Splits(0).Columns(22).HeadingStyle:id=80,.parent=88"
      _StyleDefs(126) =   "Splits(0).Columns(22).FooterStyle:id=81,.parent=89"
      _StyleDefs(127) =   "Splits(0).Columns(22).EditorStyle:id=82,.parent=91"
      _StyleDefs(128) =   "Splits(0).Columns(23).Style:id=119,.parent=87,.alignment=1"
      _StyleDefs(129) =   "Splits(0).Columns(23).HeadingStyle:id=104,.parent=88"
      _StyleDefs(130) =   "Splits(0).Columns(23).FooterStyle:id=105,.parent=89"
      _StyleDefs(131) =   "Splits(0).Columns(23).EditorStyle:id=106,.parent=91"
      _StyleDefs(132) =   "Named:id=33:Normal"
      _StyleDefs(133) =   ":id=33,.parent=0"
      _StyleDefs(134) =   "Named:id=34:Heading"
      _StyleDefs(135) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(136) =   ":id=34,.wraptext=-1"
      _StyleDefs(137) =   "Named:id=35:Footing"
      _StyleDefs(138) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(139) =   "Named:id=36:Selected"
      _StyleDefs(140) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(141) =   "Named:id=37:Caption"
      _StyleDefs(142) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(143) =   "Named:id=38:HighlightRow"
      _StyleDefs(144) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(145) =   "Named:id=39:EvenRow"
      _StyleDefs(146) =   ":id=39,.parent=33,.bgcolor=&H80FF80&"
      _StyleDefs(147) =   "Named:id=40:OddRow"
      _StyleDefs(148) =   ":id=40,.parent=33,.bgcolor=&H40FF00&"
      _StyleDefs(149) =   "Named:id=41:RecordSelector"
      _StyleDefs(150) =   ":id=41,.parent=34"
      _StyleDefs(151) =   "Named:id=42:FilterBar"
      _StyleDefs(152) =   ":id=42,.parent=33"
      _StyleDefs(153) =   "Named:id=13:LockItem"
      _StyleDefs(154) =   ":id=13,.parent=39"
   End
   Begin VB.Label Lab_Dsp 
      Alignment       =   1  '�E����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   2
      Left            =   10050
      TabIndex        =   21
      Top             =   600
      Width           =   5025
   End
   Begin VB.Label Lab_Dsp 
      Alignment       =   1  '�E����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   9540
      TabIndex        =   17
      Top             =   960
      Width           =   5475
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      Caption         =   "���ƕ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   10245
      TabIndex        =   16
      Top             =   135
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      Caption         =   "�d����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   11895
      TabIndex        =   14
      Top             =   165
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      Caption         =   "�g�p��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4260
      TabIndex        =   11
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Lab_Dsp 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1980
      TabIndex        =   10
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      Caption         =   "�S����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   420
      TabIndex        =   9
      Top             =   840
      Width           =   720
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "�����I��"
      Begin VB.Menu SHORI 
         Caption         =   "�\��"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "�ŐV"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "�X�V"
         Index           =   2
      End
      Begin VB.Menu SHORI 
         Caption         =   "��ʈ��"
         Index           =   3
      End
      Begin VB.Menu SHORI 
         Caption         =   "�I��"
         Index           =   4
      End
   End
End
Attribute VB_Name = "ODR30101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NAIGAI_CODE()   As String * 1
Private NAIGAI_NAME()   As String

'�R���{�p�Y��
Private Const pcmbJI = 0            '���ƕ�
Private Const pcmbSM = 1            '�d������

'�e�L�X�g�p�Y��
Private Const ptxTOP% = 0
Private Const ptxLAST% = 1

Private Const ptxTANTO_CD% = 0
Private Const ptxUSE_YY% = 1

'���x���p�Y��
Private Const plabTANTO_NM% = 0

'�R�}���h�{�^���p�Y��
Private Const FuncDIS% = 0      '�\��
Private Const FuncKIBOU% = 1    '��]�[��
Private Const FuncCOR% = 2      '�X�V

Private Const FuncORDER% = 3    '���������s
Private Const FuncLIST% = 4     '���X�g���s


Private Const FuncEND% = 5       '�I��

'ListBox�Y��


'�O���b�h�X�V�}�[�N
Dim Grid_Cor_M      As Integer
Dim Grid_Req_M      As Integer



Private Const Min_Row% = 1              '�ŏ��s��
'Private Max_Row As Long                '�ő�\���s��
Private Const Max_Row = 9999            '�ő�s��

Private Const Min_Col% = 0              '�ŏ���
Private Const Max_Col% = 23             '�ő��

'Private Const Col_DEL% = 0                  '�폜�}�[�N


Dim row         As Long                 '�Ώہ@�s

Dim Cor_Row     As Long                 '�J�����g�s

Dim Init_F_30101    As Integer

Private NOUNYU          As String * 5


'Private Const Last_Update_Day$ = "�������� [ODR3010] 2016.03.14 09:00"
Private Const Last_Update_Day$ = "�������� [ODR3010] 2016.05.06 09:00"

Private Function ERR_CHK(Index As Integer)
Dim sts         As Integer
Dim yn          As Integer

Dim W_STR       As String


    ERR_CHK = True
    
                        '���͕������`�F�b�N
    If LenB(StrConv(Text1(Index), vbFromUnicode)) > Text1(Index).MaxLength Then
        MsgBox "���͂������ڂ́i�����ӂ�G���[�j�ł��B", vbExclamation
        Exit Function
    End If
    
    Select Case Index
        Case ptxTANTO_CD%
            Lab_Dsp(plabTANTO_NM) = ""
            If Trim(Text1(Index)) = "" Then
                MsgBox "�S���҂��w�肵�ĉ������B", vbExclamation
                Exit Function
            End If
            
            If Trim(Text1(Index)) = "admin" Then
                Lab_Dsp(plabTANTO_NM) = "�Ǘ��Ҍ���"
            Else
                
                Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(Index))
                Do
                    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrKeyNotFound       '���R�[�h����
                            MsgBox "�S���ҁ@���o�^�I", vbExclamation
                            Exit Function
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                            Beep
                            yn = MsgBox("���Ŏg�p���ł��I<TANTO>" & Chr(13) & Chr(10) & _
                                        "�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                            If yn = vbNo Then Exit Function
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "TANTO")
                            Exit Function
                    End Select
                Loop
                Lab_Dsp(plabTANTO_NM) = Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
            
            End If
            
        Case ptxUSE_YY%
            If Trim(Text1(Index)) = "" Then
                MsgBox "�g�p�N�����w�肵�ĉ������B", vbExclamation
                Exit Function
            End If
            
            W_STR = Text1(ptxUSE_YY%) & "/01"
            
            If Not IsDate(W_STR) Then
                MsgBox "�g�p���G���[�I", vbExclamation
                Exit Function
            End If
            
            W_STR = Format(W_STR, "yyyy/mm/dd")
            Text1(ptxUSE_YY%) = Left(W_STR, 7)
            
            If Left(W_STR, 4) < "2005" Then
                MsgBox "�g�p���@���@2005�N�G���[�I", vbExclamation
                Exit Function
            End If
            If Left(W_STR, 4) > "2100" Then
                MsgBox "�g�p���@���@2100�N�G���[�I", vbExclamation
                Exit Function
            End If
            
            
    End Select
    
    
    ERR_CHK = False
End Function

Private Function Grid_Err_Chk(Index As Integer, W_Aft As String)
'       �O���b�h���͓��e�G���[�`�F�b�N
'
Dim sts         As Integer
Dim yn          As Integer
Dim W_STR       As String
Dim W_Shime     As String

    Grid_Err_Chk = True
    
    Select Case Index
        'Case Col_DEL%                   '�폜�}�[�N
        '    If ORDR_GRID(Cor_Row, Index) Then
        '        W_STR = Trim(ORDR_GRID(Cor_Row, Col_KEY))
        '        If W_STR <> "" Then
        '            MsgBox "�����ς݁��폜�s�I", vbExclamation
        '            ORDR_GRID(Cor_Row, Index) = False
        '            TDBGrid1.ReBind
        '            TDBGrid1.Update
        '            'TDBGrid1.MoveFirst
        '            TDBGrid1.ScrollBars = dbgAutomatic
        '            Exit Function
        '        End If
        '    End If
            
        Case Col_ORDR_QTY%              '��������
            If Trim(W_Aft) <> "" Then
                If Not IsNumeric(W_Aft) Then
                    MsgBox Cor_Row & "�s�ځ@�������ʁ@���l�G���[�I", vbExclamation
                    Exit Function
                End If
            
            Else
                ORDR_GRID(Cor_Row, Col_KIBOU_DT) = ""
            
                TDBGrid1.ReBind
                TDBGrid1.Update
                'TDBGrid1.MoveFirst
                TDBGrid1.ScrollBars = dbgAutomatic
                DoEvents
            
            
            End If
        Case Col_KIBOU_DT%              '��]�[��
            
            If Trim(ORDR_GRID(Cor_Row, Col_ORDR_QTY)) <> "" Then
                W_STR = Format(W_Aft, "yyyy/mm/dd")
                If IsDate(W_Aft) Then
                    'W_STR = Format(W_Aft, "yyyy/mm/dd")
                    ORDR_GRID(Cor_Row, Index) = W_STR
                    
                    W_Shime = Left(W_STR, 4) & Mid(W_STR, 6, 2) & Right(W_STR, 2)
                    
                    If W_Shime < GW_SHIMEBI Then
                        MsgBox Cor_Row & "�s�� ��]�[���@���t�G���[�I", vbExclamation
                        Exit Function
                    End If
                    
                    TDBGrid1.ReBind
                    TDBGrid1.Update
                    'TDBGrid1.MoveFirst
                    TDBGrid1.ScrollBars = dbgAutomatic
                    DoEvents
                    
                    If Trim(ORDR_GRID(Cor_Row, Col_ORDR_QTY)) = "" Then
                        MsgBox Cor_Row & "�s�� ��]�[���@�s�v�I", vbExclamation
                        Exit Function
                    End If
                Else
                
                    MsgBox Cor_Row & "�s�� ��]�[���@���t�G���[�I", vbExclamation
                    Exit Function
                End If
                
                
            End If
            
        Case Col_KAITO_DT%              '�񓚔[��
            'If Trim(W_Aft) <> "" Then
            '    If IsDate(W_Aft) Then
            '        W_Str = Format(W_Aft, "yyyy/mm/dd")
            '        ORDR_GRID(Cor_Row, Index) = W_Str
            '        TDBGrid1.ReBind
            '        TDBGrid1.Update
            '        'TDBGrid1.MoveFirst
            '        TDBGrid1.ScrollBars = dbgAutomatic
            '
            '    If Trim(ORDR_GRID(Cor_Row, Col_ORDR_QTY)) = "" Then
            '        MsgBox Cor_Row & "�s�� �񓚔[���@�s�v�I", vbExclamation
            '        Exit Function
            '    End If
            '    Else
            '        MsgBox Cor_Row & "�s�ځ@�񓚔[���@���t�G���[�I", vbExclamation
            '        Exit Function
            '    End If
            'End If
        
        Case Col_KEY                    'Key�@��
            
            
        Case Col_DELI_CD
        
            If Trim(ORDR_GRID(Cor_Row, Col_ORDR_QTY)) <> "" Then
            
                Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, ORDR_GRID(Cor_Row, Index))
                sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                Select Case sts
                    Case BtNoErr
                        ORDR_GRID(Cor_Row, Index + 1) = Trim(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode))
                    
                    
                        TDBGrid1.ReBind
                        TDBGrid1.Update
                        'TDBGrid1.MoveFirst
                        TDBGrid1.ScrollBars = dbgAutomatic
                    
                    
                    Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                        'MsgBox Cor_Row & "�s�� �[���斢�o�^�ł��I", vbExclamation
                        
                        ORDR_GRID(Cor_Row, Index + 1) = ""
                        TDBGrid1.ReBind
                        TDBGrid1.Update
                        'TDBGrid1.MoveFirst
                        TDBGrid1.ScrollBars = dbgAutomatic
                        
                        
                        'Exit Function
                                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "P_UKEHARAI")
                        Exit Function
                End Select
            
            End If
            
    End Select
    
    DoEvents
    

    Grid_Err_Chk = False

End Function

Private Function Data_Disp()
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer


Dim W_Key       As String

Dim W_QTY       As Double
Dim W_STR       As String

Dim cnt         As Integer

    Data_Disp = True
    
    row = Min_Row - 1
    Call Input_Lock                             '��ʍ��ڃ��b�N
    
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�q���i�@�����������@�������I�@��Data_Disp��", Me.hwnd, 0)
    DoEvents
    
    If ODR_KENTO_Open(BtOpenNomal) Then
        MsgBox "�����𒆒f���܂��B", vbExclamation
        Call Input_UnLock                       '��ʍ��ڃA�����b�N 2016.05.06
        Exit Function
    End If
    
    sts = BTRV(BtOpGetFirst, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
    Select Case sts
        Case BtNoErr
                
        Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                
        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
            yn = MsgBox("���Ŏg�p���ł��I<ODR_KENTO>" & Chr(13) & Chr(10) & _
                            "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
            If yn = vbNo Then GoTo Err_exit
        Case Else
            Call File_Error(sts, BtOpGetFirst, "ODR_KENTO")
            GoTo Err_exit
    End Select
    If sts = BtNoErr Then
        W_STR = Trim(StrConv(ODR_KNT_R.ITEM_NM, vbUnicode)) & " ���݂̏��"
    Else
        W_STR = ""
    End If
    Lab_Dsp(1) = W_STR
    
    
    Set ORDR_GRID = Nothing
    
    W_Key = Left(Text1(ptxUSE_YY), 4) & Right(Text1(ptxUSE_YY), 2)
    Call UniCode_Conv(K0_ODR_KENTO.USE_YM, W_Key)
    Call UniCode_Conv(K0_ODR_KENTO.KO_JGYOBU, "")
    Call UniCode_Conv(K0_ODR_KENTO.KO_NAIGAI, "")
    Call UniCode_Conv(K0_ODR_KENTO.KO_HIN_GAI, "")
    com = BtOpGetGreaterEqual
    Do
        sts = BTRV(com, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                
            Case Else
                Call File_Error(sts, com, "ODR_KENTO")
                GoTo Err_exit
        End Select
        If sts <> BtNoErr Then Exit Do
        
        If StrConv(ODR_KNT_R.USE_YM, vbUnicode) <> W_Key Then Exit Do
        
        
            '   2008/11/15  �Ǘ����R�[�h�i�i�����󔒁j�̕\�����Ȃ��I
        If Trim(StrConv(ODR_KNT_R.KO_HIN_GAI, vbUnicode)) <> "" Then
        
        
        
            '   �K�v�����O�̂ݕ\���ΏہI
        
            If CDbl(Trim(StrConv(ODR_KNT_R.NED_QTY, vbUnicode))) <> 0 Or _
                Text1(ptxTANTO_CD) = "admin" Then
            
                
                If Trim(StrConv(ODR_KNT_R.KO_HIN_GAI, vbUnicode)) = "B533" Then
                    W_STR = ""
                End If
            
                '�q���i�R�[�h
                DIS_ITEM = Trim(StrConv(ODR_KNT_R.KO_HIN_GAI, vbUnicode))
                '�q���i��
                DIS_ITEM_NM = Trim(StrConv(ODR_KNT_R.ITEM_NM, vbUnicode))
                'DIS_ITEM_NM = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                '�g�p����
                DIS_USE_QTY = Format(CDbl(Trim(StrConv(ODR_KNT_R.USE_QTY, vbUnicode))), "##,##0.00")
                '�K�v��
                DIS_MRP_QTY = Format(CDbl(Trim(StrConv(ODR_KNT_R.NED_QTY, vbUnicode))), "##,##0.00")
                '�����݌�
                DIS_ZAI_QTY = Format(CDbl(Trim(StrConv(ODR_KNT_R.ZAI_QTY, vbUnicode))), "##,##0.00")
                '�s����
                DIS_FUSOKU = Format(CDbl(Trim(StrConv(ODR_KNT_R.MAI_QTY, vbUnicode))), "##,##0.00")
                    
                '������
                'If W_Qty <= 0 Then
                '    W_Qty = CDbl(Trim(StrConv(ODR_KNT_R.MAI_QTY, vbUnicode)))
                'End If
                DIS_ORDR_QTY = Format(CDbl(Trim(StrConv(ODR_KNT_R.ODR_QTY, vbUnicode))), "##,###0.00")
                '�d���c
                W_QTY = CDbl(Trim(StrConv(ODR_KNT_R.SHI_QTY1, vbUnicode)))
                W_QTY = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.SHI_QTY2, vbUnicode)))
                W_QTY = W_QTY + CDbl(Trim(StrConv(ODR_KNT_R.SHI_QTY3, vbUnicode)))
                DIS_ZAN_QTY = Format(W_QTY, "##,###0.00")
        
            
                '�����i
                DIS_HANSEIHIN_QTY = Format(CDbl(Trim(StrConv(ODR_KNT_R.HANSEIHIN_QTY, vbUnicode))), "##,###0.00")
                
                '�ݒ��}
                DIS_TEI_QTY = Format(CDbl(Trim(StrConv(ODR_KNT_R.ZAITEI_QTY, vbUnicode))), "##,###0.00")
                
                '���b�g��
                DIS_LOT_QTY = Format(CDbl(Trim(StrConv(ODR_KNT_R.LOT_QTY, vbUnicode))), "##,###0.00")
                
                '�d����
                DIS_SECT_CD = Trim(StrConv(ODR_KNT_R.SECT, vbUnicode))
                '�d���於
                DIS_SECT_NM = ""
        
                If DIS_SECT_CD <> "" Then
                    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, DIS_SECT_CD)
                    yn = 0
                    Do
                        sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                                Sleep (500)
                                yn = yn + 1
                                If yn >= 500 Then
                                    yn = MsgBox("���Ŏg�p���ł��I<�󕥐�}�X�^>" & Chr(13) & Chr(10) & _
                                                "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                                            
                                    If yn = vbNo Then GoTo Err_exit
                                End If
                                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "P_UKEHARAI")
                                GoTo Err_exit
                        End Select
                    Loop
                    If sts = BtNoErr Then
                        DIS_SECT_NM = Trim(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode))
                    End If
                
                End If
            
                '�d����P��
                If IsNumeric(StrConv(ODR_KNT_R.TANKA, vbUnicode)) Then
                    DIS_TANKA = Format(CDbl(Trim(StrConv(ODR_KNT_R.TANKA, vbUnicode))), "##,##0.00")
                Else
                    DIS_TANKA = "0.00"
                End If
        
                '��]�[��
                W_STR = Trim(StrConv(ODR_KNT_R.NOUKI, vbUnicode))
                If W_STR <> "" Then
                    W_STR = Left(W_STR, 4) & "/" & Mid(W_STR, 5, 2) & "/" & Right(W_STR, 2)
                End If
                DIS_KIBOU_DT = W_STR
        
               '�񓚔[��
                W_STR = Trim(StrConv(ODR_KNT_R.KAITO, vbUnicode))
                If W_STR <> "" Then
                    W_STR = Mid(W_STR, 3, 2) & "/" & Mid(W_STR, 5, 2) & "/" & Right(W_STR, 2)
                End If
            
                If CInt(StrConv(ODR_KNT_R.ZAN_CNT, vbUnicode)) > 1 Then
                    W_STR = W_STR & "*"
                End If
               
                DIS_KAITO_DE = W_STR
        
                '�[����
                DIS_DELI_CD = NOUNYU
                    
                Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, NOUNYU)
                sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                                    
                        Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
                                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "P_UKEHARAI")
                        GoTo Err_exit
                End Select
                DIS_DELI_NM = Trim(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode))
            
            
            
                DIS_Item_Zaiko = Format(CDbl(Trim(StrConv(ODR_KNT_R.ITEM_Z_QTY, vbUnicode))), "##,##0.00")
                
                
                DIS_ZAIKO_ODR = Format(CDbl(Trim(StrConv(ODR_KNT_R.MINASHI1, vbUnicode))), "##,##0.00")
                DIS_ZAIKO_UKE = Format(CDbl(Trim(StrConv(ODR_KNT_R.MINASHI2, vbUnicode))), "##,##0.00")
            
        
                '�j�d�x����
                DIS_KEY = Trim(StrConv(ODR_KNT_R.USE_YM, vbUnicode))
            
                '���ƕ�
                Key_JIGYOBU = Trim(StrConv(ODR_KNT_R.KO_JGYOBU, vbUnicode))
                '�����O
                Key_NAIGAI = Trim(StrConv(ODR_KNT_R.KO_NAIGAI, vbUnicode))
                
                row = row + 1
                If row > Max_Row Then
                    MsgBox "�ő�\���s���𒴂��܂����B"
                    Exit Do
                End If
                            
                If Grid_Set_Proc() Then
                    GoTo Err_exit
                End If
        
            End If
        
        
        End If
        
        
        com = BtOpGetNext
    Loop
    If row > 1 Then
        ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                            Col_FUSOKU%, XORDER_ASCEND, XTYPE_DOUBLE, _
                            Col_ORDR_QTY%, XORDER_DESCEND, XTYPE_DOUBLE, _
                            Col_ITEM%, XORDER_ASCEND, XTYPE_STRING
    End If
    
    Set TDBGrid1.Array = ORDR_GRID
    
    'TDBGrid1.style.Locked = True
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�q���i�@�����������@�\�����܂����B�@��Data_Disp��", Me.hwnd, 0)
    DoEvents
    Data_Disp = False
    
Err_exit:
    Call Input_UnLock                             '��ʍ��ڃ��b�N
    sts = BTRV(BtOpClose, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_KENTO")
        End If
    End If
    
    
End Function

Private Function Grid_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   �O���b�h�\���i�ړ����f�[�^���e�j
'               Row   �s��
'               mode�@FALSE:����OFF  TRUE:����ON
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim W_Row       As Long
Dim W_STR       As String
Dim W_QTY       As Double

    Grid_Set_Proc = True
    W_Row = row - 1

    ORDR_GRID.ReDim Min_Row, row, Min_Col, Max_Col
    
    'ORDR_GRID(row, Col_No) = Trim(row)                              '�s��
    
    If Right(DIS_USE_QTY, 3) = ".00" Then
        DIS_USE_QTY = Left(Trim(DIS_USE_QTY), Len(Trim(DIS_USE_QTY)) - 3)
    End If
    If Right(DIS_MRP_QTY, 3) = ".00" Then
        DIS_MRP_QTY = Left(Trim(DIS_MRP_QTY), Len(Trim(DIS_MRP_QTY)) - 3)
    End If
    If Right(DIS_ZAI_QTY, 3) = ".00" Then
        DIS_ZAI_QTY = Left(Trim(DIS_ZAI_QTY), Len(Trim(DIS_ZAI_QTY)) - 3)
    End If
    If Right(DIS_FUSOKU, 3) = ".00" Then
        DIS_FUSOKU = Left(Trim(DIS_FUSOKU), Len(Trim(DIS_FUSOKU)) - 3)
    End If
    If Right(DIS_ORDR_QTY, 3) = ".00" Then
        DIS_ORDR_QTY = Left(Trim(DIS_ORDR_QTY), Len(Trim(DIS_ORDR_QTY)) - 3)
    End If
    If Right(DIS_ZAN_QTY, 3) = ".00" Then
        DIS_ZAN_QTY = Left(Trim(DIS_ZAN_QTY), Len(Trim(DIS_ZAN_QTY)) - 3)
    End If
    
    If Right(DIS_Item_Zaiko, 3) = ".00" Then
        DIS_Item_Zaiko = Left(Trim(DIS_Item_Zaiko), Len(Trim(DIS_Item_Zaiko)) - 3)
    End If
    
    If Right(DIS_ZAIKO_ODR, 3) = ".00" Then
        DIS_ZAIKO_ODR = Left(Trim(DIS_ZAIKO_ODR), Len(Trim(DIS_ZAIKO_ODR)) - 3)
    End If
    If CDbl(Trim(DIS_ZAIKO_ODR)) = 0 Then DIS_ZAIKO_ODR = ""
    
    If Right(DIS_ZAIKO_UKE, 3) = ".00" Then
        DIS_ZAIKO_UKE = Left(Trim(DIS_ZAIKO_UKE), Len(Trim(DIS_ZAIKO_UKE)) - 3)
    End If
    If CDbl(Trim(DIS_ZAIKO_UKE)) = 0 Then DIS_ZAIKO_UKE = ""
    
    
    If Right(DIS_HANSEIHIN_QTY, 3) = ".00" Then
        DIS_HANSEIHIN_QTY = Left(Trim(DIS_HANSEIHIN_QTY), Len(Trim(DIS_HANSEIHIN_QTY)) - 3)
    End If
    
    If Right(DIS_TEI_QTY, 3) = ".00" Then
        DIS_TEI_QTY = Left(Trim(DIS_TEI_QTY), Len(Trim(DIS_TEI_QTY)) - 3)
    End If
    
    
    If Right(DIS_LOT_QTY, 3) = ".00" Then
        DIS_LOT_QTY = Left(Trim(DIS_LOT_QTY), Len(Trim(DIS_LOT_QTY)) - 3)
    End If
    
    
    
    ORDR_GRID(row, Col_ITEM) = Trim(DIS_ITEM)               '�q���i�R�[�h
    ORDR_GRID(row, Col_ITEM_NM) = Trim(DIS_ITEM_NM)         '�q���i��
    If CDbl(Trim(DIS_USE_QTY)) = 0 Then
        ORDR_GRID(row, Col_USE_QTY) = ""
    Else
        ORDR_GRID(row, Col_USE_QTY) = Trim(DIS_USE_QTY)     '�g�p����
    End If
    
    ORDR_GRID(row, Col_MRP_QTY) = Trim(DIS_MRP_QTY)         '�K�v��
    ORDR_GRID(row, Col_ZAI_QTY) = Trim(DIS_ZAI_QTY)         '�����݌�
    ORDR_GRID(row, Col_FUSOKU) = Trim(DIS_FUSOKU)           '�s����
    
    If CDbl(Trim(DIS_ORDR_QTY)) = 0 Then
        ORDR_GRID(row, Col_ORDR_QTY) = ""
    Else
        ORDR_GRID(row, Col_ORDR_QTY) = Trim(DIS_ORDR_QTY)   '������
    End If
    
    
    ORDR_GRID(row, Col_ZAN_QTY) = Trim(DIS_ZAN_QTY)         '�d���c
    
                                                            '�����i
    ORDR_GRID(row, Col_HANSEIHIN_QTY) = Trim(DIS_HANSEIHIN_QTY)
    
                                                            '�ݒ��}
    ORDR_GRID(row, Col_TEI_QTY) = Trim(DIS_TEI_QTY)
    
    ORDR_GRID(row, Col_LOT_QTY) = Trim(DIS_LOT_QTY)         '���b�g��
    ORDR_GRID(row, Col_SECT_CD) = Trim(DIS_SECT_CD)         '�d����
    ORDR_GRID(row, Col_SECT_NM) = Trim(DIS_SECT_NM)         '�d���於
    ORDR_GRID(row, Col_TANKA) = Trim(DIS_TANKA)             '�d���P��
    ORDR_GRID(row, Col_KIBOU_DT) = Trim(DIS_KIBOU_DT)       '��]�[��
    ORDR_GRID(row, Col_KAITO_DT) = Trim(DIS_KAITO_DE)       '�񓚔[��
    ORDR_GRID(row, Col_KEY) = DIS_KEY                       '�j�d�x����
    ORDR_GRID(row, Col_JIGYOBU) = Key_JIGYOBU               '���ƕ�
    ORDR_GRID(row, Col_NAIGAI) = Key_NAIGAI                 '���O
  
    ORDR_GRID(row, Col_DELI_CD) = Trim(DIS_DELI_CD)         '�[����
    ORDR_GRID(row, Col_DELI_NM) = Trim(DIS_DELI_NM)         '�[���於
  
    ORDR_GRID(row, Col_Item_Zaiko) = Trim(DIS_Item_Zaiko)   '�O�����݌�
    ORDR_GRID(row, Col_ZAIKO_ODR) = Trim(DIS_ZAIKO_ODR)       '�݌Ɂ{������
    ORDR_GRID(row, Col_ZAIKO_UKE) = Trim(DIS_ZAIKO_UKE)       '�݌Ɂ{������
  
    Grid_Set_Proc = False

End Function


Private Function KENTO_UPDT(QTY_0 As Integer) As Integer
                        
                        '   ����    QTY_0 = True �� 0�N���A�I
                        
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim X_i         As Integer

Dim W_Key       As String
Dim W_No        As String
Dim W_STR       As String
Dim W_Date      As String

Dim W_Moto      As Double
Dim W_QTY       As Double

Dim W_Test      As String

    KENTO_UPDT = True
    
    Call Input_Lock
    
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i�@�������@�X�V���I�@��KENTO_UPDT��", Me.hwnd, 0)
    
    If ODR_KENTO_Open(BtOpenExec) Then
        MsgBox "�����𒆒f���܂��B", vbExclamation
        Call Input_UnLock           '2016.05.06
        Exit Function
    End If
    
    
    X_i = ORDR_GRID.UpperBound(1)
    
    For Cor_Row = Min_Row To ORDR_GRID.UpperBound(1)
        If Trim(ORDR_GRID(Cor_Row, Col_ORDR_QTY)) = "" Then
            W_QTY = 0
        Else
            W_QTY = CDbl(Trim(ORDR_GRID(Cor_Row, Col_ORDR_QTY)))
        End If
        
        If W_QTY > 0 Then           '�������ʁ��O�̂ݑΏہI
        
            DIS_KEY = Trim(ORDR_GRID(Cor_Row, Col_KEY))
            Key_JIGYOBU = Trim(ORDR_GRID(Cor_Row, Col_JIGYOBU))
            Key_NAIGAI = Trim(ORDR_GRID(Cor_Row, Col_NAIGAI))
            Key_HinGai = Trim(ORDR_GRID(Cor_Row, Col_ITEM))
            
            Call UniCode_Conv(K0_ODR_KENTO.USE_YM, DIS_KEY)
            Call UniCode_Conv(K0_ODR_KENTO.KO_JGYOBU, Key_JIGYOBU)
            Call UniCode_Conv(K0_ODR_KENTO.KO_NAIGAI, Key_NAIGAI)
            Call UniCode_Conv(K0_ODR_KENTO.KO_HIN_GAI, Key_HinGai)
        
            Do
                sts = BTRV(BtOpGetEqual, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                        
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                        yn = MsgBox("���Ŏg�p���ł��I<ODR_KENTO>" & Chr(13) & Chr(10) & _
                                    "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                        If yn = vbNo Then Exit Do
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "ODR_KENTO")
                        Exit Do
                End Select
            Loop
            
            If sts = BtNoErr Then
                
                '�ǉ����d���c�ɉ��Z
                W_Moto = CDbl(StrConv(ODR_KNT_R.SHI_QTY3, vbUnicode))
                W_Moto = W_Moto + W_QTY
                Call UniCode_Conv(ODR_KNT_R.SHI_QTY3, CStr(W_Moto))
                
                
                If QTY_0 = True Then
                    '���������[���ɁI
                    Call UniCode_Conv(ODR_KNT_R.ODR_QTY, "00000000.00")
                Else
                    '�������ɉ�ʍ��ڒl���Z�b�g     '08/09/18
                    W_STR = CStr(W_QTY)
                    Call UniCode_Conv(ODR_KNT_R.ODR_QTY, W_STR)
                End If
                               
                               
                If Trim(ORDR_GRID(Cor_Row, Col_KIBOU_DT)) <> "" Then
                    Call UniCode_Conv(ODR_KNT_R.NOUKI, _
                        Format(Trim(ORDR_GRID(Cor_Row, Col_KIBOU_DT)), "yyyymmdd"))
                End If
                
                'If Trim(ORDR_GRID(Cor_Row, Col_KAITO_DT)) <> "" Then
                '    Call UniCode_Conv(ODR_KNT_R.KAITO, _
                '        Format(Trim(ORDR_GRID(Cor_Row, Col_KAITO_DT)), "yyyymmdd"))
                'End If
                
                '�[����
                Call UniCode_Conv(ODR_KNT_R.NONYU, Trim(ORDR_GRID(Cor_Row, Col_DELI_CD)))
                
                                '�󕥐�Ͻ��Ǎ���        '2009/04/03
                                                        '�}�X�^�[�`�F�b�N���āA�X�V���Ȃ��I
                                                        '
                Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, ORDR_GRID(Cor_Row, Col_SECT_CD))
                sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                
                If sts = BtNoErr Then
                    Do
                        sts = BTRV(BtOpUpdate, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
                        Select Case sts
                            Case BtNoErr
                                
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                                Sleep (500)
                            Case Else
                                Call File_Error(sts, BtOpUpdate, "ODR_KENTO")
                                GoTo Err_exit
                        End Select
                    Loop
                Else
                    W_Test = Trim(ORDR_GRID(Cor_Row, Col_SECT_CD))
                End If
                
''                If P_SHORDER_PUT(W_QTY) Then
''                    MsgBox "���ޔ����f�[�^�@�o�̓G���[�I", vbExclamation
''                    GoTo Err_Exit
''                End If
                
                
            End If
            
        
        
        End If
        
    Next Cor_Row
    
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i�������@�X�V�I���B�@��KENTO_UPDT��", Me.hwnd, 0)
 
    KENTO_UPDT = False
    
Err_exit:
        
    Call Input_UnLock
    
    
    sts = BTRV(BtOpClose, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_KENTO")
        End If
    End If
    
End Function

Private Function KENTO_UPDT2()
                        
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim X_i         As Integer

Dim W_Key       As String
Dim W_No        As String
Dim W_STR       As String
Dim W_Date      As String

Dim W_Moto      As Double
Dim W_QTY       As Double

Dim W_Test      As String

    KENTO_UPDT2 = True
    
    Call Input_Lock
    
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i�@�������@�X�V���I�@��KENTO_UPDT2��", Me.hwnd, 0)
    
    If ODR_KENTO_Open(BtOpenExec) Then
        MsgBox "�����𒆒f���܂��B", vbExclamation
        Call Input_UnLock       '2016.05.06
        Exit Function
    End If
    
    
    X_i = ORDR_GRID.UpperBound(1)
    
    For Cor_Row = Min_Row To ORDR_GRID.UpperBound(1)
        If Trim(ORDR_GRID(Cor_Row, Col_ORDR_QTY)) = "" Then
            W_QTY = 0
        Else
            W_QTY = CDbl(Trim(ORDR_GRID(Cor_Row, Col_ORDR_QTY)))
        End If
        
        'If W_QTY > 0 Then           '��������>�O�̂ݑΏہI     '2010/01/19 �������߂��I
        
            DIS_KEY = Trim(ORDR_GRID(Cor_Row, Col_KEY))
            Key_JIGYOBU = Trim(ORDR_GRID(Cor_Row, Col_JIGYOBU))
            Key_NAIGAI = Trim(ORDR_GRID(Cor_Row, Col_NAIGAI))
            Key_HinGai = Trim(ORDR_GRID(Cor_Row, Col_ITEM))
            
            Call UniCode_Conv(K0_ODR_KENTO.USE_YM, DIS_KEY)
            Call UniCode_Conv(K0_ODR_KENTO.KO_JGYOBU, Key_JIGYOBU)
            Call UniCode_Conv(K0_ODR_KENTO.KO_NAIGAI, Key_NAIGAI)
            Call UniCode_Conv(K0_ODR_KENTO.KO_HIN_GAI, Key_HinGai)
        
            Do
                sts = BTRV(BtOpGetEqual, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                        
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                        yn = MsgBox("���Ŏg�p���ł��I<ODR_KENTO>" & Chr(13) & Chr(10) & _
                                    "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                        If yn = vbNo Then Exit Do
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "ODR_KENTO")
                        Exit Do
                End Select
            Loop
            
            If sts = BtNoErr Then
                    '�������ɉ�ʍ��ڒl���Z�b�g     '08/09/18
                    W_STR = CStr(W_QTY)
                    Call UniCode_Conv(ODR_KNT_R.ODR_QTY, W_STR)
                               
                               
                If Trim(ORDR_GRID(Cor_Row, Col_KIBOU_DT)) <> "" Then
                    Call UniCode_Conv(ODR_KNT_R.NOUKI, _
                        Format(Trim(ORDR_GRID(Cor_Row, Col_KIBOU_DT)), "yyyymmdd"))
                End If
                
                '�[����
                Call UniCode_Conv(ODR_KNT_R.NONYU, Trim(ORDR_GRID(Cor_Row, Col_DELI_CD)))
                
                Do
                    sts = BTRV(BtOpUpdate, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
                    Select Case sts
                        Case BtNoErr
                                
                                Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                                Sleep (500)
                        Case Else
                            Call File_Error(sts, BtOpUpdate, "ODR_KENTO")
                            GoTo Err_exit
                    End Select
                Loop
                
            End If
            
        
        
        'End If
        
    Next Cor_Row
    
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i�������@�X�V�I���B�@��KENTO_UPDT2��", Me.hwnd, 0)
 
    KENTO_UPDT2 = False
    
Err_exit:
        
    Call Input_UnLock
    
    
    sts = BTRV(BtOpClose, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_KENTO")
        End If
    End If
End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------
Dim i As Integer

    ODR30101.MousePointer = vbHourglass

    Call Ctrl_Lock(ODR30101)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(ODR30101)


    ODR30101.MousePointer = vbDefault

End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Tab_Ctrl(Shift)        '�ړ�
    End If

End Sub

Private Sub Command1_Click(Index As Integer)
Dim yn      As Integer

Dim X_i     As Integer
Dim X_j     As Integer

Dim W_After     As String

    Select Case Index
        Case FuncDIS%
        '-------------------    �\��
            yn = vbYes
            
            If Grid_Cor_M = True Then
                yn = MsgBox("�X�V����Ă��܂���I" & Chr(13) & Chr(10) & _
                            "�@�ĕ\�����܂����H", vbYesNo + vbExclamation, "�m�F����")
            End If
            
            If yn = vbYes Then
                '�ŐV�\������
                If Data_Disp Then
                    MsgBox "�w������̒������@�\�����s�I", vbExclamation
                    Call Text1_GotFocus(ptxTOP%)
                    Text1(ptxTOP%).SetFocus
                    Exit Sub
                End If
                
                TDBGrid1.SetFocus
                
                Grid_Cor_M = False
                
            End If
                        
            Exit Sub
    
        Case FuncKIBOU%
        '-------------------    ��]�[��
            'If Grid_Cor_M <> True Then
            '    Exit Sub
            'End If
            
            If IsNull(TDBGrid1.Bookmark) Then Exit Sub
            
            If TDBGrid1.Bookmark <= 0 Then Exit Sub
            
            For X_i = ptxTOP To ptxLAST
                If ERR_CHK(X_i) Then
                    Text1(X_i).SetFocus
                    Call Text1_GotFocus(X_i)
                    Exit Sub
                End If
                
            Next X_i
        
            ODR30105.Show vbModal
            
            If ODR30105_Return = True Then Exit Sub     '�L�����Z��
            
            'MsgBox "��]�[���F" & KIBOU_DT
                        
            If KIBOU_UPDT Then
                MsgBox "��]�[���@�ꊇ�X�V�G���[�I", vbExclamation
            End If
                        
            Exit Sub
        
        Case FuncCOR
        '-------------------    �X�V
            If IsNull(TDBGrid1.Bookmark) Then Exit Sub
            For X_i = ptxTOP To ptxLAST
                If ERR_CHK(X_i) Then
                    Text1(X_i).SetFocus
                    Call Text1_GotFocus(X_i)
                    Exit Sub
                End If
            Next X_i
            
            If Grid_Cor_M <> True Then
                Exit Sub
            End If
            
            TDBGrid1.Update
            Set ORDR_GRID = TDBGrid1.Array
    
            For X_j = Min_Row To ORDR_GRID.UpperBound(1)
            
                For X_i = Col_ITEM To Col_DELI_CD%
                    
                    W_After = ORDR_GRID(X_j, X_i)
                    
                    Cor_Row = X_j
                    
                    If Grid_Err_Chk(X_i, W_After) Then
                        row = X_j
                        TDBGrid1.SetFocus
                        Exit Sub
                    End If
                
                Next X_i
                
            Next X_j
            
            
            yn = MsgBox("�X�V���܂����H", vbYesNo + vbDefaultButton2 + vbQuestion, "�m�F����")
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            '�X�V����
            If KENTO_UPDT(False) Then
                MsgBox "�X�V���s���܂����B", vbExclamation
                
                Exit Sub
            End If
            '�ŐV�\������
            If Data_Disp() Then
                MsgBox "�w������̒������@�\�����s�I", vbExclamation
                Call Text1_GotFocus(ptxTOP%)
                Text1(ptxTOP%).SetFocus
                Exit Sub
            End If
            
            Grid_Cor_M = False
            
            TDBGrid1.SetFocus
                        
            Exit Sub
            
            
        Case FuncORDER
        '-------------------    ���������s
            If IsNull(TDBGrid1.Bookmark) Then Exit Sub
            
            For X_i = ptxTOP To ptxLAST
                If ERR_CHK(X_i) Then
                    Text1(X_i).SetFocus
                    Call Text1_GotFocus(X_i)
                    Exit Sub
                End If
                
            Next X_i
            
            If row < Min_Row Then
                Exit Sub
            End If
            
            TDBGrid1.Update
            Set ORDR_GRID = TDBGrid1.Array
    
    
    
            For X_j = Min_Row To ORDR_GRID.UpperBound(1)
            
                For X_i = Col_ITEM To Col_KAITO_DT%
                    
                    W_After = ORDR_GRID(X_j, X_i)
                    
                    Cor_Row = X_j
                    
                    If Grid_Err_Chk(X_i, W_After) Then
                        row = X_j
                        TDBGrid1.SetFocus
                        Exit Sub
                    End If
                
                Next X_i
                
            Next X_j
            
            
            yn = MsgBox("���������s���܂����H", vbYesNo + vbDefaultButton2 + vbQuestion, "�m�F����")
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            '�X�V����
            
            If KENTO_UPDT(True) Then
                MsgBox "�X�V���s���܂����B", vbExclamation
                
                Exit Sub
            End If
            
            If SHORDER_Update() Then
                Unload Me
            End If
            
            If Print_Proc() Then
                Unload Me
            End If
        
        
            '�ŐV�\������
            If Data_Disp() Then
                MsgBox "�w������̒������@�\�����s�I", vbExclamation
                Call Text1_GotFocus(ptxTOP%)
                Text1(ptxTOP%).SetFocus
                Exit Sub
            End If
            
            Grid_Cor_M = False
            
            TDBGrid1.SetFocus
                        
            Exit Sub
            
        
        Case FuncLIST
        '-------------------    ���X�g���s
        
            If IsNull(TDBGrid1.Bookmark) Then Exit Sub
            
            
            
            yn = MsgBox("���X�g�o�͂��܂����H", vbYesNo + vbDefaultButton2 + vbQuestion, "�m�F����")
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
            "��ʕ\�����e�@�G���[�`�F�b�N���I�@��ERR_CHK��", Me.hwnd, 0)
            DoEvents
            For X_i = ptxTOP To ptxLAST
                If ERR_CHK(X_i) Then
                    Text1(X_i).SetFocus
                    Call Text1_GotFocus(X_i)
                    Exit Sub
                End If
                
            Next X_i
            
            TDBGrid1.Update
            Set ORDR_GRID = TDBGrid1.Array
    
    
            If row < Min_Row Then
                Exit Sub
            End If
    
            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
            "��ʕ\�����e�@�G���[�`�F�b�N���I�@��Grid_Err_Chk��", Me.hwnd, 0)
            DoEvents
            For X_j = Min_Row To ORDR_GRID.UpperBound(1)
            
                For X_i = Col_ITEM To Col_KAITO_DT%
                    
                    W_After = ORDR_GRID(X_j, X_i)
                    
                    Cor_Row = X_j                           '2009/04/04
                    
                    If Grid_Err_Chk(X_i, W_After) Then
                        row = X_j
                        TDBGrid1.SetFocus
                        Exit Sub
                    End If
                
                Next X_i
                
            Next X_j
            
            '               2010/01/18  �㕔�Ɉړ��B
            'hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
            '"���X�g�o�͊m�F", Me.hwnd, 0)
            'DoEvents
            'yn = MsgBox("���X�g�o�͂��܂����H", vbYesNo + vbDefaultButton2 + vbQuestion, "�m�F����")
            'If yn = vbNo Then
            '    Exit Sub
            'End If
            
            
            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
            "�\���f�[�^�X�V���B", Me.hwnd, 0)
            DoEvents
            
            If KENTO_UPDT2 Then                             '��ʓ��e���X�V����̂݁I   2010/01/19
                MsgBox "�X�V���s���܂����B", vbExclamation
                
                Exit Sub
            End If
            
            
            'If SHORDER_Update() Then   2010/01/19
            '    Unload Me
            'End If
            
            
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2010/01/16
            Key_USE_YM = Left(Text1(ptxUSE_YY), 4) & Right(Text1(ptxUSE_YY), 2)
            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
            "�Ώۃf�[�^�@������B", Me.hwnd, 0)
            DoEvents
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            
            If List_Print_Proc() Then
                Unload Me
            End If
        
            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
            "�Ώۃf�[�^�@����I���B", Me.hwnd, 0)
            DoEvents
            '�ŐV�\������
            If Data_Disp() Then
                MsgBox "�w������̒������@�\�����s�I", vbExclamation
                Call Text1_GotFocus(ptxTOP%)
                Text1(ptxTOP%).SetFocus
                Exit Sub
            End If
            
            Grid_Cor_M = False
            
            TDBGrid1.SetFocus
                        
            Exit Sub
        
        
        
        
        
        
        Case FuncEND%
            
            'yn = MsgBox("�I�����܂����H", vbYesNo + vbDefaultButton1 + vbQuestion, "�m�F����")
            yn = vbYes
            If Grid_Cor_M = True Then
                yn = MsgBox("�X�V����Ă��܂���I�I" & Chr(13) & Chr(10) & _
                            "�@�I�����܂����H", vbYesNo + vbDefaultButton2 + vbQuestion, "�m�F����")
            End If
            
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            Unload Me
    
    End Select

End Sub

Private Sub Form_Activate()
    
    Text1(ptxTOP).SetFocus          '2015.11.13

End Sub

Private Sub Form_Load()

Dim i       As Integer
Dim c       As String * 128
Dim sts     As Integer

Dim sBuffer As String * 255
Dim com     As String

Dim W_Date  As String

Dim wYY     As String * 4
Dim wMM     As String * 2
Dim wDD     As String * 2

Init_F_30101 = 0

    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "������������", Me.hwnd, 0)
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2016.0314
'    WS_NO = Space(255)
'    If GetComputerNameA(WS_NO, 255) <> 0 Then
'        WS_NO = Left(WS_NO, InStr(WS_NO, vbNullChar) - 1)
'    Else
'        WS_NO = "000"
'    End If

    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2016.0314

'��ʏ�������
    Show
    
    '                               '2009.01.15�v�]�ɂ��A�~�߂��I(-_-;)
    'If App.PrevInstance Then
    '    Beep
    '    MsgBox "����v���O�������s���ł��B", vbExclamation
    '    End
    'End If
    
    
    
    
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)

                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '�g�p���ʌ����݌ɂe
    'If ODR_ZAIKO_Open(BtOpenNomal) Then
    '    Unload Me
    'End If
    
                                '�Ǘ��}�X�^�n�o�d�m
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '�R�[�h�}�X�^�n�o�d�m
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '�����f�[�^�n�o�d�m
    If P_SHORDER_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '���ގ���ް��n�o�d�m
    If P_SHUKEIRE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�d����}�X�^�n�o�d�m
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If


                                '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                
                                '�e�i�Ԓ����e�n�o�d�m
    If ODR_ORDER_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '���������n�o�d�m
    'If ODR_KENTO_Open(BtOpenNomal) Then
    '    Unload Me
    'End If
                                
                                
                                '���ޒ����ް��n�o�d�m(���߲���)
    If wP_SHORDER_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                
                                '���ƕ��̊l��
    If JGYOB_TB_Set() Then
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B"
        End
    End If
    
    Combo1(pcmbJI).Clear
    
    For i = 0 To UBound(JGYOBU_T) - 1
        Combo1(pcmbJI).AddItem JGYOBU_T(i).NAME & Space(5) & JGYOBU_T(i).CODE
    
    Next i
    Combo1(pcmbJI).ListIndex = 0

                                '�����O�Ǘ��̊l��
    i = 0
    Do
        i = i + 1
        If GetIni(App.EXEName, "NAIGAI" & Format(i, "0"), "SYS_ODR3010", c) Then
            Exit Do
        End If
        ReDim Preserve NAIGAI_CODE(i - 1)
        NAIGAI_CODE(i - 1) = Trim(c)
    
    Loop
    If i = 1 Then
        MsgBox "�����O�̊l���Ɏ��s���܂����B"
        End
    End If
    
    
    '����Ͻ���`
    Call P_CODE_TBL_Proc
'���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    End If
    
    If SET_JGYOBU_T Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    End If
    
    
                                '���l�P��荞��
    If GetIni(App.EXEName, "BIKOU_1", "P_SYS", c) Then
        pubBikou_1 = ""
    Else
        pubBikou_1 = Trim(c)
    End If
                                '���l�Q��荞��
    If GetIni(App.EXEName, "BIKOU_2", "P_SYS", c) Then
        pubBikou_2 = ""
    Else
        pubBikou_2 = Trim(c)
    End If
                                '���l�R��荞��
    If GetIni(App.EXEName, "BIKOU_3", "P_SYS", c) Then
        pubBikou_3 = ""
    Else
        pubBikou_3 = Trim(c)
    End If
    
    
    
                                '�[����
    If GetIni(App.EXEName, "DELI_CODE", "SYS_ODR3010", c) Then
        NOUNYU = ""
    Else
        NOUNYU = Trim(c)
    End If
    
    
    
    
    
    
    
    GW_SIMUKE = Left(Right(Combo1(pcmbSM).Text, 4), 2)
    'GW_JIGYOBU = Right(Combo1(pcmbJI).Text, 1)
    GW_JIGYOBU = Mid(Right(Combo1(pcmbSM).Text, 4), 3, 1)
    GW_NAIGAI = Right(Combo1(pcmbSM).Text, 1)
        
    GW_HINGAI = ""
    
    
    '                       2008/07/02  GetIni�̌�Ɉړ��B
    'If Require_Set Then
    '    MsgBox "�ŐV���X�V�@���s�I", vbExclamation
    '    Call Text1_GotFocus(ptxTOP%)
    '    Text1(ptxTOP%).SetFocus
    '    Exit Sub
    'End If
    

'    If GetIni("PR00030", "LAST_SHIME_DT01", "P_SYS", c) Then           '2016.01.12
    If GetIni("PR00030", "LAST_SHIME_DT01", "PR00030", c) Then          '2016.01.12
        GW_TOUGETU = Left(Format(Date, "yyyymmdd"), 6)
        GW_SHIMEBI = Format(Date, "yyyymmdd")
    Else
        GW_TOUGETU = Left(Format(Trim(c), "yyyymmdd"), 6)
            
        wYY = Left(GW_TOUGETU, 4)
        wMM = Right(GW_TOUGETU, 2)
        wDD = Right(Format(Trim(c), "yyyymmdd"), 2)
                    
        GW_SHIMEBI = Format(Trim(c), "yyyymmdd")
        'If wDD <= "20" Then
        '
        'Else
        '
        '    wMM = Format(CInt(wMM) + 1, "00")
        '
        '    If wMM > "12" Then
        '        wYY = Format(CInt(wYY) + 1, "0000")
        '        wMM = "01"
        '    End If
        'End If
    
        GW_TOUGETU = wYY & wMM
    
    End If
    
    W_Date = Left(GW_SHIMEBI, 4) & "/" & Mid(GW_SHIMEBI, 5, 2) & "/" & Right(GW_SHIMEBI, 2)
    Lab_Dsp(2) = "�J�z����" & W_Date
    
    Text1(ptxUSE_YY) = Left(GW_TOUGETU, 4) & "/" & Right(GW_TOUGETU, 2)
    
    '�ő���t�̐ݒ�
    
    GW_MAX_YYMM = Left(Format(DateAdd("m", 20, W_Date), "yyyy/mm/dd"), 7)

    'If Require_Set Then
    '    MsgBox "�ŐV���X�V�@���s�I", vbExclamation
    '    Call Text1_GotFocus(ptxTOP%)
    '    Text1(ptxTOP%).SetFocus
    '    Exit Sub
    'End If
    

    Grid_Cor_M = False
    Grid_Req_M = False
    
    row = Min_Row - 1
    
    TDBGrid1.Bookmark = -1
    
    ODR30101.Caption = Last_Update_Day      '2016.01.12
    
    Load ODR30102
    Load ODR30103
    Load ODR30104
    Load ODR30105
    
    'Combo1(pcmbSM).SetFocus
'    Text1(ptxTOP).SetFocus         2015.11.13
        

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim yn      As Integer

    If UnloadMode = 1 Then Exit Sub
    
    yn = MsgBox("�I�����܂����H", vbYesNo + vbDefaultButton1 + vbQuestion, "�m�F����")
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

    
    sts = BTRV(BtOpClose, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_ORDER, Len(K0_ODR_ORDER), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_ORDER")
        End If
    End If

    'sts = BTRV(BtOpClose, ODR_ZK_POS, ODR_ZK_R, Len(ODR_ZK_R), K0_ODR_ZK, Len(K0_ODR_ZK), 0)
    'If sts Then
    '    If sts <> BtErrNoOpen Then
    '        Call File_Error(sts, BtOpClose, "ODR_ZAIKO")
    '    End If
    'End If
    
    
    sts = BTRV(BtOpClose, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_KENTO")
        End If
    End If

    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "P_CODE")
        End If
    End If

    
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "P_KANRI")
        End If
    End If
    
    sts = BTRV(BtOpClose, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K0_P_SHUKEIRE, Len(K0_P_SHUKEIRE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "P_SHUKEIRE")
        End If
    End If
    
    
    sts = BTRV(BtOpClose, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "P_SHORDER")
        End If
    End If
    
    sts = BTRV(BtOpClose, wP_SHORDER_POS, wP_SHORDER_REC, Len(wP_SHORDER_REC), K2_wP_SHORDER, Len(K2_wP_SHORDER), 0)
    
    
    
    
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "P_UKEHARAI")
        End If
    End If
    
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "TANTO")
        End If
    End If
    


    End
End Sub

Private Sub SHORI_Click(Index As Integer)
Dim yn      As Integer


    Select Case Index
    
        Case 0      '�\��
            Call Command1_Click(FuncDIS)
            
        Case 1      '�ŐV
            Call Command1_Click(FuncKIBOU)
        
        Case 2      '�X�V
            Call Command1_Click(FuncCOR)
        
        Case 3      '��ʈ��
            yn = MsgBox("��ʈ�����܂����H", vbYesNo + vbDefaultButton2 + vbQuestion, "�m�F����")
            If yn = vbNo Then
                
                Exit Sub
            End If
            
            Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)
        
        
        Case 4      '�I��
            Call Command1_Click(FuncEND)
    
    End Select


End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
Dim W_STR       As String
    
Dim W_Before    As String
Dim W_After     As String
    
    If IsNull(TDBGrid1.Bookmark) Then Exit Sub
    If TDBGrid1.Bookmark <= 0 Then Exit Sub
    
    Cor_Row = TDBGrid1.Bookmark
    
    'W_Before = Trim(ORDR_GRID(Cor_Row, ColIndex))
    W_After = Trim(TDBGrid1.Text)
    
    
    TDBGrid1.Update
    Set ORDR_GRID = TDBGrid1.Array
    'If W_Before <> W_After Then
    '    Grid_Cor_M = True
    'End If
    
    'If Grid_Err_Chk(ColIndex, W_Before, W_After) Then
        
    '    Exit Sub
    'End If
    

End Sub

Private Sub TDBGrid1_BeforeInsert(Cancel As Integer)

    ORDR_GRID.ReDim Min_Row, ORDR_GRID.Count(1), Min_Col, Max_Col
    
End Sub
Private Sub TDBGrid1_Change()

    Grid_Cor_M = True

End Sub

Private Sub TDBGrid1_DblClick()

Dim W_SHIRE_ZAN     As String

    If IsNull(TDBGrid1.Bookmark) Then Exit Sub
    
    If TDBGrid1.Bookmark = -1 Then
    
    
    Else

        
        If Option1(0).Value Then
        
        
            Set ORDR_GRID = TDBGrid1.Array
            
            '       ���[�̉ۃ`�F�b�N
            '
            '       �@�e���i���������w�肵�Ă��鎖�I
            '       �A�������̎��I
            '       �B���[�̐e�i��j�����w�����鎖�I?
            '
            
            
            W_SHIRE_ZAN = Trim(ORDR_GRID(TDBGrid1.Bookmark, Col_ZAN_QTY%)) '�d���c��
            If W_SHIRE_ZAN = "" Then W_SHIRE_ZAN = "0"
            If CDbl(W_SHIRE_ZAN) = 0 Then
                MsgBox "�d���c�́A����܂���", vbExclamation
            '    Exit Sub
            End If
            
    
    
            
            'Key_BUN_NO = Trim(ORDR_GRID(TDBGrid1.Bookmark, Col_BUNNO%))
                    
            '           ���[�w����ʂɈڍs�I
            Key_USE_YM = ORDR_GRID(TDBGrid1.Bookmark, Col_KEY%)
            Key_JIGYOBU = ORDR_GRID(TDBGrid1.Bookmark, Col_JIGYOBU%)
            Key_NAIGAI = ORDR_GRID(TDBGrid1.Bookmark, Col_NAIGAI%)
            Key_HinGai = ORDR_GRID(TDBGrid1.Bookmark, Col_ITEM%)
    
    
    
            DIS_ITEM = ORDR_GRID(TDBGrid1.Bookmark, Col_ITEM%)          '�q���i�R�[�h
            DIS_ITEM_NM = ORDR_GRID(TDBGrid1.Bookmark, Col_ITEM_NM%)    '�q���i��
            DIS_USE_QTY = ORDR_GRID(TDBGrid1.Bookmark, Col_USE_QTY%)    '�g�p����
            DIS_MRP_QTY = ORDR_GRID(TDBGrid1.Bookmark, Col_MRP_QTY%)    '�K�v��
            DIS_ZAI_QTY = ORDR_GRID(TDBGrid1.Bookmark, Col_ZAI_QTY%)    '�����݌�
            DIS_FUSOKU = ORDR_GRID(TDBGrid1.Bookmark, Col_FUSOKU%)      '�s����
            DIS_ORDR_QTY = ORDR_GRID(TDBGrid1.Bookmark, Col_ORDR_QTY%)  '������
            
            
            
            DIS_ZAN_QTY = ORDR_GRID(TDBGrid1.Bookmark, Col_ZAN_QTY%)    '�d���c
            
                                                                        '�����i
            DIS_HANSEIHIN_QTY = ORDR_GRID(TDBGrid1.Bookmark, Col_HANSEIHIN_QTY%)
            
            
            
            
            DIS_LOT_QTY = ORDR_GRID(TDBGrid1.Bookmark, Col_LOT_QTY%)    '���b�g��
            DIS_SECT_CD = ORDR_GRID(TDBGrid1.Bookmark, Col_SECT_CD%)    '�d����
            DIS_SECT_NM = ORDR_GRID(TDBGrid1.Bookmark, Col_SECT_NM%)    '�d���於
            DIS_TANKA = ORDR_GRID(TDBGrid1.Bookmark, Col_TANKA%)        '�d���P��
            DIS_KIBOU_DT = ORDR_GRID(TDBGrid1.Bookmark, Col_KIBOU_DT%)  '��]�[��
            DIS_KAITO_DE = ORDR_GRID(TDBGrid1.Bookmark, Col_KAITO_DT%)  '�񓚔[��
            
            
            DoEvents
            
            ODR30102.Show vbModal
            
            If ODR30102_Return = True Then Exit Sub     '�L�����Z��
            
            '���[���𔽉f���čĕ\������B
            
            'If Data_Disp Then
            '    MsgBox "�w������̒������ŁA�\�����s�I", vbExclamation
            '    Call Input_UnLock                             '��ʍ��ڃ��b�N
            '    Call Text1_GotFocus(ptxTOP%)
            '    Text1(ptxTOP%).SetFocus
            '    Exit Sub
            'End If
        
        End If
        
        
        If Option1(1).Value Then
        
            Key_USE_YM = ORDR_GRID(TDBGrid1.Bookmark, Col_KEY%)
            Key_JIGYOBU = ORDR_GRID(TDBGrid1.Bookmark, Col_JIGYOBU%)
            Key_NAIGAI = ORDR_GRID(TDBGrid1.Bookmark, Col_NAIGAI%)
            Key_HinGai = ORDR_GRID(TDBGrid1.Bookmark, Col_ITEM%)
        
            
            ODR30104.Show vbModal
            
            If ODR30104_Return = True Then Exit Sub     '�L�����Z��
        
        
        End If
        
        
    End If

End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
Dim yn          As Integer
Dim W_Index     As Integer

    'TDBGrid1.Bookmark = -1
    W_Index = ColIndex
    
    If row <= 1 Then Exit Sub
        
    yn = MsgBox("���׊����܂����H", vbYesNo + vbExclamation, "�m�F����")
    If yn <> vbYes Then Exit Sub
    Select Case ColIndex
        Case Col_ITEM%                  '�q���i�R�[�h
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_ITEM%, XORDER_ASCEND, XTYPE_STRING
                        
        Case Col_USE_QTY%               '�g�p����
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_USE_QTY%, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ITEM%, XORDER_ASCEND, XTYPE_STRING
              
        Case Col_ZAI_QTY%                '�����݌�
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_ZAI_QTY%, XORDER_ASCEND, XTYPE_DOUBLE, _
                        Col_ITEM%, XORDER_ASCEND, XTYPE_STRING
                        
        Case Col_FUSOKU%                '�s����
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_FUSOKU%, XORDER_ASCEND, XTYPE_DOUBLE, _
                        Col_ITEM%, XORDER_ASCEND, XTYPE_STRING
              
        Case Col_ORDR_QTY%               '������
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_ORDR_QTY%, XORDER_ASCEND, XTYPE_DOUBLE, _
                        Col_ITEM%, XORDER_ASCEND, XTYPE_STRING
        
        Case Col_SECT_CD%               '�d���於
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_SECT_CD%, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ITEM%, XORDER_ASCEND, XTYPE_STRING
              
        Case Col_TANKA%                 '�d���P��
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_TANKA%, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ITEM%, XORDER_ASCEND, XTYPE_STRING
                        
        Case Col_KIBOU_DT%               '��]�[��
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_KIBOU_DT%, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ITEM%, XORDER_ASCEND, XTYPE_STRING
              
        Case Col_KAITO_DT%               '�񓚔[��
            ORDR_GRID.QuickSort Min_Row, (ORDR_GRID.UpperBound(1)), _
                        Col_KAITO_DT%, XORDER_ASCEND, XTYPE_STRING, _
                        Col_ITEM%, XORDER_ASCEND, XTYPE_STRING
        Case Else
            MsgBox "���ב֎w�� ���O���ځI", vbExclamation
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
    
    If Text1(Index).Locked = True Then      '���b�N�����ڂȂ珈�����Ȃ�
        Call Tab_Ctrl(Shift)    '�ړ�
        Exit Sub
    End If
                        '���͕������`�F�b�N
    If ERR_CHK(Index) Then
        Call Text1_GotFocus(Index)
        Text1(Index).SetFocus
        Exit Sub
    End If
    
    If Index = ptxTOP And Init_F_30101 = 0 Then
        If Data_Disp Then
            MsgBox "�w������̒������ŁA�\�����s�I", vbExclamation
            Call Text1_GotFocus(ptxTOP%)
            Text1(ptxTOP%).SetFocus
            Exit Sub
        End If
        Init_F_30101 = 1
        Call Text1_GotFocus(ptxUSE_YY)
        Text1(ptxUSE_YY).SetFocus
        Exit Sub
    End If
    
    Call Tab_Ctrl(Shift)    '�ړ�
    
End Sub
Private Function KIBOU_UPDT() As Integer
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer
    
Dim X_i             As Long
    
    KIBOU_UPDT = True
    
                                        
    Set TDBGrid1.Array = ORDR_GRID
    TDBGrid1.Refresh

    TDBGrid1.Update

                                        
    For X_i = 1 To ORDR_GRID.UpperBound(1)
      '�������̗L��s�̂ݑΏۂƂ���!     �l�s
        If IsNumeric(ORDR_GRID(X_i, Col_ORDR_QTY%)) Then
            If CLng(ORDR_GRID(X_i, Col_ORDR_QTY%)) > 0 Then
                ORDR_GRID(X_i, Col_KIBOU_DT) = KIBOU_DT       '��]�[��
            
            End If
        End If
    Next X_i
    
    Set TDBGrid1.Array = ORDR_GRID
    
    'TDBGrid1.style.Locked = True
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    'TDBGrid1.MoveFirst
    'TDBGrid1.ScrollBars = dbgAutomatic
    
    
    KIBOU_UPDT = False
End Function

Private Function SHORDER_Update() As Integer
'----------------------------------------------------------------------------
'                  ���ޒ����ް��X�V
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer

Dim ORDERNO         As Integer

Dim i               As Integer
Dim j               As Integer

Dim W_Test          As String

    SHORDER_Update = True
                                        
    Call Input_Lock
                                        
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If

    Set TDBGrid1.Array = ORDR_GRID
    TDBGrid1.Refresh

    TDBGrid1.Update

                                        
    For i = 1 To ORDR_GRID.UpperBound(1)
        
        
      '2008/09/19 �K�v���̗L��s�̂ݑΏۂƂ���!     �l�s
      If IsNumeric(ORDR_GRID(i, Col_MRP_QTY)) Then
        If CLng(ORDR_GRID(i, Col_MRP_QTY)) > 0 Then
        
        
        
          If IsNumeric(ORDR_GRID(i, Col_ORDR_QTY)) Then
            If CLng(ORDR_GRID(i, Col_ORDR_QTY)) > 0 Then
    
    
                            
                '�󕥐�Ͻ��Ǎ���            '2009/04/03
                                            '�}�X�^�[�`�F�b�N���āA������Ή������Ȃ��I
                                            '
            Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, ORDR_GRID(i, Col_SECT_CD))
            sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
              
              If sts = BtNoErr Then
            
    
                '�Ǘ��t�@�C����莑�ޒ����ԍ��̊l��
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
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                SHORDER_Update = True
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�Ǘ��}�X�^")
                            GoTo Abort_Tran
                    
                    End Select
                
                
                Loop
            
                '�������{�P
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
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                sts = BTRV(BtOpUnlock, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                                If sts Then
                                    Call File_Error(sts, BtOpUnlock, "�Ǘ��}�X�^")
                                End If
                                GoTo Abort_Tran
                            End If
                        Case Else
                            Call File_Error(sts, BtOpUpdate, "�Ǘ��}�X�^")
                            GoTo Abort_Tran
                    End Select
                Loop
        
                ORDERNO = CLng(StrConv(P_KANRIREC.ORDER_NO, vbUnicode))
    
    
                                        
                '������
                Call UniCode_Conv(P_SHORDER_REC.ORDER_NO, Format(ORDERNO, "00000"))
                '������
                Call UniCode_Conv(P_SHORDER_REC.ORDER_DT, Format(Now, "YYYYMMDD"))
                '���s����
                Call UniCode_Conv(P_SHORDER_REC.Print_datetime, "")
                '�S���Һ���
                Call UniCode_Conv(P_SHORDER_REC.TANTO_CODE, Text1(ptxTANTO_CD).Text)
                '���ƕ�
                Call UniCode_Conv(P_SHORDER_REC.JGYOBU, ORDR_GRID(i, Col_JIGYOBU))
                '�����O
                Call UniCode_Conv(P_SHORDER_REC.NAIGAI, ORDR_GRID(i, Col_NAIGAI))
                '�i��
                Call UniCode_Conv(P_SHORDER_REC.HIN_GAI, ORDR_GRID(i, Col_ITEM))
                '������
                Call UniCode_Conv(P_SHORDER_REC.ORDER_CODE, ORDR_GRID(i, Col_SECT_CD))
                '�[����
                Call UniCode_Conv(P_SHORDER_REC.DELI_CODE, ORDR_GRID(i, Col_DELI_CD))
                '������
                Call UniCode_Conv(P_SHORDER_REC.ORDER_QTY, Format(CDbl(ORDR_GRID(i, Col_ORDR_QTY)), "00000000.00"))
                '�\��[��
                Call UniCode_Conv(P_SHORDER_REC.Y_NOUKI_DT, Format(ORDR_GRID(i, Col_KIBOU_DT), "YYYYMMDD"))
                '�����P��
                Call UniCode_Conv(P_SHORDER_REC.TANKA, Format(CDbl(ORDR_GRID(i, Col_TANKA)), "00000000.00"))
                '����ۯ�
                Call UniCode_Conv(P_SHORDER_REC.LOT, Format(CDbl(ORDR_GRID(i, Col_LOT_QTY)), "00000000"))
                    
                Call UniCode_Conv(P_SHORDER_REC.KAN_F, P_KAN_OFF)                       '�����׸�
                    
                Call UniCode_Conv(P_SHORDER_REC.KAN_DT, "")                             '������
                    
                Call UniCode_Conv(P_SHORDER_REC.BUNNOU_CNT, "00")                       '�����
                    
                Call UniCode_Conv(P_SHORDER_REC.UKEIRE_QTY, "00000000.00")              '�����
                
                Call UniCode_Conv(P_SHORDER_REC.CANCEL_F, P_CANCEL_OFF)                 '��ݾ��׸�
                    
                Call UniCode_Conv(P_SHORDER_REC.CANCEL_DATETIME, "")                    '��ݾٓ���
                
                Call UniCode_Conv(P_SHORDER_REC.PRINT_F, P_PRINT_OFF)                   '����׸�
                
                Call UniCode_Conv(P_SHORDER_REC.WS_NO, WS_NO)                           '���͒[��
                
                
                '�i��Ͻ��Ǎ���
                Call UniCode_Conv(K0_ITEM.JGYOBU, ORDR_GRID(i, Col_JIGYOBU))
                Call UniCode_Conv(K0_ITEM.NAIGAI, ORDR_GRID(i, Col_NAIGAI))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, ORDR_GRID(i, Col_ITEM))
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    
                        MsgBox "�i�ڃ}�X�^�����[���ŕύX����܂����B�X�V�����𒆎~���܂��B"
                        GoTo Abort_Tran
                                                
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                        GoTo Abort_Tran
                End Select
                '�d���敪
                Call UniCode_Conv(P_SHORDER_REC.G_SHIIRE_KBN, StrConv(ITEMREC.G_SHIIRE_KBN, vbUnicode))
                '���x�P��
                Call UniCode_Conv(P_SHORDER_REC.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))
                
                
                '�󕥐�Ͻ��Ǎ���
                Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, ORDR_GRID(i, Col_SECT_CD))
                sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                    
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        MsgBox "�󕥐�}�X�^�����[���ŕύX����܂����B�X�V�����𒆎~���܂��B" & "<" & ORDR_GRID(i, Col_SECT_CD) & ">"
                        GoTo Abort_Tran
                        
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�󕥐�Ͻ�")
                        GoTo Abort_Tran
                End Select
            
                                                                                            '�����敪
                Call UniCode_Conv(P_SHORDER_REC.TORI_KBN, StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode))
            
                Call UniCode_Conv(P_SHORDER_REC.FILLER, "")
                                                                                            '�X�V����
                Call UniCode_Conv(P_SHORDER_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                
                
                
                Call UniCode_Conv(P_SHORDER_REC.ANS_NOUKI_DT, "")                           '�[���񓚓�
                                                                                            '�g�p�N��
                Call UniCode_Conv(P_SHORDER_REC.USE_YM, Mid(Format(Text1(ptxUSE_YY).Text & "/01", "YYYYMMDD"), 1, 7))
                
                
                
                
                Do
                    
                    DoEvents
                    
                    sts = BTRV(BtOpInsert, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SHORDER.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                If com = BtOpUpdate Then
                                    sts = BTRV(BtOpUnlock, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
                                    If sts Then
                                        Call File_Error(sts, BtOpUnlock, "���ޒ����ް�")
                                    End If
                                End If
                                GoTo Abort_Tran
                            End If
                        Case Else
                            Call File_Error(sts, BtOpInsert, "���ޒ����ް�")
                            GoTo Abort_Tran
                    End Select
                
                Loop
                
                
                '---------------------------------------------------    '�i�ڃ}�X�^�X�V
                Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
                
                Do
                
                    sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        
                    Select Case sts
                        Case BtNoErr
                        
                            Exit Do
                        
                        Case BtErrKeyNotFound
                        
                            MsgBox "�i�ڃ}�X�^���폜����Ă��܂��B�X�V�𒆎~���܂��B"
                            GoTo Abort_Tran
                        
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                GoTo Abort_Tran
                            End If
                        
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                            GoTo Abort_Tran
                    End Select
            
                Loop
                
                For j = 0 To 2
                
                    If Trim(StrConv(ITEMREC.G_SHIIRE_TBL(j).CODE, vbUnicode)) = Trim(StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode)) Then
                        Exit For
                    End If
                Next j
                
                
                If j <= 2 Then
                    '�O�񒍕���
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LAST_ORDER_DT, StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode))
                    '�O�񒍕���
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LAST_ORDER_QTY, StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode))
                End If
            
            
                Do
                    
                    DoEvents
                    
                    sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                If sts Then
                                    Call File_Error(sts, BtOpUnlock, "�i��Ͻ�")
                                End If
                            End If
                            GoTo Abort_Tran
                        Case Else
                            Call File_Error(sts, BtOpUpdate, "�i��Ͻ�")
                            GoTo Abort_Tran
                    End Select
                
                Loop
                End If
                
                Else
                    W_Test = Trim(ORDR_GRID(i, Col_SECT_CD))
                
              End If
            
        
            End If      '2009/04/03
        
          End If
        End If
        
    Next i
        

End_Tran:
                                        '�g�����U�N�V�����I��
    sts = BTRV(BtOpEndTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
    Call Input_UnLock
    
    SHORDER_Update = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    
    Call Input_UnLock

End Function

Private Function Print_Proc() As Integer
'----------------------------------------------------------------------------
'           �������������
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
Dim Save_Order_Code As String * 5
                
Dim rpt         As New ODR3010F1
Dim f           As New ODR30103

                
    Print_Proc = True
                
    Call Input_Lock
                
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
                Call Input_UnLock
                Call File_Error(sts, BtOpGetEqual, "�󕥐�}�X�^")
                Exit Function
        End Select
    
        If Trim(Save_Order_Code) = "" Then
    
            Set rpt = New ODR3010F1
        
            '���|�[�g��������܂��B�itrue�F����_�C�A���O���� false�F�Ȃ��j
            rpt.PrintReport False
        
            Set rpt = Nothing
    
    
    
            'f.RunReport rpt
            'f.Show vbModal
    
            Save_Order_Code = StrConv(wP_SHORDER_REC.ORDER_CODE, vbUnicode)
    
    
        End If
    
        If Save_Order_Code <> StrConv(wP_SHORDER_REC.ORDER_CODE, vbUnicode) Then
    
            Set rpt = New ODR3010F1
        
            '���|�[�g��������܂��B�itrue�F����_�C�A���O���� false�F�Ȃ��j
            rpt.PrintReport False
        
            Set rpt = Nothing


'            f.RunReport rpt
'            f.Show
    
            Save_Order_Code = StrConv(wP_SHORDER_REC.ORDER_CODE, vbUnicode)
    
    
        End If
    
        com = BtOpGetNext
    
    Loop
                

    Call Input_UnLock
    Print_Proc = False

End Function


Private Function List_Print_Proc() As Integer
'----------------------------------------------------------------------------
'           �����������X�g�������
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
                
Dim rpt         As New ODR3010F2
Dim f           As New ODR30103

                
    List_Print_Proc = True
    Call Input_Lock
    
    
    If ODR_KENTO_Open(BtOpenNomal) Then
        MsgBox "�����𒆒f���܂��B", vbExclamation
        Call Input_UnLock           '2016.05.06
        Exit Function
    End If
    
    
    
    Set rpt = New ODR3010F2
        
    '���|�[�g��������܂��B�itrue�F����_�C�A���O���� false�F�Ȃ��j
    'rpt.PrintReport False
    rpt.PrintReport True
    
    Set rpt = Nothing
    
    
    sts = BTRV(BtOpClose, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_KENTO")
        End If
    End If
    
    
    
    Call Input_UnLock
    List_Print_Proc = False

End Function

