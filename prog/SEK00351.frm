VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form SEK00351 
   Caption         =   "�ϐ��n�E�X�@�ʒ����ް��@�����ƍ�"
   ClientHeight    =   4920
   ClientLeft      =   2025
   ClientTop       =   -3210
   ClientWidth     =   8745
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  '�蓮
   ScaleHeight     =   4920
   ScaleWidth      =   8745
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton Command1 
      Caption         =   "INI�\��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   16680
      TabIndex        =   10
      ToolTipText     =   "�������I�����܂�"
      Top             =   240
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�I ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   3600
      TabIndex        =   9
      ToolTipText     =   "�������I�����܂�"
      Top             =   360
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�o ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   5160
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�o�דo�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�\ ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   6
      ToolTipText     =   "�ŐV����\�����܂�"
      Top             =   360
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   372
      Left            =   16680
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�@�ʓo�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1380
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   972
      Left            =   240
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   1720
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "�폜"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "��������"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "��ݾ�"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "�o�ׁ@�@�@�`�[��"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "�o�ׁ@�@�@�@�h�c��"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "�@�f�[�^�@�@�@�쐬��"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "�@�f�[�^�@�@�@�쐬����"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "�A��"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "�󒍓�"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "�@�@�[���@�@����ꏊ"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "�@�@�[���@�@�@�@�@����ꏊ��"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "���Ӑ�@�@�@�R�[�h"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "���[��@�@�@�R�[�h"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "���Ӑ�i�ԁ@�@���i�ԁi��j"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "�i�ԁ@�@�@�@�@���i�ԁi���j"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "�������@�@�@�@�@�@���w�}���i��j"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "�o�׏��ԁ@�@�@�@�@�@���w�}���i���E���j"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "�@���@�@�@�@�@�@�@���w�}���i���E�E�j"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "�󒍐���"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "�o�׊m���"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "�[����"
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "���Ǉ��@�@�@�@�@�@���Ǘ����i��j"
      Columns(21).DataField=   ""
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "�i�Ǉ��@�@�@�@�@���Ǘ����i���j"
      Columns(22).DataField=   ""
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "�P�i�敪"
      Columns(23).DataField=   ""
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "�@������ID"
      Columns(24).DataField=   ""
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).Caption=   "����"
      Columns(25).DataField=   ""
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(26)._VlistStyle=   0
      Columns(26)._MaxComboItems=   5
      Columns(26).Caption=   "���o�ɐ�"
      Columns(26).DataField=   ""
      Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(27)._VlistStyle=   0
      Columns(27)._MaxComboItems=   5
      Columns(27).Caption=   "�o�Ɂ@       �S����"
      Columns(27).DataField=   ""
      Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(28)._VlistStyle=   0
      Columns(28)._MaxComboItems=   5
      Columns(28).Caption=   "�o�Ɂ@�@�@�@�@����"
      Columns(28).DataField=   ""
      Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(29)._VlistStyle=   0
      Columns(29)._MaxComboItems=   5
      Columns(29).Caption=   "����@�@�@�S����"
      Columns(29).DataField=   ""
      Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(30)._VlistStyle=   0
      Columns(30)._MaxComboItems=   5
      Columns(30).Caption=   "����@�@�@�@�@����"
      Columns(30).DataField=   ""
      Columns(30)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(31)._VlistStyle=   0
      Columns(31)._MaxComboItems=   5
      Columns(31).Caption=   "����ID"
      Columns(31).DataField=   ""
      Columns(31)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(32)._VlistStyle=   0
      Columns(32)._MaxComboItems=   5
      Columns(32).Caption=   "�W���@�@�@�S����"
      Columns(32).DataField=   ""
      Columns(32)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(33)._VlistStyle=   0
      Columns(33)._MaxComboItems=   5
      Columns(33).Caption=   "�W���@�@�@�@�@����"
      Columns(33).DataField=   ""
      Columns(33)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(34)._VlistStyle=   0
      Columns(34)._MaxComboItems=   5
      Columns(34).Caption=   "����"
      Columns(34).DataField=   ""
      Columns(34)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(35)._VlistStyle=   0
      Columns(35)._MaxComboItems=   5
      Columns(35).Caption=   "�ː�"
      Columns(35).DataField=   ""
      Columns(35)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(36)._VlistStyle=   0
      Columns(36)._MaxComboItems=   5
      Columns(36).Caption=   "�}��"
      Columns(36).DataField=   ""
      Columns(36)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(37)._VlistStyle=   0
      Columns(37)._MaxComboItems=   5
      Columns(37).Caption=   "�ƍ��@�@�@�S����"
      Columns(37).DataField=   ""
      Columns(37)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(38)._VlistStyle=   0
      Columns(38)._MaxComboItems=   5
      Columns(38).Caption=   "�ƍ� �@�@�@����"
      Columns(38).DataField=   ""
      Columns(38)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(39)._VlistStyle=   0
      Columns(39)._MaxComboItems=   5
      Columns(39).Caption=   "���i�@�@�@�@�S����"
      Columns(39).DataField=   ""
      Columns(39)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(40)._VlistStyle=   0
      Columns(40)._MaxComboItems=   5
      Columns(40).Caption=   "���i�@�@�@�@�@����"
      Columns(40).DataField=   ""
      Columns(40)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(41)._VlistStyle=   0
      Columns(41)._MaxComboItems=   5
      Columns(41).Caption=   "�ǉ��@�@�@�S����"
      Columns(41).DataField=   ""
      Columns(41)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(42)._VlistStyle=   0
      Columns(42)._MaxComboItems=   5
      Columns(42).Caption=   "�ǉ��@�@�@�@�@����"
      Columns(42).DataField=   ""
      Columns(42)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(43)._VlistStyle=   0
      Columns(43)._MaxComboItems=   5
      Columns(43).Caption=   "�X�V�@�@�@�S����"
      Columns(43).DataField=   ""
      Columns(43)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(44)._VlistStyle=   0
      Columns(44)._MaxComboItems=   5
      Columns(44).Caption=   "�X�V�@�@�@�@�@����"
      Columns(44).DataField=   ""
      Columns(44)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   45
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   714
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=45"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=556"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=450"
      Splits(0)._ColumnProps(4)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=5768"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5662"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1111"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1005"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=1958"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1852"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=2487"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=2381"
      Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(24)=   "Column(5).Width=2381"
      Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=2275"
      Splits(0)._ColumnProps(27)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(28)=   "Column(6).Width=2381"
      Splits(0)._ColumnProps(29)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(6)._WidthInPix=2275"
      Splits(0)._ColumnProps(31)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(32)=   "Column(7).Width=1376"
      Splits(0)._ColumnProps(33)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(7)._WidthInPix=1270"
      Splits(0)._ColumnProps(35)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(36)=   "Column(8).Width=1588"
      Splits(0)._ColumnProps(37)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(8)._WidthInPix=1482"
      Splits(0)._ColumnProps(39)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(40)=   "Column(9).Width=2064"
      Splits(0)._ColumnProps(41)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(9)._WidthInPix=1958"
      Splits(0)._ColumnProps(43)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(44)=   "Column(10).Width=2831"
      Splits(0)._ColumnProps(45)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(10)._WidthInPix=2725"
      Splits(0)._ColumnProps(47)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(48)=   "Column(11).Width=1958"
      Splits(0)._ColumnProps(49)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(11)._WidthInPix=1852"
      Splits(0)._ColumnProps(51)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(52)=   "Column(12).Width=1958"
      Splits(0)._ColumnProps(53)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(12)._WidthInPix=1852"
      Splits(0)._ColumnProps(55)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(56)=   "Column(13).Width=2593"
      Splits(0)._ColumnProps(57)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(13)._WidthInPix=2487"
      Splits(0)._ColumnProps(59)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(60)=   "Column(14).Width=2593"
      Splits(0)._ColumnProps(61)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(14)._WidthInPix=2487"
      Splits(0)._ColumnProps(63)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(64)=   "Column(15).Width=3281"
      Splits(0)._ColumnProps(65)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(66)=   "Column(15)._WidthInPix=3175"
      Splits(0)._ColumnProps(67)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(68)=   "Column(16).Width=3545"
      Splits(0)._ColumnProps(69)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(70)=   "Column(16)._WidthInPix=3440"
      Splits(0)._ColumnProps(71)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(72)=   "Column(17).Width=3334"
      Splits(0)._ColumnProps(73)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(74)=   "Column(17)._WidthInPix=3228"
      Splits(0)._ColumnProps(75)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(76)=   "Column(18).Width=2752"
      Splits(0)._ColumnProps(77)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(78)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(79)=   "Column(18)._ColStyle=2"
      Splits(0)._ColumnProps(80)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(81)=   "Column(19).Width=2514"
      Splits(0)._ColumnProps(82)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(83)=   "Column(19)._WidthInPix=2408"
      Splits(0)._ColumnProps(84)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(85)=   "Column(20).Width=2064"
      Splits(0)._ColumnProps(86)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(87)=   "Column(20)._WidthInPix=1958"
      Splits(0)._ColumnProps(88)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(89)=   "Column(21).Width=3254"
      Splits(0)._ColumnProps(90)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(21)._WidthInPix=3149"
      Splits(0)._ColumnProps(92)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(93)=   "Column(22).Width=2858"
      Splits(0)._ColumnProps(94)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(95)=   "Column(22)._WidthInPix=2752"
      Splits(0)._ColumnProps(96)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(97)=   "Column(23).Width=1296"
      Splits(0)._ColumnProps(98)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(99)=   "Column(23)._WidthInPix=1191"
      Splits(0)._ColumnProps(100)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(101)=   "Column(24).Width=2725"
      Splits(0)._ColumnProps(102)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(103)=   "Column(24)._WidthInPix=2619"
      Splits(0)._ColumnProps(104)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(105)=   "Column(25).Width=1508"
      Splits(0)._ColumnProps(106)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(107)=   "Column(25)._WidthInPix=1402"
      Splits(0)._ColumnProps(108)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(109)=   "Column(26).Width=2752"
      Splits(0)._ColumnProps(110)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(111)=   "Column(26)._WidthInPix=2646"
      Splits(0)._ColumnProps(112)=   "Column(26)._ColStyle=2"
      Splits(0)._ColumnProps(113)=   "Column(26).Visible=0"
      Splits(0)._ColumnProps(114)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(115)=   "Column(27).Width=1773"
      Splits(0)._ColumnProps(116)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(117)=   "Column(27)._WidthInPix=1667"
      Splits(0)._ColumnProps(118)=   "Column(27).Visible=0"
      Splits(0)._ColumnProps(119)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(120)=   "Column(28).Width=2328"
      Splits(0)._ColumnProps(121)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(122)=   "Column(28)._WidthInPix=2223"
      Splits(0)._ColumnProps(123)=   "Column(28).Visible=0"
      Splits(0)._ColumnProps(124)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(125)=   "Column(29).Width=1773"
      Splits(0)._ColumnProps(126)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(127)=   "Column(29)._WidthInPix=1667"
      Splits(0)._ColumnProps(128)=   "Column(29).Order=30"
      Splits(0)._ColumnProps(129)=   "Column(30).Width=2328"
      Splits(0)._ColumnProps(130)=   "Column(30).DividerColor=0"
      Splits(0)._ColumnProps(131)=   "Column(30)._WidthInPix=2223"
      Splits(0)._ColumnProps(132)=   "Column(30).Order=31"
      Splits(0)._ColumnProps(133)=   "Column(31).Width=3493"
      Splits(0)._ColumnProps(134)=   "Column(31).DividerColor=0"
      Splits(0)._ColumnProps(135)=   "Column(31)._WidthInPix=3387"
      Splits(0)._ColumnProps(136)=   "Column(31).Order=32"
      Splits(0)._ColumnProps(137)=   "Column(32).Width=1773"
      Splits(0)._ColumnProps(138)=   "Column(32).DividerColor=0"
      Splits(0)._ColumnProps(139)=   "Column(32)._WidthInPix=1667"
      Splits(0)._ColumnProps(140)=   "Column(32).Order=33"
      Splits(0)._ColumnProps(141)=   "Column(33).Width=2328"
      Splits(0)._ColumnProps(142)=   "Column(33).DividerColor=0"
      Splits(0)._ColumnProps(143)=   "Column(33)._WidthInPix=2223"
      Splits(0)._ColumnProps(144)=   "Column(33).Order=34"
      Splits(0)._ColumnProps(145)=   "Column(34).Width=1640"
      Splits(0)._ColumnProps(146)=   "Column(34).DividerColor=0"
      Splits(0)._ColumnProps(147)=   "Column(34)._WidthInPix=1535"
      Splits(0)._ColumnProps(148)=   "Column(34)._ColStyle=2"
      Splits(0)._ColumnProps(149)=   "Column(34).Order=35"
      Splits(0)._ColumnProps(150)=   "Column(35).Width=1640"
      Splits(0)._ColumnProps(151)=   "Column(35).DividerColor=0"
      Splits(0)._ColumnProps(152)=   "Column(35)._WidthInPix=1535"
      Splits(0)._ColumnProps(153)=   "Column(35)._ColStyle=2"
      Splits(0)._ColumnProps(154)=   "Column(35).Order=36"
      Splits(0)._ColumnProps(155)=   "Column(36).Width=3493"
      Splits(0)._ColumnProps(156)=   "Column(36).DividerColor=0"
      Splits(0)._ColumnProps(157)=   "Column(36)._WidthInPix=3387"
      Splits(0)._ColumnProps(158)=   "Column(36)._ColStyle=1"
      Splits(0)._ColumnProps(159)=   "Column(36).Visible=0"
      Splits(0)._ColumnProps(160)=   "Column(36).Order=37"
      Splits(0)._ColumnProps(161)=   "Column(37).Width=1773"
      Splits(0)._ColumnProps(162)=   "Column(37).DividerColor=0"
      Splits(0)._ColumnProps(163)=   "Column(37)._WidthInPix=1667"
      Splits(0)._ColumnProps(164)=   "Column(37).Order=38"
      Splits(0)._ColumnProps(165)=   "Column(38).Width=1984"
      Splits(0)._ColumnProps(166)=   "Column(38).DividerColor=0"
      Splits(0)._ColumnProps(167)=   "Column(38)._WidthInPix=1879"
      Splits(0)._ColumnProps(168)=   "Column(38).Order=39"
      Splits(0)._ColumnProps(169)=   "Column(39).Width=1905"
      Splits(0)._ColumnProps(170)=   "Column(39).DividerColor=0"
      Splits(0)._ColumnProps(171)=   "Column(39)._WidthInPix=1799"
      Splits(0)._ColumnProps(172)=   "Column(39).Order=40"
      Splits(0)._ColumnProps(173)=   "Column(40).Width=2328"
      Splits(0)._ColumnProps(174)=   "Column(40).DividerColor=0"
      Splits(0)._ColumnProps(175)=   "Column(40)._WidthInPix=2223"
      Splits(0)._ColumnProps(176)=   "Column(40).Order=41"
      Splits(0)._ColumnProps(177)=   "Column(41).Width=1773"
      Splits(0)._ColumnProps(178)=   "Column(41).DividerColor=0"
      Splits(0)._ColumnProps(179)=   "Column(41)._WidthInPix=1667"
      Splits(0)._ColumnProps(180)=   "Column(41).Order=42"
      Splits(0)._ColumnProps(181)=   "Column(42).Width=2328"
      Splits(0)._ColumnProps(182)=   "Column(42).DividerColor=0"
      Splits(0)._ColumnProps(183)=   "Column(42)._WidthInPix=2223"
      Splits(0)._ColumnProps(184)=   "Column(42).Order=43"
      Splits(0)._ColumnProps(185)=   "Column(43).Width=1773"
      Splits(0)._ColumnProps(186)=   "Column(43).DividerColor=0"
      Splits(0)._ColumnProps(187)=   "Column(43)._WidthInPix=1667"
      Splits(0)._ColumnProps(188)=   "Column(43).Order=44"
      Splits(0)._ColumnProps(189)=   "Column(44).Width=2328"
      Splits(0)._ColumnProps(190)=   "Column(44).DividerColor=0"
      Splits(0)._ColumnProps(191)=   "Column(44)._WidthInPix=2223"
      Splits(0)._ColumnProps(192)=   "Column(44).Order=45"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �S�V�b�N"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �S�V�b�N"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      Appearance      =   0
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      OLEDropMode     =   1
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
      _StyleDefs(5)   =   ":id=0,.fontname=�l�r �S�V�b�N"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=�l�r �S�V�b�N"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=�l�r �S�V�b�N"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=�l�r �S�V�b�N"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFFF00&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFFFFFF&"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&HFFFF00&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(27)  =   ":id=14,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(28)  =   ":id=14,.fontname=�l�r �S�V�b�N"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFFFF&"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFF00&"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=162,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=159,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=160,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=161,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=214,.parent=13,.alignment=2"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=211,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=212,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=213,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=170,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=167,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=168,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=169,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=166,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=163,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=164,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=165,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=32,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=46,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=50,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=110,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=107,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=108,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=109,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=54,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=58,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=55,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=56,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=57,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=62,.parent=13,.alignment=3"
      _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=59,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=60,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=61,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(12).Style:id=66,.parent=13,.alignment=3"
      _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=63,.parent=14"
      _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=64,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=65,.parent=17"
      _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=70,.parent=13"
      _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=67,.parent=14"
      _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=68,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=69,.parent=17"
      _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=74,.parent=13"
      _StyleDefs(95)  =   "Splits(0).Columns(14).HeadingStyle:id=71,.parent=14"
      _StyleDefs(96)  =   "Splits(0).Columns(14).FooterStyle:id=72,.parent=15"
      _StyleDefs(97)  =   "Splits(0).Columns(14).EditorStyle:id=73,.parent=17"
      _StyleDefs(98)  =   "Splits(0).Columns(15).Style:id=78,.parent=13"
      _StyleDefs(99)  =   "Splits(0).Columns(15).HeadingStyle:id=75,.parent=14"
      _StyleDefs(100) =   "Splits(0).Columns(15).FooterStyle:id=76,.parent=15"
      _StyleDefs(101) =   "Splits(0).Columns(15).EditorStyle:id=77,.parent=17"
      _StyleDefs(102) =   "Splits(0).Columns(16).Style:id=82,.parent=13"
      _StyleDefs(103) =   "Splits(0).Columns(16).HeadingStyle:id=79,.parent=14"
      _StyleDefs(104) =   "Splits(0).Columns(16).FooterStyle:id=80,.parent=15"
      _StyleDefs(105) =   "Splits(0).Columns(16).EditorStyle:id=81,.parent=17"
      _StyleDefs(106) =   "Splits(0).Columns(17).Style:id=86,.parent=13"
      _StyleDefs(107) =   "Splits(0).Columns(17).HeadingStyle:id=83,.parent=14"
      _StyleDefs(108) =   "Splits(0).Columns(17).FooterStyle:id=84,.parent=15"
      _StyleDefs(109) =   "Splits(0).Columns(17).EditorStyle:id=85,.parent=17"
      _StyleDefs(110) =   "Splits(0).Columns(18).Style:id=90,.parent=13,.alignment=1"
      _StyleDefs(111) =   "Splits(0).Columns(18).HeadingStyle:id=87,.parent=14"
      _StyleDefs(112) =   "Splits(0).Columns(18).FooterStyle:id=88,.parent=15"
      _StyleDefs(113) =   "Splits(0).Columns(18).EditorStyle:id=89,.parent=17"
      _StyleDefs(114) =   "Splits(0).Columns(19).Style:id=94,.parent=13"
      _StyleDefs(115) =   "Splits(0).Columns(19).HeadingStyle:id=91,.parent=14"
      _StyleDefs(116) =   "Splits(0).Columns(19).FooterStyle:id=92,.parent=15"
      _StyleDefs(117) =   "Splits(0).Columns(19).EditorStyle:id=93,.parent=17"
      _StyleDefs(118) =   "Splits(0).Columns(20).Style:id=98,.parent=13"
      _StyleDefs(119) =   "Splits(0).Columns(20).HeadingStyle:id=95,.parent=14"
      _StyleDefs(120) =   "Splits(0).Columns(20).FooterStyle:id=96,.parent=15"
      _StyleDefs(121) =   "Splits(0).Columns(20).EditorStyle:id=97,.parent=17"
      _StyleDefs(122) =   "Splits(0).Columns(21).Style:id=102,.parent=13"
      _StyleDefs(123) =   "Splits(0).Columns(21).HeadingStyle:id=99,.parent=14"
      _StyleDefs(124) =   "Splits(0).Columns(21).FooterStyle:id=100,.parent=15"
      _StyleDefs(125) =   "Splits(0).Columns(21).EditorStyle:id=101,.parent=17"
      _StyleDefs(126) =   "Splits(0).Columns(22).Style:id=106,.parent=13"
      _StyleDefs(127) =   "Splits(0).Columns(22).HeadingStyle:id=103,.parent=14"
      _StyleDefs(128) =   "Splits(0).Columns(22).FooterStyle:id=104,.parent=15"
      _StyleDefs(129) =   "Splits(0).Columns(22).EditorStyle:id=105,.parent=17"
      _StyleDefs(130) =   "Splits(0).Columns(23).Style:id=114,.parent=13"
      _StyleDefs(131) =   "Splits(0).Columns(23).HeadingStyle:id=111,.parent=14"
      _StyleDefs(132) =   "Splits(0).Columns(23).FooterStyle:id=112,.parent=15"
      _StyleDefs(133) =   "Splits(0).Columns(23).EditorStyle:id=113,.parent=17"
      _StyleDefs(134) =   "Splits(0).Columns(24).Style:id=118,.parent=13"
      _StyleDefs(135) =   "Splits(0).Columns(24).HeadingStyle:id=115,.parent=14"
      _StyleDefs(136) =   "Splits(0).Columns(24).FooterStyle:id=116,.parent=15"
      _StyleDefs(137) =   "Splits(0).Columns(24).EditorStyle:id=117,.parent=17"
      _StyleDefs(138) =   "Splits(0).Columns(25).Style:id=122,.parent=13"
      _StyleDefs(139) =   "Splits(0).Columns(25).HeadingStyle:id=119,.parent=14"
      _StyleDefs(140) =   "Splits(0).Columns(25).FooterStyle:id=120,.parent=15"
      _StyleDefs(141) =   "Splits(0).Columns(25).EditorStyle:id=121,.parent=17"
      _StyleDefs(142) =   "Splits(0).Columns(26).Style:id=126,.parent=13,.alignment=1"
      _StyleDefs(143) =   "Splits(0).Columns(26).HeadingStyle:id=123,.parent=14"
      _StyleDefs(144) =   "Splits(0).Columns(26).FooterStyle:id=124,.parent=15"
      _StyleDefs(145) =   "Splits(0).Columns(26).EditorStyle:id=125,.parent=17"
      _StyleDefs(146) =   "Splits(0).Columns(27).Style:id=130,.parent=13"
      _StyleDefs(147) =   "Splits(0).Columns(27).HeadingStyle:id=127,.parent=14"
      _StyleDefs(148) =   "Splits(0).Columns(27).FooterStyle:id=128,.parent=15"
      _StyleDefs(149) =   "Splits(0).Columns(27).EditorStyle:id=129,.parent=17"
      _StyleDefs(150) =   "Splits(0).Columns(28).Style:id=134,.parent=13"
      _StyleDefs(151) =   "Splits(0).Columns(28).HeadingStyle:id=131,.parent=14"
      _StyleDefs(152) =   "Splits(0).Columns(28).FooterStyle:id=132,.parent=15"
      _StyleDefs(153) =   "Splits(0).Columns(28).EditorStyle:id=133,.parent=17"
      _StyleDefs(154) =   "Splits(0).Columns(29).Style:id=138,.parent=13"
      _StyleDefs(155) =   "Splits(0).Columns(29).HeadingStyle:id=135,.parent=14"
      _StyleDefs(156) =   "Splits(0).Columns(29).FooterStyle:id=136,.parent=15"
      _StyleDefs(157) =   "Splits(0).Columns(29).EditorStyle:id=137,.parent=17"
      _StyleDefs(158) =   "Splits(0).Columns(30).Style:id=142,.parent=13"
      _StyleDefs(159) =   "Splits(0).Columns(30).HeadingStyle:id=139,.parent=14"
      _StyleDefs(160) =   "Splits(0).Columns(30).FooterStyle:id=140,.parent=15"
      _StyleDefs(161) =   "Splits(0).Columns(30).EditorStyle:id=141,.parent=17"
      _StyleDefs(162) =   "Splits(0).Columns(31).Style:id=190,.parent=13"
      _StyleDefs(163) =   "Splits(0).Columns(31).HeadingStyle:id=187,.parent=14"
      _StyleDefs(164) =   "Splits(0).Columns(31).FooterStyle:id=188,.parent=15"
      _StyleDefs(165) =   "Splits(0).Columns(31).EditorStyle:id=189,.parent=17"
      _StyleDefs(166) =   "Splits(0).Columns(32).Style:id=202,.parent=13"
      _StyleDefs(167) =   "Splits(0).Columns(32).HeadingStyle:id=199,.parent=14"
      _StyleDefs(168) =   "Splits(0).Columns(32).FooterStyle:id=200,.parent=15"
      _StyleDefs(169) =   "Splits(0).Columns(32).EditorStyle:id=201,.parent=17"
      _StyleDefs(170) =   "Splits(0).Columns(33).Style:id=198,.parent=13"
      _StyleDefs(171) =   "Splits(0).Columns(33).HeadingStyle:id=195,.parent=14"
      _StyleDefs(172) =   "Splits(0).Columns(33).FooterStyle:id=196,.parent=15"
      _StyleDefs(173) =   "Splits(0).Columns(33).EditorStyle:id=197,.parent=17"
      _StyleDefs(174) =   "Splits(0).Columns(34).Style:id=194,.parent=13,.alignment=1"
      _StyleDefs(175) =   "Splits(0).Columns(34).HeadingStyle:id=191,.parent=14"
      _StyleDefs(176) =   "Splits(0).Columns(34).FooterStyle:id=192,.parent=15"
      _StyleDefs(177) =   "Splits(0).Columns(34).EditorStyle:id=193,.parent=17"
      _StyleDefs(178) =   "Splits(0).Columns(35).Style:id=174,.parent=13,.alignment=1"
      _StyleDefs(179) =   "Splits(0).Columns(35).HeadingStyle:id=171,.parent=14"
      _StyleDefs(180) =   "Splits(0).Columns(35).FooterStyle:id=172,.parent=15"
      _StyleDefs(181) =   "Splits(0).Columns(35).EditorStyle:id=173,.parent=17"
      _StyleDefs(182) =   "Splits(0).Columns(36).Style:id=210,.parent=13,.alignment=2"
      _StyleDefs(183) =   "Splits(0).Columns(36).HeadingStyle:id=207,.parent=14"
      _StyleDefs(184) =   "Splits(0).Columns(36).FooterStyle:id=208,.parent=15"
      _StyleDefs(185) =   "Splits(0).Columns(36).EditorStyle:id=209,.parent=17"
      _StyleDefs(186) =   "Splits(0).Columns(37).Style:id=206,.parent=13"
      _StyleDefs(187) =   "Splits(0).Columns(37).HeadingStyle:id=203,.parent=14"
      _StyleDefs(188) =   "Splits(0).Columns(37).FooterStyle:id=204,.parent=15"
      _StyleDefs(189) =   "Splits(0).Columns(37).EditorStyle:id=205,.parent=17"
      _StyleDefs(190) =   "Splits(0).Columns(38).Style:id=178,.parent=13"
      _StyleDefs(191) =   "Splits(0).Columns(38).HeadingStyle:id=175,.parent=14"
      _StyleDefs(192) =   "Splits(0).Columns(38).FooterStyle:id=176,.parent=15"
      _StyleDefs(193) =   "Splits(0).Columns(38).EditorStyle:id=177,.parent=17"
      _StyleDefs(194) =   "Splits(0).Columns(39).Style:id=186,.parent=13"
      _StyleDefs(195) =   "Splits(0).Columns(39).HeadingStyle:id=183,.parent=14"
      _StyleDefs(196) =   "Splits(0).Columns(39).FooterStyle:id=184,.parent=15"
      _StyleDefs(197) =   "Splits(0).Columns(39).EditorStyle:id=185,.parent=17"
      _StyleDefs(198) =   "Splits(0).Columns(40).Style:id=182,.parent=13"
      _StyleDefs(199) =   "Splits(0).Columns(40).HeadingStyle:id=179,.parent=14"
      _StyleDefs(200) =   "Splits(0).Columns(40).FooterStyle:id=180,.parent=15"
      _StyleDefs(201) =   "Splits(0).Columns(40).EditorStyle:id=181,.parent=17"
      _StyleDefs(202) =   "Splits(0).Columns(41).Style:id=146,.parent=13"
      _StyleDefs(203) =   "Splits(0).Columns(41).HeadingStyle:id=143,.parent=14"
      _StyleDefs(204) =   "Splits(0).Columns(41).FooterStyle:id=144,.parent=15"
      _StyleDefs(205) =   "Splits(0).Columns(41).EditorStyle:id=145,.parent=17"
      _StyleDefs(206) =   "Splits(0).Columns(42).Style:id=150,.parent=13"
      _StyleDefs(207) =   "Splits(0).Columns(42).HeadingStyle:id=147,.parent=14"
      _StyleDefs(208) =   "Splits(0).Columns(42).FooterStyle:id=148,.parent=15"
      _StyleDefs(209) =   "Splits(0).Columns(42).EditorStyle:id=149,.parent=17"
      _StyleDefs(210) =   "Splits(0).Columns(43).Style:id=154,.parent=13"
      _StyleDefs(211) =   "Splits(0).Columns(43).HeadingStyle:id=151,.parent=14"
      _StyleDefs(212) =   "Splits(0).Columns(43).FooterStyle:id=152,.parent=15"
      _StyleDefs(213) =   "Splits(0).Columns(43).EditorStyle:id=153,.parent=17"
      _StyleDefs(214) =   "Splits(0).Columns(44).Style:id=158,.parent=13"
      _StyleDefs(215) =   "Splits(0).Columns(44).HeadingStyle:id=155,.parent=14"
      _StyleDefs(216) =   "Splits(0).Columns(44).FooterStyle:id=156,.parent=15"
      _StyleDefs(217) =   "Splits(0).Columns(44).EditorStyle:id=157,.parent=17"
      _StyleDefs(218) =   "Named:id=33:Normal"
      _StyleDefs(219) =   ":id=33,.parent=0"
      _StyleDefs(220) =   "Named:id=34:Heading"
      _StyleDefs(221) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(222) =   ":id=34,.wraptext=-1"
      _StyleDefs(223) =   "Named:id=35:Footing"
      _StyleDefs(224) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(225) =   "Named:id=36:Selected"
      _StyleDefs(226) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(227) =   "Named:id=37:Caption"
      _StyleDefs(228) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(229) =   "Named:id=38:HighlightRow"
      _StyleDefs(230) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(231) =   "Named:id=39:EvenRow"
      _StyleDefs(232) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(233) =   "Named:id=40:OddRow"
      _StyleDefs(234) =   ":id=40,.parent=33"
      _StyleDefs(235) =   "Named:id=41:RecordSelector"
      _StyleDefs(236) =   ":id=41,.parent=34"
      _StyleDefs(237) =   "Named:id=42:FilterBar"
      _StyleDefs(238) =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�����쐬"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3600
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2088
      TabIndex        =   0
      ToolTipText     =   "�݌Ɏ������������s���܂�"
      Top             =   360
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label2 
      Caption         =   "�ϐ��n�E�X�@�ʒ����f�[�^�@�ƍ��@�@�@�@�@�@�@���s��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   840
      TabIndex        =   11
      Top             =   1680
      Width           =   5412
   End
   Begin VB.Label Label1 
      Caption         =   "�Ǎ��݌���"
      Height          =   252
      Left            =   15360
      TabIndex        =   5
      Top             =   840
      Width           =   1212
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "�����I��"
      Begin VB.Menu SHORI 
         Caption         =   "�\��"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "�ƍ�"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "�I��"
         Index           =   5
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "SEK00351"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Y_Syuka_TEI     As New XArrayDB

Private Const Min_Row% = 1              '�ŏ��s��

Private Max_Row    As Integer           '�O���b�h�ő�\������


Private Const Min_Col% = 0              '�ŏ���
Private Const Max_Col% = 44             '�ő��

Private Const colDEL_FLG% = 0           '�폜�׸�
Private Const colSHORI% = 1             '��������

Private Const colCANCEL_F% = 2          '��ݾ�


Private Const colDEN_NO% = 3            '�o�ח\��`�[��
Private Const colID_NO% = 4             '�o�ח\��ID��

Private Const colSND_YMD% = 5           '�f�[�^�쐬��
Private Const colSND_HMS% = 6           '�f�[�^�쐬����
Private Const colSEQ_NO% = 7            '�A��
Private Const colJUC_YMD% = 8           '�󒍓�
Private Const colNOU_CD% = 9            '�[�������
Private Const colNOU_NM% = 10            '�[������ꖼ
Private Const colTOK_CD% = 11           '���Ӑ溰��
Private Const colCHO_CD% = 12           '���[�溰��
Private Const colTHINB_CD% = 13         '���Ӑ�i�ԁ@���i��(��)
Private Const colHINB_CD% = 14          '�i�ԁ@      ���i��(��)
Private Const colCHU_CD% = 15           '�������@    ���w�}��(��)
Private Const colSYU_JUN% = 16          '�o�׏��ԁ@  ���w�}��(���E��)
Private Const colTEI_NM% = 17           '�@���@      ���w�}��(���E�E)
Private Const colJUC_SUU% = 18          '�󒍐���
Private Const colSYU_YMD% = 19          '�o�׊m���
Private Const colNOU_YMD% = 20          '�[����
Private Const colKEN_NO% = 21           '���Ǉ��@�@�@���Ǘ���(��)
Private Const colHIN_NO% = 22           '���Ǉ��@�@�@���Ǘ���(��)
Private Const colTANP_KB% = 23          '�P�i�敪

Private Const colTEI_LABELID% = 24      '�@������ID
Private Const colHAKO_NO% = 25          '����



Private Const colJITU_SUU% = 26         '���o�ɐ�
Private Const colJITU_TANTO% = 27       '�o�Ɂ@�S����
Private Const colJITU_DATETIME% = 28    '�o�Ɂ@����


Private Const colKONPO_TANTO% = 29      '����@�S����
Private Const colKONPO_DATETIME% = 30   '����@����

Private Const colKONPO_ID% = 31         '����ID

Private Const colSYUGO_KONPO_TANTO% = 32    '�W������@�S����
Private Const colSYUGO_KONPO_DATETIME% = 33 '�W������@����



Private Const colKUTI_SU% = 34          '����
Private Const colSAI_SU% = 35           '�ː�

Private Const colOKURI_NO_SEQ% = 36     '�}��


Private Const colSHOGO_TANTO% = 37      '�ƍ��@�S����
Private Const colSHOGO_DATETIME% = 38   '�ƍ��@����

Private Const colKENPIN_TANTO% = 39     '���i�@�S����
Private Const colKENPIN_DATETIME% = 40  '���i�@����

Private Const colINS_TANTO% = 41        '�ǉ��@�S����
Private Const colINS_DateTime% = 42     '�ǉ��@����
Private Const colUPD_TANTO% = 43        '�X�V�@�S����
Private Const colUPD_DATETIME% = 44     '�X�V�@����


Private Sort_Tbl(Min_Col To colUPD_DATETIME) _
            As Integer                  '��Ă̐��� 0:���� 1:�~��





Private Zaiko_Tanaban   As String * 8   '�݌Ɉ������p�I��

Private MENU_NO         As String * 2   '���j���[��

Private YOIN_CODE       As String * 2   '�v���R�[�h

Private WS_NO           As String * 3   'ܰ��ð��ݔԍ�

Private Disp_Mode       As Integer      '���ו\��Ӱ��

Private KONPO_F         As Integer      '��������ƍ��Ώۂ��邩�H

Private SYUGO_KONPO_F   As Integer      '���W��������ƍ��Ώۂ��邩�H

Private KENPIN_F        As Integer      '���i�ς݂�ΏۂƂ��邩�H

Private ZAIKO_F         As Integer      '�݌ɏ󋵂��`�F�b�N���邩�H



Private SEK_MUKE_CODE  As Variant   '���������溰��


Dim wkDisp_Mode     As Integer
Dim wkKONPO_F       As Integer
Dim wkSYUGO_KONPO_F As Integer
Dim wkKENPIN_F      As Integer
Dim wkZAIKO_F       As Integer


'Private Const LAST_UPDATE_DAY$ = "[SEK0035] 2018.04.06 13:30"
Private Const LAST_UPDATE_DAY$ = "[SEK0035] 2018.04.06 16:15"
Private Sub Command1_Click(Index As Integer)
    
    
    Select Case Index
        
        
        Case 0          '�\��
        
            If List_Disp_Proc() Then
                Unload Me
            End If
        
        
            Command1(Index).SetFocus
        
        
        Case 1          '�ƍ�
        
        
            If Matching_Proc() Then
                Unload Me
            End If
        
            Command1(Index).SetFocus
        
        Case 2          '�@�ʓo�^
        
            If TEI_Update_Proc() Then
                Unload Me
            End If
        
            Command1(Index).SetFocus
        
        
        Case 3          '�\��o�^
        
            If YOTEI_Update_Proc() Then
                Unload Me
            End If
        
            Command1(Index).SetFocus
        
        
        Case 4          '�����쐬
        
        
            If Data_Link_Proc() Then
                Unload Me
            End If
        
            Command1(Index).SetFocus
        
        
        Case 5          'CSV�쐬
        
        
            If CSV_DATA_OUT_PROC() Then
                Unload Me
            End If
        
            Command1(Index).SetFocus
        
        
        
        Case 6          '�I��
            
            Unload Me
    
    
        Case 7          'INI�\��
    
    
    
    
    
    
            MsgBox "Zaiko_Tanaban=" & Zaiko_Tanaban & Chr(13) & Chr(10) & _
                    "DISP_MODE=" & wkDisp_Mode & Chr(13) & Chr(10) & _
                    "KONPO_F=" & wkKONPO_F & Chr(13) & Chr(10) & _
                    "SYUGO_KONPO_F=" & wkSYUGO_KONPO_F & Chr(13) & Chr(10) & _
                    "KENPIN_F=" & wkKENPIN_F & Chr(13) & Chr(10) & _
                    "ZAIKO_F=" & wkZAIKO_F
    
    
    
    
    End Select
    
    
    
    
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �j���� �c������ �O����
'----------------------------------------------------------------------------
    
    
    If Shift = vbAltMask Then
    
    
        If TDBGrid1.Columns(colDEL_FLG).Visible Then
            TDBGrid1.Columns(colDEL_FLG).Visible = False
        Else
            TDBGrid1.Columns(colDEL_FLG).Visible = True
        End If

    
        If Command1(2).Visible Then
            Command1(2).Visible = False
        Else
            Command1(2).Visible = True
        End If
    
    
        If Command1(3).Visible Then
            Command1(3).Visible = False
        Else
            Command1(3).Visible = True
        End If
    
    End If
    
    
    If Shift = vbAltMask + vbCtrlMask Then
    
        If Command1(4).Visible Then
            Command1(4).Visible = False
        Else
            Command1(4).Visible = True
        End If
    
        If Command1(5).Visible Then
            Command1(5).Visible = False
        Else
            Command1(5).Visible = True
        End If
    
    
    
    End If
    
    
    If Shift = vbShiftMask Then
        If TDBGrid1.Columns(35).Visible Then
            TDBGrid1.Columns(35).Visible = False
        Else
            TDBGrid1.Columns(35).Visible = True
    
        End If
    End If
    
    Select Case KeyCode
        Case vbKeyF12
            Unload Me
    End Select

End Sub

Private Sub Form_Load()


Dim c       As String * 128

Dim i       As Integer

Dim sBuffer As String * 255
Dim com     As String
    
    
    
    
    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�ϐ��n�E�X�@�ʒ����ް��@�ƍ�", Me.hwnd, 0)
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)



    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)


    '�݌ɗ��p�I�Ԋl��
    If GetIni(App.EXEName, "Zaiko_Tanaban", App.EXEName, c) Then
        Beep
        MsgBox "�݌Ɉ����p�I�Ԃ��w�肵�ĉ�����[SEK0035] Zaiko_Tanaban="
        End
    End If
    Zaiko_Tanaban = RTrim(c)

    '�o�ɗv���l��
    If GetIni(App.EXEName, "YOIN", App.EXEName, c) Then
        Beep
        MsgBox "�o�ɗv�����w�肵�ĉ�����[SEK0035] YOIN="
        End
    End If
    YOIN_CODE = RTrim(c)

    '�ϐ������溰��
    If GetIni(App.EXEName, "MUKE_CODE", App.EXEName, c) Then
        Beep
        MsgBox "�ϐ������溰�ނ��w�肵�ĉ�����[SEK0035] MUKE_CODE="
        End
    End If
    SEK_MUKE_CODE = Split(Trim(c), ",", -1)
    
    



    '�\��Ӱ��
    If GetIni(App.EXEName, "DISP_MODE", App.EXEName, c) Then
        Disp_Mode = False
        wkDisp_Mode = 0
    Else
        If Trim(c) = "1" Then
            Disp_Mode = True
            wkDisp_Mode = 1
        Else
            Disp_Mode = False
            wkDisp_Mode = 0
        End If
    End If


    '��������ƍ��Ώۂ��邩�H
    If GetIni(App.EXEName, "KONPO_F", App.EXEName, c) Then
        KONPO_F = True
        wkKONPO_F = 0
    Else
        If Trim(c) = "1" Then
            KONPO_F = False
            wkKONPO_F = 1
        Else
            KONPO_F = True
            wkKONPO_F = 0
        End If
    End If


    '���W��������ƍ��Ώۂ��邩�H
    If GetIni(App.EXEName, "SYUGO_KONPO_F", App.EXEName, c) Then
        SYUGO_KONPO_F = True
        wkSYUGO_KONPO_F = 0
    Else
        If Trim(c) = "1" Then
            SYUGO_KONPO_F = False
            wkSYUGO_KONPO_F = 1
        Else
            SYUGO_KONPO_F = True
            wkSYUGO_KONPO_F = 0
        End If
    End If
    '���i�ς݂��ƍ��Ώۂ��邩�H
    If GetIni(App.EXEName, "KENPIN_F", App.EXEName, c) Then
        KENPIN_F = 0
        wkKENPIN_F = 0
    Else
        If Trim(c) = "0" Or Trim(c) = "1" Or Trim(c) = "2" Then
            KENPIN_F = Val(Trim(c))
            wkKENPIN_F = Val(Trim(c))
        Else
            wkKENPIN_F = 0
        End If
    End If
    
    
    
    '�݌ɏ󋵂��`�F�b�N���邩�H
    If GetIni(App.EXEName, "ZAIKO_F", App.EXEName, c) Then
        ZAIKO_F = True
        wkZAIKO_F = 0
    Else
        If Trim(c) = "1" Then
            ZAIKO_F = False
            wkZAIKO_F = 1
        Else
            ZAIKO_F = True
            wkZAIKO_F = 0
        End If
    End If
    
    
    
    
    
    
    
    '�ƭ����l��
    If GetIni(App.EXEName, "MENU_NO", App.EXEName, c) Then
        MENU_NO = ""
    Else
        MENU_NO = Trim(c)
    End If


                                'ܰ��ð��ݔԍ���荞��
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)




    SEK00351.Caption = SEK00351.Caption & " " & LAST_UPDATE_DAY

                                '�@�ʒ����ް��n�o�d�m
    If Y_SYU_TEI_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '�o�ח\���ް��n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '�o�ח\��(H)�ް��n�o�d�m
    If Y_SYU_H_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '���ԗp�ް��n�o�d�m
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '�݌Ɉړ����n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '�i��Ͻ��n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '�i�ڃ}�X�^(���[�N)�n�o�d�m
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If


                                '������Ͻ��n�o�d�m
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '���۸ނn�o�d�m
    If P_SAGYO_LOG_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�S����Ͻ��n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�v��Ͻ��n�o�d�m
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌��ް��n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    '�݌ɂ̗\�����
    If Data_Clear_Proc(0) Then
        Unload Me
    End If
    
    
    '��ď��̏�����
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0             '��̫�ď���
    Next i

'    Command1(0).SetFocus




    If List_Disp_Proc() Then
        Unload Me
    End If
        
        
        
    If Matching_Proc() Then
        Unload Me
    End If

    Unload Me

End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    
Dim sts     As Integer
    
    sts = BTRV(BtOpClose, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�@�ʒ����ް�")
        End If
    End If
    
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\���ް�")
        End If
    End If
    
    sts = BTRV(BtOpClose, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��(H)�ް�")
        End If
    End If
    
    sts = BTRV(BtOpClose, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ԗp�ް�")
        End If
    End If
    
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If
    
    
    
    
    
    
    sts = BTRV(BtOpReset, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If


    Set SEK00351 = Nothing



    End

End Sub

Private Sub SHORI_Click(Index As Integer)

    Select Case Index
    
        Case 0
            
            
            Command1(0).Value = True
        Case 1
            
            Command1(1).Value = True
        Case 2
            If Command1(2).Enabled Then
            
                Command1(2).Value = True
            End If
        Case 6
            Command1(6).Value = True
    End Select



End Sub



Private Function TEI_Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   �u�����f�[�^�v�o�^����
'----------------------------------------------------------------------------

Dim sts             As Integer
Dim ans             As Integer
    
Dim com             As Integer
Dim Upd_com         As Integer

Dim Skip_Flg        As Integer

Dim Row             As Long



    If Y_Syuka_TEI.Count(1) = 0 Then
        Exit Function
    End If
    
    TEI_Update_Proc = True
    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�ϐ��n�E�X�@�ʒ����ް��ƍ��@[�����ް�]�X�V�����J�n�I�I", Me.hwnd, 0)

                                    
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts <> BtNoErr Then
        Call Input_Lock
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        TEI_Update_Proc = SYS_ERR
        Exit Function
    End If
                                    
                                    
                                    
                                    
    Set TDBGrid1.Array = Y_Syuka_TEI
    TDBGrid1.Refresh
    
    TDBGrid1.Update
                                    
                                    
                                    
                                    
                                    
                                    '�e�[�u�����Z�b�g
    For Row = 1 To Y_Syuka_TEI.UpperBound(1)
        
        
        DoEvents
        
        
        Call UniCode_Conv(K0_Y_SYU_TEI.SND_YMD, Y_Syuka_TEI(Row, colSND_YMD))
        Call UniCode_Conv(K0_Y_SYU_TEI.SND_HMS, Y_Syuka_TEI(Row, colSND_HMS))
        Call UniCode_Conv(K0_Y_SYU_TEI.SEQ_NO, Y_Syuka_TEI(Row, colSEQ_NO))
        
        
        Skip_Flg = False
        
        sts = BTRV(BtOpGetEqual, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
        Select Case sts
            Case BtNoErr
            
                Upd_com = BtOpUpdate
            
                If Y_Syuka_TEI(Row, colDEL_FLG) = "1" Then
            
                    Upd_com = BtOpDelete
                            
                End If
            
            Case BtErrKeyNotFound
                Upd_com = BtOpInsert
            
                If Y_Syuka_TEI(Row, colDEL_FLG) = "1" Then
            
                    Skip_Flg = True
                            
                End If
            
            
            Case Else
                Call Input_UnLock
                Call File_Error(sts, BtOpInsert, "�@�ʒ����f�[�^")
                GoTo Abort_Tran
        End Select
        
        
        
        
        If Not Skip_Flg And Upd_com <> BtOpDelete Then
        
        
        
            Call UniCode_Conv(Y_SYU_TEI_REC.SND_YMD, Y_Syuka_TEI(Row, colSND_YMD))          '�f�[�^�쐬��
            Call UniCode_Conv(Y_SYU_TEI_REC.SND_HMS, Y_Syuka_TEI(Row, colSND_HMS))          '�f�[�^�쐬����
            Call UniCode_Conv(Y_SYU_TEI_REC.SEQ_NO, Y_Syuka_TEI(Row, colSEQ_NO))            '�A��
            Call UniCode_Conv(Y_SYU_TEI_REC.JUC_YMD, Y_Syuka_TEI(Row, colJUC_YMD))          '�󒍓�
            Call UniCode_Conv(Y_SYU_TEI_REC.NOU_CD, Y_Syuka_TEI(Row, colNOU_CD))            '�[�������
            Call UniCode_Conv(Y_SYU_TEI_REC.NOU_NM, Y_Syuka_TEI(Row, colNOU_NM))            '�[������ꖼ
            Call UniCode_Conv(Y_SYU_TEI_REC.TOK_CD, Y_Syuka_TEI(Row, colTOK_CD))            '���Ӑ溰��
            Call UniCode_Conv(Y_SYU_TEI_REC.CHO_CD, Y_Syuka_TEI(Row, colCHO_CD))            '���[�溰��
            Call UniCode_Conv(Y_SYU_TEI_REC.THINB_CD, Y_Syuka_TEI(Row, colTHINB_CD))        '���Ӑ�i�ԁ@���i��(��)
            Call UniCode_Conv(Y_SYU_TEI_REC.HINB_CD, Y_Syuka_TEI(Row, colHINB_CD))          '�i�ԁ@���i��(��)
            Call UniCode_Conv(Y_SYU_TEI_REC.CHU_CD, Y_Syuka_TEI(Row, colCHU_CD))            '�������@    ���w�}��(��)
            Call UniCode_Conv(Y_SYU_TEI_REC.SYU_JUN, Y_Syuka_TEI(Row, colSYU_JUN))          '�o�׏��ԁ@  ���w�}��(���E��)
            Call UniCode_Conv(Y_SYU_TEI_REC.TEI_NM, Y_Syuka_TEI(Row, colTEI_NM))            '�@���@      ���w�}��(���E�E)
                                                                                            '�󒍐���
            Call UniCode_Conv(Y_SYU_TEI_REC.JUC_SUU, Format(Val(Y_Syuka_TEI(Row, colJUC_SUU)), "00000000"))
            Call UniCode_Conv(Y_SYU_TEI_REC.SYU_YMD, Y_Syuka_TEI(Row, colSYU_YMD))          '�o�׊m���
            Call UniCode_Conv(Y_SYU_TEI_REC.NOU_YMD, Y_Syuka_TEI(Row, colNOU_YMD))          '�[����
            Call UniCode_Conv(Y_SYU_TEI_REC.KEN_NO, Y_Syuka_TEI(Row, colKEN_NO))            '���Ǉ��@�@�@���Ǘ���(��)
            Call UniCode_Conv(Y_SYU_TEI_REC.HIN_NO, Y_Syuka_TEI(Row, colHIN_NO))            '���Ǉ��@�@�@���Ǘ���(��)
            Call UniCode_Conv(Y_SYU_TEI_REC.TANP_KB, Y_Syuka_TEI(Row, colTANP_KB))          '�P�i�敪
            Call UniCode_Conv(Y_SYU_TEI_REC.YOBI1_NM, "")                                   '�\��
            
            
            If Upd_com = BtOpInsert Then
                Call UniCode_Conv(Y_SYU_TEI_REC.GSEQ_NO, "00000")                           '������
            End If
            
            
            
            
            Call UniCode_Conv(Y_SYU_TEI_REC.TEI_LABELID, Y_Syuka_TEI(Row, colTEI_LABELID))  '�@������ID(���������w�}��(��)+����)
            
            
            Call UniCode_Conv(Y_SYU_TEI_REC.HAKO_NO, Y_Syuka_TEI(Row, colHAKO_NO))          '����
            
            
            Call UniCode_Conv(Y_SYU_TEI_REC.KONPO_ID, Y_Syuka_TEI(Row, colKONPO_ID))        '����ID
            
            Call UniCode_Conv(Y_SYU_TEI_REC.KONPO_TANTO, Y_Syuka_TEI(Row, colKONPO_TANTO))  '����@�S����
            
                                                                                            '����@����
            If Trim(Y_Syuka_TEI(Row, colKONPO_TANTO)) <> "" Then
                If Trim(Y_Syuka_TEI(Row, colKONPO_DATETIME)) = "" Then
                    Call UniCode_Conv(Y_SYU_TEI_REC.KONPO_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                Else
                    Call UniCode_Conv(Y_SYU_TEI_REC.KONPO_DATETIME, Y_Syuka_TEI(Row, colKONPO_DATETIME))
                End If
            Else
                Call UniCode_Conv(Y_SYU_TEI_REC.KONPO_DATETIME, "")
            End If
            
            
            
            Call UniCode_Conv(Y_SYU_TEI_REC.JITU_SUU, "")                                   '���o�ɐ�(�����ւ̏o�ɐ� ���ݖ��g�p)
            Call UniCode_Conv(Y_SYU_TEI_REC.JITU_TANTO, "")                                 '�o�Ɂ@�S����(���ݖ��g�p)
            Call UniCode_Conv(Y_SYU_TEI_REC.JITU_DATETIME, "")                              '�o�Ɂ@����(���ݖ��g�p)
            Call UniCode_Conv(Y_SYU_TEI_REC.SYUGO_KONPO_TANTO, Y_Syuka_TEI(Row, colSYUGO_KONPO_TANTO))  '�W������@�S����
            
                                                                                            '�W������@����
            If Trim(Y_Syuka_TEI(Row, colSYUGO_KONPO_TANTO)) <> "" Then
                If Trim(Y_Syuka_TEI(Row, colSYUGO_KONPO_DATETIME)) = "" Then
                    Call UniCode_Conv(Y_SYU_TEI_REC.SYUGO_KONPO_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                Else
                    Call UniCode_Conv(Y_SYU_TEI_REC.SYUGO_KONPO_DATETIME, Y_Syuka_TEI(Row, colSYUGO_KONPO_DATETIME))
                End If
            Else
                Call UniCode_Conv(Y_SYU_TEI_REC.SYUGO_KONPO_DATETIME, "")
            End If
            
            
                                                                                            '����
            If IsNumeric(Y_Syuka_TEI(Row, colKUTI_SU)) Then
                Call UniCode_Conv(Y_SYU_TEI_REC.KUTI_SU, Format(Val(Y_Syuka_TEI(Row, colKUTI_SU)), "0000"))
            End If
            
                                                                                            '�ː�
            If IsNumeric(Y_Syuka_TEI(Row, colSAI_SU)) Then
                Call UniCode_Conv(Y_SYU_TEI_REC.SAI_SU, Format(Val(Y_Syuka_TEI(Row, colSAI_SU)), "000.00"))
            End If
            
            
            
            
            Call UniCode_Conv(Y_SYU_TEI_REC.SHOGO_TANTO, Y_Syuka_TEI(Row, colSHOGO_TANTO))      '�ƍ��@�S����
                                                                                                '�ƍ�  ����
            If Trim(Y_Syuka_TEI(Row, colSHOGO_TANTO)) <> "" Then
                If Trim(Y_Syuka_TEI(Row, colSHOGO_DATETIME)) = "" Then
                    Call UniCode_Conv(Y_SYU_TEI_REC.SHOGO_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                Else
                    Call UniCode_Conv(Y_SYU_TEI_REC.SHOGO_DATETIME, Y_Syuka_TEI(Row, colSHOGO_DATETIME))
                End If
            Else
                Call UniCode_Conv(Y_SYU_TEI_REC.SHOGO_DATETIME, "")
            End If
            

            Call UniCode_Conv(Y_SYU_TEI_REC.KENPIN_TANTO, Y_Syuka_TEI(Row, colKENPIN_TANTO))    '���i�@�S����
                                                                                                '���i  ����
            If Trim(Y_Syuka_TEI(Row, colKENPIN_TANTO)) <> "" Then
                If Trim(Y_Syuka_TEI(Row, colKENPIN_DATETIME)) = "" Then
                    Call UniCode_Conv(Y_SYU_TEI_REC.KENPIN_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                Else
                    Call UniCode_Conv(Y_SYU_TEI_REC.KENPIN_DATETIME, Y_Syuka_TEI(Row, colKENPIN_DATETIME))
                End If
            Else
                Call UniCode_Conv(Y_SYU_TEI_REC.KENPIN_DATETIME, "")
            End If



            
            
            
            Call UniCode_Conv(Y_SYU_TEI_REC.FILLER, "")                                     'FILLER
            
            
            If Upd_com = BtOpInsert Then
                                                                                            '�ǉ��S����
                Call UniCode_Conv(Y_SYU_TEI_REC.INS_TANTO, StrConv(App.EXEName, vbUpperCase))
                                                                                            '�ǉ�����
                Call UniCode_Conv(Y_SYU_TEI_REC.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))
            End If
            
            Call UniCode_Conv(Y_SYU_TEI_REC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))   '�X�V�S����
            Call UniCode_Conv(Y_SYU_TEI_REC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))    '�X�V����
                    
                    
        End If
            
            
        Do
            sts = BTRV(Upd_com, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    
                    Beep
                    ans = MsgBox("�u�@�ʒ����f�[�^�v���[���Ńf�[�^�g�p���ł��B<Y_SYUKA_TEI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Call Input_UnLock
                        
                        GoTo Abort_Tran
                    End If
                
                Case Else
                    Call Input_UnLock
                    Call File_Error(sts, Upd_com, "�@�ʒ����f�[�^")
                    GoTo Abort_Tran
            End Select
        
        Loop
            
        TDBGrid1.Bookmark = Row
        

    Next Row

    Set TDBGrid1.Array = Y_Syuka_TEI
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst



                                        '�g�����U�N�V�����I��
    sts = BTRV(BtOpEndTransaction, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts <> BtNoErr Then
        Call Input_UnLock
        Call File_Error(sts, BtOpEndTransaction, "", 0)
        TEI_Update_Proc = SYS_ERR
        GoTo Abort_Tran
    End If
    
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�ϐ��n�E�X�@�ʒ����ް��ƍ��@[�����ް�]�X�V�����I���I�I", Me.hwnd, 0)
    
    Call Input_UnLock

    TEI_Update_Proc = False
    
    Exit Function


Abort_Tran:
    
    Call Input_UnLock
    sts = BTRV(BtOpAbortTransaction, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�ϐ��n�E�X�@�ʒ����ް��ƍ��@[�����ް�]�X�V�����ُ�I���I�I", Me.hwnd, 0)



End Function
Private Function YOTEI_Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   �u�o�ח\��f�[�^�v�o�^����
'----------------------------------------------------------------------------

Dim sts             As Integer
Dim ans             As Integer
    
Dim com             As Integer
Dim Upd_com         As Integer

Dim Skip_Flg        As Integer

Dim Row             As Long



    If Y_Syuka_TEI.Count(1) = 0 Then
        Exit Function
    End If
    
    YOTEI_Update_Proc = True
    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�ϐ��n�E�X�@�ʒ����ް��ƍ��@[�o�ח\��]�X�V�����J�n�I�I", Me.hwnd, 0)

                                    
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts <> BtNoErr Then
        Call Input_Lock
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        YOTEI_Update_Proc = SYS_ERR
        Exit Function
    End If
                                    
                                    
                                    
                                    
    Set TDBGrid1.Array = Y_Syuka_TEI
    TDBGrid1.Refresh
    
    TDBGrid1.Update
                                    
                                    
                                    
                                    
                                    
                                    '�e�[�u�����Z�b�g
    For Row = 1 To Y_Syuka_TEI.UpperBound(1)
        
        
        DoEvents
        
        
        Call UniCode_Conv(K4_Y_SYU_H.ID_NO, Y_Syuka_TEI(Row, colID_NO))
        
        
        Skip_Flg = False
        
        sts = BTRV(BtOpGetEqual, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
        Select Case sts
            Case BtNoErr
            
            
            
            
                Call UniCode_Conv(Y_SYU_HREC.SEK_KEN_NO, Y_Syuka_TEI(Row, colKEN_NO))           '���Ǉ��@�@�@���Ǘ���(��)
                Call UniCode_Conv(Y_SYU_HREC.SEK_HIN_NO, Y_Syuka_TEI(Row, colHIN_NO))           '���Ǉ��@�@�@���Ǘ���(��)
                    
                    
                    
                    
                Call UniCode_Conv(Y_SYU_HREC.SEK_SHOGO_TANTO, Y_Syuka_TEI(Row, colSHOGO_TANTO)) '�ƍ��@�S����
                                                                                                '�ƍ�  ����
                If Trim(Y_Syuka_TEI(Row, colSHOGO_TANTO)) <> "" Then
                    If Trim(Y_Syuka_TEI(Row, colSHOGO_DATETIME)) = "" Then
                        Call UniCode_Conv(Y_SYU_HREC.SEK_SHOGO_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                    Else
                        Call UniCode_Conv(Y_SYU_HREC.SEK_SHOGO_DATETIME, Y_Syuka_TEI(Row, colKONPO_DATETIME))
                    End If
                Else
                    Call UniCode_Conv(Y_SYU_HREC.SEK_SHOGO_DATETIME, "")
                End If
                    
                    
                    
                If IsNumeric(Left(Y_Syuka_TEI(Row, colOKURI_NO_SEQ), 3)) Then
                    Call UniCode_Conv(Y_SYU_HREC.OKURI_NO_SEQ, Left(Y_Syuka_TEI(Row, colOKURI_NO_SEQ), 3))
                    Call UniCode_Conv(Y_SYU_HREC.OKURI_NO_SEQ_TO, Right(Y_Syuka_TEI(Row, colOKURI_NO_SEQ), 3))
                Else
                    Call UniCode_Conv(Y_SYU_HREC.OKURI_NO_SEQ, "")
                    Call UniCode_Conv(Y_SYU_HREC.OKURI_NO_SEQ_TO, "")
                End If
                    
                    
                    
                    
                    
                Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))      '�X�V�S����
                Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))       '�X�V����
                            
                    
                Do
                    sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            
                            Beep
                            ans = MsgBox("�u�o�ח\��f�[�^�v���[���Ńf�[�^�g�p���ł��B<Y_SYUKA_TEI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Call Input_UnLock
                                
                                GoTo Abort_Tran
                            End If
                        
                        Case Else
                            Call Input_UnLock
                            Call File_Error(sts, BtOpUpdate, "�o�ח\��(H)")
                            GoTo Abort_Tran
                    End Select
                
                Loop
            
            Case BtErrKeyNotFound
            
            
            Case Else
                Call Input_UnLock
                
                Call File_Error(sts, BtOpGetEqual, "�o�ח\��(H)")
                GoTo Abort_Tran
        End Select
        
        
        
        
        
        Set TDBGrid1.Array = Y_Syuka_TEI
        TDBGrid1.ReBind
        
        TDBGrid1.Update
        TDBGrid1.Bookmark = Row
        
        
        

    Next Row

    Set TDBGrid1.Array = Y_Syuka_TEI
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst



                                        '�g�����U�N�V�����I��
    sts = BTRV(BtOpEndTransaction, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts <> BtNoErr Then
        Call Input_UnLock
        Call File_Error(sts, BtOpEndTransaction, "", 0)
        YOTEI_Update_Proc = SYS_ERR
        GoTo Abort_Tran
    End If
    
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�ϐ��n�E�X�@�ʒ����ް��ƍ��@[�o�ח\��]�X�V�����I���I�I", Me.hwnd, 0)
    
    Call Input_UnLock

    YOTEI_Update_Proc = False
    
    Exit Function


Abort_Tran:
    
    Call Input_UnLock
    sts = BTRV(BtOpAbortTransaction, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�ϐ��n�E�X�@�ʒ����ް��ƍ��@[�o�ח\��]�X�V�����ُ�I���I�I", Me.hwnd, 0)



End Function





Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   �u�����f�[�^�v�Ǎ��ݏ���
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer
    


Dim Row             As Long

Dim i               As Integer

Dim SKIP_F          As Boolean


    List_Disp_Proc = True
    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�ϐ��n�E�X�@�ʒ����ް��ƍ��@�\�������J�n�I�I", Me.hwnd, 0)
    DoEvents

                                    '�e�[�u�����Z�b�g
    Set Y_Syuka_TEI = Nothing
    Row = Min_Row - 1
    
    
    
    
    
    
    com = BtOpGetFirst

    
        
    Do

        DoEvents


        sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K3_Y_SYU_H, Len(K3_Y_SYU_H), 3)
        Select Case sts
            Case BtNoErr
                            
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "�o�ח\���ް�(H)")
                Exit Function
        End Select




        Call UniCode_Conv(K0_Y_SYU.JGYOBU, SETSUBI)
        Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, StrConv(Y_SYU_HREC.ID_NO, vbUnicode))

        SKIP_F = False
        
        sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
                For i = 0 To UBound(SEK_MUKE_CODE)
                    If Trim(SEK_MUKE_CODE(i)) = Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) Then
                        Exit For
                    End If
                Next i
            
                If i > UBound(SEK_MUKE_CODE) Then
                    SKIP_F = True
                End If
            
            Case BtErrKeyNotFound
                SKIP_F = True
            Case Else
                Call Input_UnLock
                Call File_Error(sts, BtOpGetEqual, "�o�ח\��f�[�^")
                Exit Function
        End Select


        If Not SKIP_F Then
        
            Call UniCode_Conv(K2_Y_SYU_TEI.KEN_NO, StrConv(Y_SYU_HREC.SEK_KEN_NO, vbUnicode))
            Call UniCode_Conv(K2_Y_SYU_TEI.HIN_NO, StrConv(Y_SYU_HREC.SEK_HIN_NO, vbUnicode))
    
    
            sts = BTRV(BtOpGetEqual, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K2_Y_SYU_TEI, Len(K2_Y_SYU_TEI), 2)
            Select Case sts
                Case BtNoErr
                
                    Row = Row + 1
                    
                    
                    Y_Syuka_TEI.ReDim Min_Row, Row, Min_Col, Max_Col
                
                
                    If Trim(StrConv(Y_SYU_TEI_REC.SHOGO_TANTO, vbUnicode)) <> "" Then
                        Y_Syuka_TEI(Row, colSHORI) = "�ƍ��ς�"
                    End If
                        
                
                
                                            
                    If Trim(StrConv(Y_SYU_HREC.CANCEL_F, vbUnicode)) = "1" Then
                        Y_Syuka_TEI(Row, colCANCEL_F) = "�~"
                    Else
                        Y_Syuka_TEI(Row, colCANCEL_F) = "�@"
                    End If
                
                
                    Y_Syuka_TEI(Row, colDEN_NO) = Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode))
                    Y_Syuka_TEI(Row, colID_NO) = Trim(StrConv(Y_SYU_HREC.ID_NO, vbUnicode))
                
                
                    Y_Syuka_TEI(Row, colSND_YMD) = Trim(StrConv(Y_SYU_TEI_REC.SND_YMD, vbUnicode))          '�f�[�^�쐬��
                    Y_Syuka_TEI(Row, colSND_HMS) = Trim(StrConv(Y_SYU_TEI_REC.SND_HMS, vbUnicode))          '�f�[�^�쐬����
                    Y_Syuka_TEI(Row, colSEQ_NO) = Trim(StrConv(Y_SYU_TEI_REC.SEQ_NO, vbUnicode))            '�A��
                    Y_Syuka_TEI(Row, colJUC_YMD) = Trim(StrConv(Y_SYU_TEI_REC.JUC_YMD, vbUnicode))          '�󒍓�
                    Y_Syuka_TEI(Row, colNOU_CD) = Trim(StrConv(Y_SYU_TEI_REC.NOU_CD, vbUnicode))            '�[�������
                    Y_Syuka_TEI(Row, colNOU_NM) = Trim(StrConv(Y_SYU_TEI_REC.NOU_NM, vbUnicode))            '�[������ꖼ
                    Y_Syuka_TEI(Row, colTOK_CD) = Trim(StrConv(Y_SYU_TEI_REC.TOK_CD, vbUnicode))            '���Ӑ溰��
                    Y_Syuka_TEI(Row, colCHO_CD) = Trim(StrConv(Y_SYU_TEI_REC.CHO_CD, vbUnicode))            '���[�溰��
                    Y_Syuka_TEI(Row, colTHINB_CD) = Trim(StrConv(Y_SYU_TEI_REC.THINB_CD, vbUnicode))        '���Ӑ�i�ԁ@���i��(��)
                    Y_Syuka_TEI(Row, colHINB_CD) = Trim(StrConv(Y_SYU_TEI_REC.HINB_CD, vbUnicode))          '�i�ԁ@���i��(��)
                    Y_Syuka_TEI(Row, colCHU_CD) = Trim(StrConv(Y_SYU_TEI_REC.CHU_CD, vbUnicode))            '�������@    ���w�}��(��)
                    Y_Syuka_TEI(Row, colSYU_JUN) = Trim(StrConv(Y_SYU_TEI_REC.SYU_JUN, vbUnicode))          '�o�׏��ԁ@  ���w�}��(���E��)
                    Y_Syuka_TEI(Row, colTEI_NM) = Trim(StrConv(Y_SYU_TEI_REC.TEI_NM, vbUnicode))            '�@���@      ���w�}��(���E�E)
                                                                                                            '�󒍐���
                    Y_Syuka_TEI(Row, colJUC_SUU) = Format(Val(StrConv(Y_SYU_TEI_REC.JUC_SUU, vbUnicode)), "#0")
                    Y_Syuka_TEI(Row, colSYU_YMD) = Trim(StrConv(Y_SYU_TEI_REC.SYU_YMD, vbUnicode))          '�o�׊m���
                    Y_Syuka_TEI(Row, colNOU_YMD) = Trim(StrConv(Y_SYU_TEI_REC.NOU_YMD, vbUnicode))          '�[����
                    Y_Syuka_TEI(Row, colKEN_NO) = Trim(StrConv(Y_SYU_TEI_REC.KEN_NO, vbUnicode))            '���Ǉ��@�@�@���Ǘ���(��)
                    Y_Syuka_TEI(Row, colHIN_NO) = Trim(StrConv(Y_SYU_TEI_REC.HIN_NO, vbUnicode))            '���Ǉ��@�@�@���Ǘ���(��)
                    Y_Syuka_TEI(Row, colTANP_KB) = Trim(StrConv(Y_SYU_TEI_REC.TANP_KB, vbUnicode))          '�P�i�敪
            
            
                                                                                                            '�@������ID
                    Y_Syuka_TEI(Row, colTEI_LABELID) = Trim(StrConv(Y_SYU_TEI_REC.TEI_LABELID, vbUnicode))
                    Y_Syuka_TEI(Row, colHAKO_NO) = Trim(StrConv(Y_SYU_TEI_REC.HAKO_NO, vbUnicode))          '����
                    
                    
                    
                    Y_Syuka_TEI(Row, colJITU_SUU) = Trim(StrConv(Y_SYU_TEI_REC.JITU_SUU, vbUnicode))        '���o�ɐ�
                                                                                                    
                    Y_Syuka_TEI(Row, colJITU_TANTO) = Trim(StrConv(Y_SYU_TEI_REC.JITU_TANTO, vbUnicode))    '�o�Ɂ@�S����
                                                                                                            '�o�Ɂ@����
                    Y_Syuka_TEI(Row, colJITU_DATETIME) = Trim(StrConv(Y_SYU_TEI_REC.JITU_DATETIME, vbUnicode))
                    
                    
                    
                    Y_Syuka_TEI(Row, colKONPO_TANTO) = Trim(StrConv(Y_SYU_TEI_REC.KONPO_TANTO, vbUnicode))  '����@�S����
                                                                                                            '����  ����
                    Y_Syuka_TEI(Row, colKONPO_DATETIME) = Trim(StrConv(Y_SYU_TEI_REC.KONPO_DATETIME, vbUnicode))
                                                                                                            
                                                                                                            
                                                                                                            
                    Y_Syuka_TEI(Row, colKONPO_ID) = Trim(StrConv(Y_SYU_TEI_REC.KONPO_ID, vbUnicode))        '����ID
                    
                                                                                                            '�W������@�S����
                    Y_Syuka_TEI(Row, colSYUGO_KONPO_TANTO) = Trim(StrConv(Y_SYU_TEI_REC.SYUGO_KONPO_TANTO, vbUnicode))
                                                                                                            '�W������ ����
                    Y_Syuka_TEI(Row, colSYUGO_KONPO_DATETIME) = Trim(StrConv(Y_SYU_TEI_REC.SYUGO_KONPO_DATETIME, vbUnicode))
                    
                    
                    
                                                                                                            
                                                                                                    
                    Y_Syuka_TEI(Row, colSHOGO_TANTO) = Trim(StrConv(Y_SYU_TEI_REC.SHOGO_TANTO, vbUnicode))  '�ƍ��@�S����
                                                                                                            '�ƍ�  ����
                    Y_Syuka_TEI(Row, colSHOGO_DATETIME) = Trim(StrConv(Y_SYU_TEI_REC.SHOGO_DATETIME, vbUnicode))
                                                                                                            '����
                    Y_Syuka_TEI(Row, colKUTI_SU) = Format(Val(StrConv(Y_SYU_TEI_REC.KUTI_SU, vbUnicode)), "#0")
                                                                                                            '�ː�
                    Y_Syuka_TEI(Row, colSAI_SU) = Format(Val(StrConv(Y_SYU_TEI_REC.SAI_SU, vbUnicode)), "#0.00")
                                                                                                    
                                                                                                            '�}��
                    Y_Syuka_TEI(Row, colOKURI_NO_SEQ) = Trim(StrConv(Y_SYU_HREC.OKURI_NO_SEQ, vbUnicode)) & "�`" & Trim(StrConv(Y_SYU_HREC.OKURI_NO_SEQ_TO, vbUnicode))
                                                                                                    
                                                                                                            
                                                                                                    
                                                                                                            '���i�@�S����
                    Y_Syuka_TEI(Row, colKENPIN_TANTO) = Trim(StrConv(Y_SYU_TEI_REC.KENPIN_TANTO, vbUnicode))
                                                                                                            '���i  ����
                    Y_Syuka_TEI(Row, colKENPIN_DATETIME) = Trim(StrConv(Y_SYU_TEI_REC.KENPIN_DATETIME, vbUnicode))
                                                                                                    
                    Y_Syuka_TEI(Row, colINS_TANTO) = Trim(StrConv(Y_SYU_TEI_REC.INS_TANTO, vbUnicode))      '�ǉ��@�S����
                                                                                                            '�ǉ�  ����
                    Y_Syuka_TEI(Row, colINS_DateTime) = Trim(StrConv(Y_SYU_TEI_REC.Ins_DateTime, vbUnicode))
                                                                                                    
                    Y_Syuka_TEI(Row, colUPD_TANTO) = Trim(StrConv(Y_SYU_TEI_REC.UPD_TANTO, vbUnicode))      '�ǉ��@�S����
                                                                                                            '�ǉ�  ����
                    Y_Syuka_TEI(Row, colUPD_DATETIME) = Trim(StrConv(Y_SYU_TEI_REC.UPD_DATETIME, vbUnicode))
                
                Case BtErrKeyNotFound
                
                    If Disp_Mode Then
                        
                        Row = Row + 1
                        Y_Syuka_TEI.ReDim Min_Row, Row, Min_Col, Max_Col
                        
                        Y_Syuka_TEI(Row, colDEN_NO) = Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode))
                        Y_Syuka_TEI(Row, colID_NO) = Trim(StrConv(Y_SYU_HREC.ID_NO, vbUnicode))
                        
                        Y_Syuka_TEI(Row, colKEN_NO) = Trim(StrConv(Y_SYU_HREC.SEK_KEN_NO, vbUnicode))           '���Ǉ��@�@�@���Ǘ���(��)
                        Y_Syuka_TEI(Row, colHIN_NO) = Trim(StrConv(Y_SYU_HREC.SEK_HIN_NO, vbUnicode))           '���Ǉ��@�@�@���Ǘ���(��)
                
                    End If
                Case Else
                    Call Input_UnLock
                    Call File_Error(sts, BtOpGetEqual, "�@�ʒ����f�[�^")
                    Exit Function
            End Select
    
            
    
    
            Text1.Text = Format(Row, "#")

        End If
        
        com = BtOpGetNext

    Loop


    Set TDBGrid1.Array = Y_Syuka_TEI
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst



hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�ϐ��n�E�X�@�ʒ����ް��ƍ��@�\�������I���I�I", Me.hwnd, 0)



    Call Input_UnLock


    List_Disp_Proc = False
    Exit Function

Error_Proc:
    
    Call Input_UnLock
    MsgBox "Err.Number= " & Err.Number & " " & Err.Description

End Function

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    
    
Dim lngPFstRow  As Long
Dim vntBmk      As Variant
Dim intLeftCol  As Integer
Dim intCol      As Integer
Dim lngCFstRow  As Long
    
    
    
    
    If Y_Syuka_TEI.Count(1) < 1 Then
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
                    
        Y_Syuka_TEI.QuickSort Min_Row, Y_Syuka_TEI.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
        
        Set TDBGrid1.Array = Y_Syuka_TEI
        
        With TDBGrid1
              .SetFocus
              lngPFstRow = TDBGrid1.FirstRow
              vntBmk = .Bookmark
              intLeftCol = .LeftCol
              intCol = .Col
              .ReBind
              .Col = intCol
              .LeftCol = intLeftCol
              .Bookmark = vntBmk
              lngCFstRow = TDBGrid1.FirstRow
              TDBGrid1.Scroll 0, lngPFstRow - lngCFstRow
        End With
        
        TDBGrid1.Update
        TDBGrid1.MoveFirst


    End If

End Sub

Private Function Matching_Proc() As Integer
'----------------------------------------------------------------------------
'                   �u�����f�[�^�v�ƍ�����
'----------------------------------------------------------------------------
Dim com             As Integer
Dim sts             As Integer

Dim Row             As Long
Dim SHORI_MSG       As String

Dim Y_SYU_NON       As Integer
Dim Y_SYU_H_NON     As Integer
Dim Y_SYU_TEI_NON   As Integer

Dim SEK_SHOGO_DATETIME _
                    As String * 14


Dim SUMI_QTY        As Long
Dim MI_QTY          As Long

Dim i               As Long
Dim SYUKA_QTY       As Long

Dim SHOGO_TANTO     As String


Dim Ins_DateTime    As String


Dim Mode            As Integer

    If Y_Syuka_TEI.Count(1) < 1 Then
        Exit Function
    End If
    
    
    Matching_Proc = True
    Call Input_Lock
    
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�ϐ��n�E�X�@�ʒ����ް��ƍ��@�ƍ������J�n�I�I", Me.hwnd, 0)

                                    
    
    Ins_DateTime = Format(Now, "YYYYMMDDHHMMSS")
    
    For Row = 1 To Y_Syuka_TEI.Count(1)
    
        DoEvents
    
        
        Y_SYU_NON = False
        Y_SYU_H_NON = False
        Y_SYU_TEI_NON = False
        SHORI_MSG = ""
        SHOGO_TANTO = ""
    
    
        
        
        '------------------------------------------------   �o�ח\��Ǎ���
        Call UniCode_Conv(K0_Y_SYU.JGYOBU, SETSUBI)
        Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Y_Syuka_TEI(Row, colID_NO))
        sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Y_SYU_NON = True
            Case Else
                Call Input_UnLock
                Call File_Error(sts, BtOpGetEqual, "�o�ח\��")
                Exit Function
        End Select
        '------------------------------------------------   �o�ח\��(H)�Ǎ���
        Call UniCode_Conv(K4_Y_SYU_H.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Y_SYU_H_NON = True
            Case Else
                Call Input_UnLock
                Call File_Error(sts, BtOpGetEqual, "�o�ח\��(H)")
                Exit Function
        End Select
        '------------------------------------------------   �@�ʒ����ް��Ǎ���
        If Trim(StrConv(Y_SYU_HREC.SEK_KEN_NO, vbUnicode)) <> "" And Trim(StrConv(Y_SYU_HREC.SEK_HIN_NO, vbUnicode)) <> "" Then
            Call UniCode_Conv(K2_Y_SYU_TEI.KEN_NO, StrConv(Y_SYU_HREC.SEK_KEN_NO, vbUnicode))
            Call UniCode_Conv(K2_Y_SYU_TEI.HIN_NO, StrConv(Y_SYU_HREC.SEK_HIN_NO, vbUnicode))
            
            sts = BTRV(BtOpGetEqual, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K2_Y_SYU_TEI, Len(K2_Y_SYU_TEI), 2)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Y_SYU_TEI_NON = True
                Case Else
                    Call Input_UnLock
                    Call File_Error(sts, BtOpGetEqual, "�@�ʒ����ް�")
                    Exit Function
            End Select
        End If
        '------------------------------------------------   �G���[�`�F�b�N  ----------------------
        
        
        '-------------------------------------- �o�ח\��o�^�ُ�I�I
        If Y_SYU_NON Or Y_SYU_H_NON Then
            SHORI_MSG = "ERR �o�ח\�薢�o�^�I�I�i�o�ח\��f�[�^�m�F�v�j"
            GoTo NEXT_LOOP
        End If
        '-------------------------------------- �����f�[�^�o�^�ُ�I�I
        If Y_SYU_TEI_NON Then
            If Disp_Mode Then
                If Trim(StrConv(Y_SYU_HREC.SEK_KEN_NO, vbUnicode)) <> "" Then
                    SHORI_MSG = "ERR �����f�[�^���o�^�I�I�i�����f�[�^�m�F�v�j"
                    GoTo NEXT_LOOP
                End If
            End If
        End If
        '-------------------------------------- �o�ח\��L�����Z���ς�
        If StrConv(Y_SYU_HREC.CANCEL_F, vbUnicode) = "1" Then
''''            SHORI_MSG = "ERR �o�ח\��L�����Z���ς�"
            If Y_SYUKA_TEI_Update(Ins_DateTime, Row, 1) Then
                Exit Function
            End If
            SHORI_MSG = "*�ƍ��@�n�j"
            
            GoTo NEXT_LOOP
        End If
        '-------------------------------------- �����f�[�^�Ȃ�
        If Y_SYU_TEI_NON Then
            SHORI_MSG = ""
            GoTo NEXT_LOOP
        End If
        '-------------------------------------- �o�ח\�茟�i�ُ�
        If IsDate(Left(StrConv(Y_SYU_HREC.KENPIN_NOW, vbUnicode), 8)) Then
            If Trim(StrConv(Y_SYU_TEI_REC.KENPIN_DATETIME, vbUnicode)) = "" Then
            
'                SHORI_MSG = "ERR �o�ח\�茟�i�ُ�"
                
                
                If Y_SYUKA_TEI_Update(Ins_DateTime, Row, 1) Then
                    Exit Function
                End If
                SHORI_MSG = "*�ƍ��@�n�j"
                
                
                GoTo NEXT_LOOP
            End If
        End If
        '-------------------------------------- �o�ח\��ƍ��ُ�
        If IsDate(Left(StrConv(Y_SYU_HREC.SEK_SHOGO_DATETIME, vbUnicode), 8)) Then
            If Trim(StrConv(Y_SYU_TEI_REC.SHOGO_DATETIME, vbUnicode)) = "" Then
 '               SHORI_MSG = "ERR �o�ח\��ƍ��ُ�"
                
                If Y_SYUKA_TEI_Update(Ins_DateTime, Row, 1) Then
                    Exit Function
                End If
                SHORI_MSG = "*�ƍ��@�n�j"
                
                
                GoTo NEXT_LOOP
            End If
        End If
        '-------------------------------------- �����ް��ƍ��ُ�
        If Trim(StrConv(Y_SYU_HREC.SEK_SHOGO_DATETIME, vbUnicode)) = "" Then
            If Trim(StrConv(Y_SYU_TEI_REC.SHOGO_DATETIME, vbUnicode)) <> "" Then
                SHORI_MSG = "ERR �����ް��ƍ��ُ�"
                GoTo NEXT_LOOP
            End If
        End If
        '-------------------------------------- �����ް��@����<>�o�ח\��@����
'        If CLng(StrConv(Y_SYU_TEI_REC.JUC_SUU, vbUnicode)) <> CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) Then
'            SHORI_MSG = "ERR �o�א��ʱ�ϯ�"
'            GoTo NEXT_LOOP
'        End If
        
        '-------------------------------------- �����ް��@�i��<>�o�ח\��@�i��
'        If Trim(StrConv(Y_SYU_TEI_REC.HINB_CD, vbUnicode)) <> Trim(StrConv(Y_SYUREC.HIN_NO, vbUnicode)) Then
'            SHORI_MSG = "ERR �i�ڱ�ϯ�"
'            GoTo NEXT_LOOP
'        End If
        
        
        '-------------------------------------- �o�ח\�芮���ς�
        If Trim(StrConv(Y_SYU_TEI_REC.SHOGO_DATETIME, vbUnicode)) = "" Then
            If StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = KAN_KBN_FIN Then
'                SHORI_MSG = "ERR �o�ח\��蓮�����ς�"
                
                
                If Y_SYUKA_TEI_Update(Ins_DateTime, Row, 1) Then
                    Exit Function
                End If
                SHORI_MSG = "*�ƍ��@�n�j"
                
                GoTo NEXT_LOOP
            End If
        End If
        '-------------------------------------- �����ް��ƍ��ς�
        If Trim(StrConv(Y_SYU_HREC.SEK_SHOGO_DATETIME, vbUnicode)) <> "" Then           '2018.04.06
            If Trim(StrConv(Y_SYU_HREC.SEK_SHOGO_DATETIME, vbUnicode)) = Trim(StrConv(Y_SYU_TEI_REC.SHOGO_DATETIME, vbUnicode)) Then
                SHORI_MSG = "�ƍ��ς�"
                GoTo NEXT_LOOP
            End If
        End If
        '-------------------------------------- �����ް�������
        If KONPO_F Then
            If Trim(StrConv(Y_SYU_TEI_REC.KONPO_DATETIME, vbUnicode)) = "" Then
                SHORI_MSG = "ERR �����ް�������"
                GoTo NEXT_LOOP
            End If
        End If
        '-------------------------------------- �����ް��W��������
        If SYUGO_KONPO_F Then
            If Trim(StrConv(Y_SYU_TEI_REC.KONPO_ID, vbUnicode)) <> "" Then
                If Trim(StrConv(Y_SYU_TEI_REC.SYUGO_KONPO_DATETIME, vbUnicode)) = "" Then
                    SHORI_MSG = "ERR �����ް��W��������"
                    GoTo NEXT_LOOP
                End If
            End If
        End If
        '-------------------------------------- ���i��Ԃɂ�锻��
        Select Case KENPIN_F
            Case 0
            
                If Trim(StrConv(Y_SYU_TEI_REC.KENPIN_DATETIME, vbUnicode)) <> "" Then
                    SHORI_MSG = "ERR �����ް����i�ς�"
                    GoTo NEXT_LOOP
                End If
            
            
            Case 1
            
                If Trim(StrConv(Y_SYU_TEI_REC.KENPIN_DATETIME, vbUnicode)) = "" Then
                    SHORI_MSG = "ERR �����ް������i"
                    GoTo NEXT_LOOP
                End If
            
            Case 2
        
        End Select
        
        
        '------------------ �݌ɂ̎g�p�\����s���A�L���݌ɐ����l������
        sts = Zaiko_Reserve_Proc(0, Zaiko_Tanaban, SETSUBI, NAIGAI_NAI, StrConv(Y_SYU_TEI_REC.HINB_CD, vbUnicode), SUMI_QTY, MI_QTY)
        Select Case sts
            Case False
            Case True           '�����ł͔������Ȃ�
                Exit Function
            Case SYS_ERR
                Exit Function
            Case SYS_CANCEL
                SHORI_MSG = "ERR �݌Ɏg�p��"
                GoTo NEXT_LOOP
        End Select
        
            
        SYUKA_QTY = 0
        For i = 1 To Y_Syuka_TEI.Count(1)
            If Trim(StrConv(Y_SYU_TEI_REC.HINB_CD, vbUnicode)) = Trim(Y_Syuka_TEI(i, colHINB_CD)) Then
                
                If Trim(Y_Syuka_TEI(i, colSHOGO_TANTO)) = "" Then
                        
                    SYUKA_QTY = SYUKA_QTY + Val(Y_Syuka_TEI(i, colJUC_SUU))
                End If
            End If
        Next i
            
        If ZAIKO_F Then
            
            If SYUKA_QTY > (SUMI_QTY + MI_QTY) Then
                SHORI_MSG = "ERR �݌ɕs��"
                GoTo NEXT_LOOP
            End If
        End If
        
        
        
        If SYUKA_QTY = 0 Then
        Else
            '----------------------------------- �f�[�^�X�V�����J�n -----------
                                                '�g�����U�N�V�����J�n
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Input_Lock
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
                Matching_Proc = SYS_ERR
                Exit Function
            End If
            '�o�ɏ���
            Mode = 0
            sts = Syuko_SEK_Update_Proc(SETSUBI, NAIGAI_NAI, StrConv(Y_SYU_TEI_REC.HINB_CD, vbUnicode), _
                    "", _
                    Zaiko_Tanaban, _
                    YOIN_CODE, _
                    0, 0, CLng(StrConv(Y_SYU_TEI_REC.JUC_SUU, vbUnicode)), _
                    WS_NO, _
                    "SEK30", _
                    10, _
                    "�ƍ��������������Ƃ�", _
                    StrConv(Y_SYUREC.CYU_KBN, vbUnicode), _
                    StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & StrConv(Y_SYUREC.SS_CODE, vbUnicode), _
                    StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), _
                    StrConv(Y_SYU_HREC.DEN_NO, vbUnicode), _
                    StrConv(Y_SYU_HREC.ID_NO, vbUnicode), _
                    MENU_NO, _
                    StrConv(Y_SYU_HREC.INS_BIN, vbUnicode), , Ins_DateTime, Mode)
            Select Case sts
                Case False
                Case True
                    GoTo Abort_Tran
                Case SYS_CANCEL
                    GoTo Abort_Tran
                Case SYS_ERR
                    GoTo Abort_Tran
            
            End Select
        
                            '�g�����U�N�V�����I��
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call File_Error(sts, BtOpEndTransaction, "")
                GoTo Abort_Tran
            End If
            '�݌ɂ̗\�����
            If Data_Clear_Proc(0) Then
                Exit Function
            End If
            
            
            SHOGO_TANTO = StrConv(App.EXEName, vbUpperCase)
            
            If Mode = 0 Then
                SHORI_MSG = "�ƍ��@�n�j"
            Else
                SHORI_MSG = "*�ƍ��@�n�j"
            End If
                    
        
        
        End If
    
    
    
    
    
    
    
    
    
    
NEXT_LOOP:
    
        Y_Syuka_TEI(Row, colSHORI) = SHORI_MSG
        Y_Syuka_TEI(Row, colSHOGO_TANTO) = SHOGO_TANTO
        
    
        Set TDBGrid1.Array = Y_Syuka_TEI
        TDBGrid1.ReBind
        
        TDBGrid1.Update
        TDBGrid1.Bookmark = Row
    
    
    
    
    
    
    Next Row
    
    
    Set TDBGrid1.Array = Y_Syuka_TEI
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst
        
    
    
    
    
    Call Input_UnLock
    
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�ϐ��n�E�X�@�ʒ����ް��ƍ��@�ƍ������I���I�I", Me.hwnd, 0)
    
    
    Matching_Proc = False
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If

    Call Input_UnLock



End Function
Private Function Data_Clear_Proc(Mode As Integer) As Integer
'-------------------------------------------------------
'
'   �w�o�ח\��^�݌ɂ̗\��L�����Z���x
'   ���ח\��̸ر��ǉ�      2007.06.07
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim ans         As Integer
    
    
    Data_Clear_Proc = True
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
    


    If Mode = 0 Then
                                        '�݌ɂ̊J��
        Call UniCode_Conv(K3_ZAIKO.WEL_ID, WS_NO)
        Call UniCode_Conv(K3_ZAIKO.PRG_ID, StrConv(App.EXEName, vbUpperCase))
        com = BtOpGetGreaterEqual
    Else
        Call UniCode_Conv(K3_ZAIKO.WEL_ID, "")
        Call UniCode_Conv(K3_ZAIKO.PRG_ID, "")
        com = BtOpGetGreater
    End If
    
    Do
        DoEvents
        
        Do
            sts = BTRV(com + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K3_ZAIKO, Len(K3_ZAIKO), 3)
            Select Case sts
                Case BtNoErr
                    
                    If Mode = 0 Then
                        If WS_NO <> StrConv(ZAIKOREC.WEL_ID, vbUnicode) Or _
                             StrConv(App.EXEName, vbUpperCase) <> Trim(StrConv(ZAIKOREC.PRG_ID, vbUnicode)) Then
                            sts = BtErrEOF
                        
                        
                            sts = BTRV(BtOpUnlock, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K3_ZAIKO, Len(K3_ZAIKO), 3)
                            If sts Then
                                Call File_Error(sts, BtOpUnlock, "�݌Ƀf�[�^")
                                GoTo Abort_Tran
                            End If
                        
                        
                        End If
                    End If
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        GoTo Abort_Tran
                    End If
                
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^")
                    GoTo Abort_Tran
            End Select
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
    
        
        
        Call UniCode_Conv(ZAIKOREC.WEL_ID, "")
        Call UniCode_Conv(ZAIKOREC.PRG_ID, "")
        Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)
        Do
        
            sts = BTRV(BtOpUpdate, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K3_ZAIKO, Len(K3_ZAIKO), 3)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        GoTo Abort_Tran
                    End If
                
                Case Else
                    Call File_Error(sts, BtOpUpdate, "�݌Ƀf�[�^")
                    GoTo Abort_Tran
            End Select
        
        Loop
    
        com = BtOpGetNext
    
    Loop





End_Tran:
                                        '�g�����U�N�V�����I��
    sts = BTRV(BtOpEndTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    Data_Clear_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If


End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------
Dim i   As Integer


    SEK00351.MousePointer = vbHourglass

    Call Ctrl_Lock(SEK00351)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(SEK00351)


    SEK00351.MousePointer = vbDefault

End Sub


Private Function Zaiko_Reserve_Proc(ID_NO As String, _
                                    FROM_LOCATION As String, _
                                    JGYOBU As String, _
                                    NAIGAI As String, _
                                    Hinban As String, _
                                    SUMI_QTY As Long, _
                                    MI_QTY As Long) As Integer
'-------------------------------------------------------
'
'   �w�݌Ƀf�[�^�̎g�p�\��x
'
'-------------------------------------------------------
Dim sts             As Integer

    Zaiko_Reserve_Proc = True
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, ZAIKO_POS, ZAIKOREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        Zaiko_Reserve_Proc = SYS_ERR
        Exit Function
    End If

    sts = Zaiko_Lock_Proc(FROM_LOCATION, JGYOBU, NAIGAI, Hinban, Format(ID_NO, "000"), SUMI_QTY, MI_QTY, 10)
    If sts Then
        Zaiko_Reserve_Proc = sts
        GoTo Abort_Tran
    End If
End_Tran:
                                        '�g�����U�N�V�����I��
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "", 0)
        Zaiko_Reserve_Proc = SYS_ERR
        GoTo Abort_Tran
    End If
    
    Zaiko_Reserve_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function



Private Function Data_Link_Proc() As Integer
'----------------------------------------------------------------------------
'                   �u�����f�[�^�v�Ɓu�o�ח\��̃����N�v����
'----------------------------------------------------------------------------
Dim Row     As Integer

Dim sts     As Integer
Dim com     As Integer
Dim ans     As Integer
    
    If Y_Syuka_TEI.Count(1) = 0 Then
        Exit Function
    End If
    
    Data_Link_Proc = True
    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�ϐ��n�E�X�@�ʒ����ް��ƍ��@[�o�ח\�聕�����f�[�^�����N�����@�e�X�g�p]�X�V�����J�n�I�I", Me.hwnd, 0)

                                    
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts <> BtNoErr Then
        Call Input_Lock
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Data_Link_Proc = SYS_ERR
        Exit Function
    End If
                                    
                                    
                                    
                                    
    Set TDBGrid1.Array = Y_Syuka_TEI
    TDBGrid1.Refresh
    
    TDBGrid1.Update
                                    
                                    
    com = BtOpGetFirst
                                    
                                    
                                    '�e�[�u�����Z�b�g
    For Row = 1 To Y_Syuka_TEI.UpperBound(1)
        
        
        DoEvents
        
        
        Call UniCode_Conv(K4_Y_SYU_H.ID_NO, Y_Syuka_TEI(Row, colID_NO))
        
        sts = BTRV(BtOpGetEqual, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
        Select Case sts
            Case BtNoErr
            
            
                sts = BTRV(com, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
                Select Case sts
                    Case BtNoErr
                        com = BtOpGetNext
                    Case BtErrEOF
                        Exit For
                    Case Else
            
                        Call Input_UnLock
                        
                        Call File_Error(sts, BtOpGetEqual, "�o�ח\��(H)")
                        GoTo Abort_Tran
            
                End Select
            
            
            
            
            
                Call UniCode_Conv(Y_SYU_HREC.SEK_KEN_NO, StrConv(Y_SYU_TEI_REC.KEN_NO, vbUnicode))      '���Ǉ��@�@�@���Ǘ���(��)
                Call UniCode_Conv(Y_SYU_HREC.SEK_HIN_NO, StrConv(Y_SYU_TEI_REC.HIN_NO, vbUnicode))      '���Ǉ��@�@�@���Ǘ���(��)
                    
                    
                                    
                    
                    
                    
                    
                Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))      '�X�V�S����
                Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))       '�X�V����
                            
                    
                Do
                    sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            
                            Beep
                            ans = MsgBox("�u�o�ח\��f�[�^�v���[���Ńf�[�^�g�p���ł��B<Y_SYUKA(H).DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Call Input_UnLock
                                
                                GoTo Abort_Tran
                            End If
                        
                        Case Else
                            Call Input_UnLock
                            Call File_Error(sts, BtOpUpdate, "�o�ח\��(H)")
                            GoTo Abort_Tran
                    End Select
                
                Loop
            
            Case BtErrKeyNotFound
            
            
            Case Else
                Call Input_UnLock
                
                Call File_Error(sts, BtOpGetEqual, "�o�ח\��(H)")
                GoTo Abort_Tran
        End Select
        
        
        
        
        Call UniCode_Conv(K0_Y_SYU.JGYOBU, SETSUBI)
        Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Y_Syuka_TEI(Row, colID_NO))
        
        sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
            
                Call UniCode_Conv(Y_SYUREC.MUKE_CODE, StrConv(Y_SYU_TEI_REC.TOK_CD, vbUnicode))         '���Ӑ溰��
                Call UniCode_Conv(Y_SYUREC.SS_CODE, StrConv(Y_SYU_TEI_REC.CHO_CD, vbUnicode))           '���[�溰��
                    
                    
                Call UniCode_Conv(Y_SYUREC.UPD_NOW, Format(Now, "YYYYMMDDHHMMSS"))                      '�X�V����
                            
                    
                Do
                    sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            
                            Beep
                            ans = MsgBox("�u�o�ח\��f�[�^�v���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Call Input_UnLock
                                
                                GoTo Abort_Tran
                            End If
                        
                        Case Else
                            Call Input_UnLock
                            Call File_Error(sts, BtOpUpdate, "�o�ח\��")
                            GoTo Abort_Tran
                    End Select
                
                Loop
            
            Case BtErrKeyNotFound
            
            
            Case Else
                Call Input_UnLock
                
                Call File_Error(sts, BtOpGetEqual, "�o�ח\��")
                GoTo Abort_Tran
        End Select
        
        
        
        
        Set TDBGrid1.Array = Y_Syuka_TEI
        TDBGrid1.ReBind
        
        TDBGrid1.Update
        TDBGrid1.Bookmark = Row
        
        
        

    Next Row

    Set TDBGrid1.Array = Y_Syuka_TEI
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst



                                        '�g�����U�N�V�����I��
    sts = BTRV(BtOpEndTransaction, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts <> BtNoErr Then
        Call Input_UnLock
        Call File_Error(sts, BtOpEndTransaction, "", 0)
        Data_Link_Proc = SYS_ERR
        GoTo Abort_Tran
    End If
    
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�ϐ��n�E�X�@�ʒ����ް��ƍ��@[�o�ח\��]�X�V�����I���I�I", Me.hwnd, 0)
    
    Call Input_UnLock

    Data_Link_Proc = False
    
    Exit Function


Abort_Tran:
    
    Call Input_UnLock
    sts = BTRV(BtOpAbortTransaction, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�ϐ��n�E�X�@�ʒ����ް��ƍ��@[�o�ח\��]�X�V�����ُ�I���I�I", Me.hwnd, 0)

End Function


Private Function CSV_DATA_OUT_PROC() As Integer
'----------------------------------------------------------------------------
'                   �u�����f�[�^�vCSV�o�͏���
'----------------------------------------------------------------------------
Dim FileNo          As Long
Dim i               As Long


    If Y_Syuka_TEI.Count(1) < 1 Then
        Exit Function
    End If



    CSV_DATA_OUT_PROC = True
    
    
    
    
    FileNo = FreeFile
    
    On Error GoTo Error_Proc
    Open ("Y_Syuka_TEI.CSV") For Output As FileNo
    On Error GoTo 0

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�ϐ��n�E�X�@�ʏo�ׁ^�����ް��m�F�@�b�r�u�o�͏����J�n�I�I", Me.hwnd, 0)



    For i = 1 To Y_Syuka_TEI.Count(1)
    
        DoEvents
    
        Text1.Text = Format(i, "#0")
        
        DoEvents
    
            
            

            
    
        Write #FileNo, "*" & Trim(Left(Y_Syuka_TEI(i, colID_NO), 7)) & "*",
    
        
        Write #FileNo, "*" & Trim(Y_Syuka_TEI(i, colHINB_CD)) & "*",
        
        Write #FileNo, Format(Val(Y_Syuka_TEI(i, colJUC_SUU)), "#0"),

        
        If Trim(Y_Syuka_TEI(i, colTEI_LABELID)) = "" Then
            Write #FileNo, ,
        Else
            Write #FileNo, "*" & Trim(Y_Syuka_TEI(i, colTEI_LABELID)) & "*",
        End If
                
        If Trim(Y_Syuka_TEI(i, colKONPO_ID)) = "" Then
            Write #FileNo, ,
        Else
            Write #FileNo, "*" & Trim(Y_Syuka_TEI(i, colKONPO_ID)) & "*",
        End If
            
            
        Write #FileNo, Val(Y_Syuka_TEI(i, colKUTI_SU)),
        Write #FileNo, Val(Y_Syuka_TEI(i, colSAI_SU)),
'
'        Write #FileNo, Y_Syuka_TEI(i, colOKURI_NO_SEQ),
            
            
'        Write #FileNo, Y_Syuka_TEI(i, colSHOGO_TANTO),
'        Write #FileNo, "/" & Y_Syuka_TEI(i, colSHOGO_DATETIME),
            
            
            
            
            
        Write #FileNo,
    Next i



    MsgBox ("�uY_Syuka_TEI.CSV�v�o�͏I���I�I")
    


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�ϐ��n�E�X�@�ʒ����ް��m�F�@�b�r�u�o�͏����I���I�I", Me.hwnd, 0)

    Close #FileNo


    CSV_DATA_OUT_PROC = False
    Exit Function
    
Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox "Y_Syuka_TEI.CSV" & "���g�p���ł��B"
    Else
        MsgBox "Err.Number= " & Err.Number & " " & Err.Description
            
    End If





End Function



Private Function Y_SYUKA_TEI_Update(Ins_DateTime As String, Row As Long, Mode As Integer) As Integer
            
'----------------------------------------------------------------------------
'                   �u�����f�[�^�v�X�V����
'----------------------------------------------------------------------------
            
Dim sts As Integer
Dim ans As Integer
            
        Y_SYUKA_TEI_Update = True
            
                                                '�g�����U�N�V�����J�n
        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        If sts <> BtNoErr Then
            Call Input_Lock
            Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
            Exit Function
        End If
            
            
        Call UniCode_Conv(K2_Y_SYU_TEI.KEN_NO, Y_Syuka_TEI(Row, colKEN_NO))
        Call UniCode_Conv(K2_Y_SYU_TEI.HIN_NO, Y_Syuka_TEI(Row, colHIN_NO))
        
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K2_Y_SYU_TEI, Len(K2_Y_SYU_TEI), 2)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYU_TEI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        GoTo Abort_Tran
                    End If
                Case BtErrKeyNotFound
                    Y_SYUKA_TEI_Update = False
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�@�ʒ����ް�")
                    GoTo Abort_Tran
            End Select
        Loop
            
            
        Call UniCode_Conv(Y_SYU_TEI_REC.CANCEL_F, StrConv(Y_SYU_HREC.CANCEL_F, vbUnicode))
            
        Call UniCode_Conv(Y_SYU_TEI_REC.SHOGO_TANTO, StrConv(App.EXEName, vbUpperCase))
        Call UniCode_Conv(Y_SYU_TEI_REC.SHOGO_DATETIME, Ins_DateTime)

    
        Call UniCode_Conv(Y_SYU_TEI_REC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
        Call UniCode_Conv(Y_SYU_TEI_REC.UPD_DATETIME, Ins_DateTime)
    
    
    
    
    
        Do
            sts = BTRV(BtOpUpdate, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K2_Y_SYU_TEI, Len(K2_Y_SYU_TEI), 2)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYU_TEI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        GoTo Abort_Tran
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "�@�ʒ����ް�")
                    GoTo Abort_Tran
            End Select
        Loop

'------------------------------------------------------

        If Mode = 1 Then
        Else

            Call UniCode_Conv(K4_Y_SYU_H.ID_NO, Y_Syuka_TEI(Row, colID_NO))
            
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYU_H.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            GoTo Abort_Tran
                        End If
                    Case BtErrKeyNotFound
                        Y_SYUKA_TEI_Update = False
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�o�ח\��(H)")
                        GoTo Abort_Tran
                End Select
            Loop
                
                
                
                
                
            Call UniCode_Conv(Y_SYU_HREC.SEK_SHOGO_TANTO, StrConv(App.EXEName, vbUpperCase))
            Call UniCode_Conv(Y_SYU_HREC.SEK_SHOGO_DATETIME, Ins_DateTime)
        
            Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
            Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, Ins_DateTime)
        
        
        
        
        
            Do
                sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYU_H.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(BtOpUpdate, BtOpUpdate, "�o�ח\��(H)")
                        GoTo Abort_Tran
    
                End Select
            Loop
        End If

        sts = BTRV(BtOpEndTransaction, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
        If sts <> BtNoErr Then
            Call File_Error(sts, BtOpEndTransaction, "")
        End If
    





        Y_SYUKA_TEI_Update = False
        Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If



End Function
