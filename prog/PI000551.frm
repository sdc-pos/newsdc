VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PI000551 
   Caption         =   "���ޔ��㏈��"
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16965
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
   ScaleHeight     =   10545
   ScaleWidth      =   16965
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox txtLOAD_LIMIT 
      Height          =   375
      Left            =   11640
      TabIndex        =   47
      Top             =   9960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   15
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   45
      Top             =   4800
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   0
      Top             =   120
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   14
      Left            =   5265
      MaxLength       =   12
      TabIndex        =   17
      Top             =   4080
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   13
      Left            =   8085
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   12
      Left            =   5265
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   11
      Left            =   5265
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   10
      Left            =   1680
      MaxLength       =   12
      TabIndex        =   13
      Top             =   4080
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   9
      Left            =   1680
      MaxLength       =   11
      TabIndex        =   12
      Top             =   3720
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      Index           =   8
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   11
      Top             =   3360
      Width           =   1485
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   2
      Left            =   2400
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   10
      Top             =   2760
      Width           =   4050
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   9
      Top             =   2760
      Width           =   750
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   1
      Left            =   2400
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   8
      Top             =   2400
      Width           =   4050
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   7
      Top             =   2400
      Width           =   750
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   5
      Left            =   4320
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1800
      Width           =   5025
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   4305
      MaxLength       =   7
      TabIndex        =   2
      Top             =   720
      Width           =   1065
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   2400
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   4
      Top             =   1320
      Width           =   4005
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   1
      Top             =   720
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�I ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10650
      TabIndex        =   29
      Top             =   9960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9810
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   9960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8970
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   9960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   8130
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   9960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6810
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   9960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5970
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   9960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   5130
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   9960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4290
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   9960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ݾ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2970
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   9960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2130
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   9960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1290
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   9960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�X �V"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   450
      TabIndex        =   18
      Top             =   9960
      Width           =   855
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   4575
      Left            =   315
      TabIndex        =   44
      Top             =   5280
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   8070
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "���㇂"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "����N����"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "�����N��"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "���Ӑ�"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "���ޕi��"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "���x�P��"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "�̔��敪"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "����"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "�P��"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "���z"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "�����"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1879"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1773"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=512"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2434"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2328"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=512"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1852"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1746"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=512"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2699"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2593"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=512"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=6112"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=6006"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=512"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=1879"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1773"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=512"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=1958"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1852"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=512"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=2064"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=1958"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=514"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=2328"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2223"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=514"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(9).Width=2064"
      Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=1958"
      Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=514"
      Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(51)=   "Column(10).Width=2064"
      Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=1958"
      Splits(0)._ColumnProps(54)=   "Column(10)._ColStyle=514"
      Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �S�V�b�N"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �S�V�b�N"
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
      _StyleDefs(5)   =   ":id=0,.fontname=�l�r �S�V�b�N"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFF&,.bold=0,.fontsize=1200"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=�l�r �S�V�b�N"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=0,.fontsize=1200"
      _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=�l�r �S�V�b�N"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=�l�r �S�V�b�N"
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
      _StyleDefs(26)  =   ":id=43,.fontname=�l�r �S�V�b�N"
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
      _StyleDefs(40)  =   ":id=58,.fontname=�l�r �S�V�b�N"
      _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=44"
      _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=45"
      _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=47"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=62,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(45)  =   ":id=62,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(46)  =   ":id=62,.fontname=�l�r �S�V�b�N"
      _StyleDefs(47)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=44"
      _StyleDefs(48)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=45"
      _StyleDefs(49)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=47"
      _StyleDefs(50)  =   "Splits(0).Columns(2).Style:id=16,.parent=43,.alignment=0"
      _StyleDefs(51)  =   "Splits(0).Columns(2).HeadingStyle:id=13,.parent=44"
      _StyleDefs(52)  =   "Splits(0).Columns(2).FooterStyle:id=14,.parent=45"
      _StyleDefs(53)  =   "Splits(0).Columns(2).EditorStyle:id=15,.parent=47"
      _StyleDefs(54)  =   "Splits(0).Columns(3).Style:id=28,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(55)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(56)  =   ":id=28,.fontname=�l�r �S�V�b�N"
      _StyleDefs(57)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=44"
      _StyleDefs(58)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=45"
      _StyleDefs(59)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=47"
      _StyleDefs(60)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(61)  =   ":id=66,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(62)  =   ":id=66,.fontname=�l�r �S�V�b�N"
      _StyleDefs(63)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44"
      _StyleDefs(64)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
      _StyleDefs(65)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
      _StyleDefs(66)  =   "Splits(0).Columns(5).Style:id=32,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(67)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(68)  =   ":id=32,.fontname=�l�r �S�V�b�N"
      _StyleDefs(69)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=44"
      _StyleDefs(70)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=45"
      _StyleDefs(71)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=47"
      _StyleDefs(72)  =   "Splits(0).Columns(6).Style:id=24,.parent=43,.alignment=0"
      _StyleDefs(73)  =   "Splits(0).Columns(6).HeadingStyle:id=21,.parent=44"
      _StyleDefs(74)  =   "Splits(0).Columns(6).FooterStyle:id=22,.parent=45"
      _StyleDefs(75)  =   "Splits(0).Columns(6).EditorStyle:id=23,.parent=47"
      _StyleDefs(76)  =   "Splits(0).Columns(7).Style:id=20,.parent=43,.alignment=1"
      _StyleDefs(77)  =   "Splits(0).Columns(7).HeadingStyle:id=17,.parent=44"
      _StyleDefs(78)  =   "Splits(0).Columns(7).FooterStyle:id=18,.parent=45"
      _StyleDefs(79)  =   "Splits(0).Columns(7).EditorStyle:id=19,.parent=47"
      _StyleDefs(80)  =   "Splits(0).Columns(8).Style:id=70,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(81)  =   ":id=70,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(82)  =   ":id=70,.fontname=�l�r �S�V�b�N"
      _StyleDefs(83)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=44"
      _StyleDefs(84)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=45"
      _StyleDefs(85)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=47"
      _StyleDefs(86)  =   "Splits(0).Columns(9).Style:id=74,.parent=43,.alignment=1"
      _StyleDefs(87)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=44"
      _StyleDefs(88)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=45"
      _StyleDefs(89)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=47"
      _StyleDefs(90)  =   "Splits(0).Columns(10).Style:id=78,.parent=43,.alignment=1"
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
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1320
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�����N������"
      Height          =   255
      Index           =   14
      Left            =   75
      TabIndex        =   46
      Top             =   4920
      Width           =   1545
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���㇂"
      Height          =   255
      Index           =   13
      Left            =   840
      TabIndex        =   43
      Top             =   240
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�����"
      Height          =   255
      Index           =   11
      Left            =   4305
      TabIndex        =   42
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�W������"
      Height          =   255
      Index           =   10
      Left            =   6720
      TabIndex        =   41
      Top             =   3840
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�W������"
      Height          =   255
      Index           =   9
      Left            =   3990
      TabIndex        =   40
      Top             =   3840
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�݌ɐ���"
      Height          =   255
      Index           =   8
      Left            =   3990
      TabIndex        =   39
      Top             =   3480
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���z"
      Height          =   255
      Index           =   7
      Left            =   945
      TabIndex        =   38
      Top             =   4200
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�P��"
      Height          =   255
      Index           =   5
      Left            =   945
      TabIndex        =   37
      Top             =   3840
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "����"
      Height          =   255
      Index           =   4
      Left            =   945
      TabIndex        =   36
      Top             =   3480
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�̔��敪"
      Height          =   255
      Index           =   3
      Left            =   420
      TabIndex        =   35
      Top             =   2880
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���x�P��"
      Height          =   255
      Index           =   2
      Left            =   420
      TabIndex        =   34
      Top             =   2520
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���ޕi��"
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   33
      Top             =   1920
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�����N��"
      Height          =   255
      Index           =   12
      Left            =   3150
      TabIndex        =   32
      Top             =   840
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���Ӑ�"
      Height          =   255
      Index           =   6
      Left            =   735
      TabIndex        =   31
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "����N����"
      Height          =   255
      Index           =   1
      Left            =   315
      TabIndex        =   30
      Top             =   840
      Width           =   1275
   End
End
Attribute VB_Name = "PI000551"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private POS_UMU     As Boolean              'POS���т̗L��
    
Private YOIN        As String * 2           'POS���і��̏o�ɗv��
Private TANTO       As String * 5           'POS���і��̒S���Һ���

    
Dim WS_NO           As String * 3
    
Dim URIAGE          As New XArrayDB
    
    
Private Const Min_Row% = 1              '�ŏ��s��
'Private Const Max_Row& = 2000           '�ő�s��
Dim Max_Row     As Long                 '���X�g�{�b�N�X�ő�\������

Private Const Min_Col% = 0              '�ŏ���
Private Const Max_Col% = 10             '�ő��

Private Const ColURIAGE_NO% = 0         '�� ���㇂
Private Const ColURIAGE_DT% = 1         '�� ������t
Private Const ColKEIJYO_YM% = 2         '�� �����N��
Private Const ColTOKUI_CODE% = 3        '�� ���Ӑ�
Private Const ColHIN_GAI% = 4           '�� ���ޕi��
Private Const ColG_SYUSHI% = 5          '�� ���x�P��
Private Const ColG_HANBAI_KBN% = 6      '�� �̔��敪
Private Const ColURIAGE_QTY% = 7        '�� ���㐔��
Private Const ColTANKA% = 8             '�� �P��
Private Const ColKINGAKU% = 9           '�� ���z
Private Const ColZEI_KIN% = 10          '�� �����

Private Sort_Tbl(Min_Col To Max_Col) _
                As Integer                  '��Ă̐��� 0:���� 1:�~��
    
    
'�e�L�X�g�p�Y��

Private Const ptxURIAGE_NO% = 0             '���㇂


Private Const ptxURIAGE_DT% = 1             '����N����
Private Const ptxKEIJYO_YM% = 2             '�v�㌎

Private Const ptxTOKUI_CODE% = 3            '���Ӑ溰��

Private Const ptxHIN_GAI% = 4               '�i��
Private Const ptxHIN_NAME% = 5              '�i��

Private Const ptxG_SYUSHI% = 6              '���x�P��
Private Const ptxG_HANBAI_KBN% = 7          '�̔��敪

Private Const ptxURIAGE_QTY% = 8            '���㐔��
Private Const ptxTANKA% = 9                 '�P��
Private Const ptxKINGAKU% = 10              '���z

Private Const ptxZAIKO_QTY% = 11            '�݌Ɏc
Private Const ptxG_ST_URITAN% = 12          '�W���e������
Private Const ptxG_ST_SHITAN% = 13          '�W���e������

Private Const ptxZEI_KIN% = 14              '�����


Private Const ptxSEL_KEIJYO_YM% = 15        '�����v�㌎

'�R���{�p�Y��
Private Const pcmbTOKUI% = 0                '���Ӑ�
Private Const pcmbG_SYUSHI% = 1             '���x�P��
Private Const pcmbG_HANBAI_KBN% = 2         '�̔��P��


'�P��   2007.07.10
Private wkTANKA     As String
'����   2007.07.10
Private wkQTY       As String


Private UKEIRE_DT       As Integer          '�㉺���ݒ� ������@2017.11.20
Private KEIJYO_YM       As Integer          '�㉺���ݒ� �v�㌎�@2017.11.20


'Private Const LAST_UPDATE_DAY$ = "[PI00055] 2017.11.27 11:00"
'Private Const LAST_UPDATE_DAY$ = "[PI00055] 2018.01.31 13:20"
'Private Const LAST_UPDATE_DAY$ = "[PI00055] 2019.10.03 18:05"
Private Const LAST_UPDATE_DAY$ = "[PI00055] 2019.10.04 09:45"

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PI000551.MousePointer = vbHourglass

    Call Ctrl_Lock(PI000551)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PI000551)


    PI000551.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
    
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer

Dim SUMI_QTY    As Long
Dim MI_QTY      As Long
    
Dim wkDate      As String * 10
Dim ckDATE      As String               '2018.01.31
    
Dim ST_Sumi_Qty As Long
Dim ST_Mi_Qty   As Long
    
Dim ZEI         As Long
Dim wkKINGAKU   As Long
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        
        
        Case ptxURIAGE_NO       '���㇂
        
            If Trim(Text1(ptxURIAGE_NO).Text) = "" Then
            Else
                If IsNumeric(Text1(ptxURIAGE_NO).Text) Then
                    Text1(ptxURIAGE_NO).Text = Format(CLng(Text1(ptxURIAGE_NO).Text), "00000")
                End If
        
        
        
                        
                
                If Item_Disp_Proc() Then
                    Exit Function
                End If
        
            End If
        
        
        Case ptxURIAGE_DT       '����N����
            
            If Not IsDate(Text1(ptxURIAGE_DT).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B(����N����)"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxURIAGE_DT).Text = Format(CDate(Text1(ptxURIAGE_DT).Text), "YYYY/MM/DD")
            
            
'>>>>>>>>>>>>>>>>>>>>>  �㉺���͈����� 2017.11.17
                If DateAdd("m", UKEIRE_DT * -1, Format(Now, "YYYY/MM/DD")) <= Text1(ptxURIAGE_DT).Text And _
                    DateAdd("m", UKEIRE_DT, Format(Now, "YYYY/MM/DD")) >= Text1(ptxURIAGE_DT).Text Then
                Else
                    MsgBox "������t�����t�͈͂𒴂��Ă��܂��B"
                    Text1(Mode).SetFocus
                    Exit Function
                End If



'>>>>>>>>>>>>>>>>>>>>>  �㉺���͈����� 2017.11.17
            
            
            
            End If
        
        Case ptxKEIJYO_YM       '�����N��
            
            
            wkDate = Text1(ptxKEIJYO_YM).Text & "/01"
            
            If Not IsDate(wkDate) Then
                MsgBox "���͂������ڂ̓G���[�ł��B�i�����N�����j"
                Text1(Mode).SetFocus
                Exit Function
            Else
                
                wkDate = Format(CDate(Text1(ptxKEIJYO_YM).Text), "YYYY/MM/DD")
                
                Text1(ptxKEIJYO_YM).Text = Mid(wkDate, 1, 7)
            
'>>>>>>>>>>>>>>>>>>>>>  �㉺���͈����� 2017.11.17
                If Format(DateAdd("m", KEIJYO_YM * -1, Format(Now, "YYYY/MM/DD")), "YYYY/MM/DD") > Text1(ptxKEIJYO_YM).Text & Right(Format(Now, "YYYY/MM/DD"), 3) Then
                    
                    MsgBox "�����N�������t�͈͂𒴂��Ă��܂��B"
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



'                If Format(DateAdd("m", KEIJYO_YM, Format(Now, "YYYY/MM/DD")), "YYYY/MM/DD") < (Text1(ptxKEIJYO_YM).Text & Right(Format(Now, "YYYY/MM/DD"), 3)) Then        '2018.01.31
                If Format(DateAdd("m", KEIJYO_YM, Format(Now, "YYYY/MM/DD")), "YYYY/MM/DD") < ckDATE Then                                                                   '2018.01.31
                    MsgBox "�����N�������t�͈͂𒴂��Ă��܂��B"
                    Text1(Mode).SetFocus
                    Exit Function
                End If



'>>>>>>>>>>>>>>>>>>>>>  �㉺���͈����� 2017.11.17
            
            
            End If
        
        Case ptxTOKUI_CODE   '���Ӑ�
            
           Text1(ptxTOKUI_CODE).Text = StrConv(Text1(ptxTOKUI_CODE).Text, vbUpperCase)      '2017.11.20
            
           Combo1(pcmbTOKUI).ListIndex = -1
           For i = 0 To Combo1(pcmbTOKUI).ListCount - 1
               If Trim(Text1(ptxTOKUI_CODE).Text) = Trim(Right(Combo1(pcmbTOKUI).List(i), 5)) Then
                   Combo1(pcmbTOKUI).ListIndex = i
                   Exit For
               End If
           
           Next i
    
           If i > Combo1(pcmbTOKUI).ListCount - 1 Then
               MsgBox "���͂������ڂ̓G���[�ł��B�i���Ӑ斢�o�^�j"
               Text1(Mode).SetFocus
               Exit Function
           End If
        
        Case ptxHIN_GAI         '�i��
    
                    
    
            Text1(ptxHIN_GAI).Text = StrConv(Text1(ptxHIN_GAI).Text, vbUpperCase)   '2017.11.20
    
            If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI And _
                StrConv(ITEMREC.NAIGAI, vbUnicode) = NAIGAI_NAI And _
                Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) = Trim(Text1(ptxHIN_GAI).Text) Then
    
            Else
                Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        
                        Text1(ptxHIN_NAME).Text = ""
                        Text1(ptxZAIKO_QTY).Text = ""
                        Text1(ptxG_ST_URITAN).Text = ""
                        Text1(ptxG_ST_SHITAN).Text = ""
                        
                        MsgBox "���͂������ڂ̓G���[�ł��B�i�i�ږ��o�^�j"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                        Exit Function
                
                End Select
                
                            
                
                Text1(ptxHIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                
                '���x�P��
                Text1(ptxG_SYUSHI).Text = StrConv(ITEMREC.G_SYUSHI, vbUnicode)
                Combo1(pcmbG_SYUSHI).ListIndex = -1
                For i = 0 To Combo1(pcmbG_SYUSHI).ListCount - 1
                    If Text1(ptxG_SYUSHI).Text = Trim(Right(Combo1(pcmbG_SYUSHI).List(i), 3)) Then
                        Combo1(pcmbG_SYUSHI).ListIndex = i
                        Exit For
                    End If
                
                Next i
                '�̔��敪
                Text1(ptxG_HANBAI_KBN).Text = StrConv(ITEMREC.G_HANBAI_KBN, vbUnicode)
                Combo1(pcmbG_HANBAI_KBN).ListIndex = -1
                For i = 0 To Combo1(pcmbG_HANBAI_KBN).ListCount - 1
                    If Text1(ptxG_HANBAI_KBN).Text = Trim(Left(Right(Combo1(pcmbG_HANBAI_KBN).List(i), 3), 2)) Then
                        Combo1(pcmbG_HANBAI_KBN).ListIndex = i
                        Exit For
                    End If
                
                Next i
                
                
                If Not POS_UMU Then
                '�o�n�r�����ŕW���I�Ԗ��ݒ�͏o�ɕs��2006.04.26
                    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" And _
                        Trim(StrConv(ITEMREC.ST_RETU, vbUnicode)) = "" And _
                        Trim(StrConv(ITEMREC.ST_REN, vbUnicode)) = "" And _
                        Trim(StrConv(ITEMREC.ST_DAN, vbUnicode)) = "" Then

                        MsgBox "�W���I�Ԃ��ݒ肳��Ă��܂���B"
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
                End If
                
                
                
                
                If Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                           StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                           StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                    Exit Function
                
                End If
                            
                            
                                        

                
                Text1(ptxZAIKO_QTY).Text = Format(SUMI_QTY + MI_QTY, "#,##0")
                If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                    Text1(ptxG_ST_URITAN).Text = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#,##0.00")
                Else
                    Text1(ptxG_ST_URITAN).Text = ""
                End If
                
                If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                    Text1(ptxTANKA).Text = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
                Else
                    Text1(ptxTANKA).Text = ""
                End If
                
                If IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
                    Text1(ptxG_ST_SHITAN).Text = Format(CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)), "#,##0.00")
                Else
                    Text1(ptxG_ST_SHITAN).Text = ""
                End If
            End If
           
            
            
                    
        
        
        
        Case ptxG_SYUSHI        '���x�P��
            
            Combo1(pcmbG_SYUSHI).ListIndex = -1
            For i = 0 To Combo1(pcmbG_SYUSHI).ListCount - 1
                If Text1(ptxG_SYUSHI).Text = Trim(Right(Combo1(pcmbG_SYUSHI).List(i), 3)) Then
                    Combo1(pcmbG_SYUSHI).ListIndex = i
                    Exit For
                End If
               
            Next i
        
            If i > Combo1(pcmbG_SYUSHI).ListCount - 1 Then
                MsgBox "���͂������ڂ̓G���[�ł��B�i���x�P�ʖ��o�^�j"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        Case ptxG_HANBAI_KBN    '�̔��敪
            
            Combo1(pcmbG_HANBAI_KBN).ListIndex = -1
            For i = 0 To Combo1(pcmbG_HANBAI_KBN).ListCount - 1
                If Text1(ptxG_HANBAI_KBN).Text = Trim(Left(Right(Combo1(pcmbG_HANBAI_KBN).List(i), 3), 2)) Then
                    Combo1(pcmbG_HANBAI_KBN).ListIndex = i
                    Exit For
                End If
           
           Next i
    
           If i > Combo1(pcmbG_HANBAI_KBN).ListCount - 1 Then
               MsgBox "���͂������ڂ̓G���[�ł��B�i�̔��敪���o�^�j"
               Text1(Mode).SetFocus
               Exit Function
           End If
        
        
        
        Case ptxURIAGE_QTY       '���㐔��
    
            If Not IsNumeric(Text1(ptxURIAGE_QTY).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B�i���㐔�ʁj"
                Text1(Mode).SetFocus
                Exit Function
            Else
''                If CLng(Text1(ptxURIAGE_QTY).Text) = 0 Then
''                    MsgBox "���͂������ڂ̓G���[�ł��B"
''                    Text1(Mode).SetFocus
''                    Exit Function
''                End If
                
                Text1(ptxURIAGE_QTY).Text = Format(CLng(Text1(ptxURIAGE_QTY).Text), "#0.00")
            
                
                If Trim(Text1(ptxURIAGE_NO).Text) = "" Then
                
                
                    If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_ON Then
                        
                        If CLng(Text1(ptxURIAGE_QTY).Text) <= 0 Then
                        Else
                            If CLng(Text1(ptxURIAGE_QTY).Text) > CLng(Text1(ptxZAIKO_QTY).Text) Then
                                MsgBox "���͂������ڂ̓G���[�ł��B�i���݌ɐ��s���j"
                                Text1(Mode).SetFocus
                                Exit Function
                            End If
                        
                        
                        
                        
                            If Not POS_UMU Then
                            '�o�n�r�����ŕW���I�ԍ݌ɂōă`�F�b�N2006.04.26
                                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" And _
                                    Trim(StrConv(ITEMREC.ST_RETU, vbUnicode)) = "" And _
                                    Trim(StrConv(ITEMREC.ST_REN, vbUnicode)) = "" And _
                                    Trim(StrConv(ITEMREC.ST_DAN, vbUnicode)) = "" Then
            
                                    MsgBox "�W���I�Ԃ��ݒ肳��Ă��܂���B"
                                    Text1(Mode).SetFocus
                                    Exit Function
            
                                End If
                            
                            
                                If Zaiko_Syukei_Proc(ST_Sumi_Qty, ST_Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                           StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                           StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                                           StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                           StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                           StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                           StrConv(ITEMREC.ST_DAN, vbUnicode)) Then
                                    Exit Function
                                
                                End If
                                
                                If CLng(Text1(ptxURIAGE_QTY).Text) > ST_Sumi_Qty + ST_Mi_Qty Then
                                    MsgBox "���͂������ڂ̓G���[�ł��B�i�W���I�ԍ݌ɐ��s���j"
                                    Text1(Mode).SetFocus
                                    Exit Function
                                End If
                            End If
                            
                            
                            
                            
                        
                        
                        
                        
                        
                        End If
                            
                
                    End If
                
                End If
            
            
                            
            
            
            
            
''                If IsNumeric(Text1(ptxTANKA).Text) Then
''
''                    If Text1(ptxKINGAKU).Text = "" Then
''                        Text1(ptxKINGAKU).Text = Format(CLng(CDbl(Text1(ptxTANKA).Text) * _
''                                                    CLng(Text1(ptxURIAGE_QTY).Text)), "#,##0")
''
''                        If CLng(Text1(ptxKINGAKU).Text) >= 0 Then
''                            If Format(Text1(ptxURIAGE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
''                                ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
''                                        CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
''                            Else
''                                ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
''                                        CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
''                            End If
''                        Else
''
''                            wkKINGAKU = CLng(Text1(ptxKINGAKU).Text) * -1
''
''                            If Format(Text1(ptxURIAGE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
''                                ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
''                                        CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
''                            Else
''                                ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
''                                        CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
''                            End If
''                            ZEI = ZEI * -1
''                        End If
''
''                        Text1(ptxZEI_KIN).Text = ZEI
''
''                    End If
'-----------------------
                
                
                
                
                
                
                
                
                
                
                
''                Else
''                    Text1(ptxKINGAKU).Text = "0"
''                End If
            End If
    
    
        Case ptxTANKA           '�P��
    
            If Not IsNumeric(Text1(ptxTANKA).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B�i�P���j"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxTANKA).Text = Format(CDbl(Text1(ptxTANKA).Text), "#0.00")
            
                If Text1(ptxKINGAKU).Text = "" Then
                    If IsNumeric(Text1(ptxURIAGE_QTY).Text) Then
                        Text1(ptxKINGAKU).Text = Format(CLng(CDbl(Text1(ptxTANKA).Text) * _
                                                    CLng(Text1(ptxURIAGE_QTY).Text)), "#,##0")
                    
                    
                    
                    
                    
                    
                        If CLng(Text1(ptxKINGAKU).Text) >= 0 Then
                            If Format(Text1(ptxURIAGE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                                ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                                        CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
                            Else
'                                ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
'                                        CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
                                '2019.10.03                          ���o�O
                                ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                        CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))
                                        
                            End If
                        Else
                            
                            wkKINGAKU = CLng(Text1(ptxKINGAKU).Text) * -1
                            
                            If Format(Text1(ptxURIAGE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                                ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                                        CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
                            Else
'                                ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
'                                        CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
                                        
                                ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                        CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))
                                        
                            End If
                            ZEI = ZEI * -1
                        End If

                        Text1(ptxZEI_KIN).Text = ZEI
                    
                    End If
                End If
            End If
    
    
        Case ptxKINGAKU         '���z
    
            If Not IsNumeric(Text1(ptxTANKA).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B�i���z�j"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxTANKA).Text = Format(CDbl(Text1(ptxTANKA).Text), "#0.00")
            
                If Text1(ptxKINGAKU).Text = "" Then
                    If IsNumeric(Text1(ptxURIAGE_QTY).Text) Then
                        Text1(ptxKINGAKU).Text = Format(CLng(CDbl(Text1(ptxTANKA).Text) * _
                                                    CLng(Text1(ptxURIAGE_QTY).Text)), "#,##0")
                    
                    
                    
                    
                    
                    
                        If CLng(Text1(ptxKINGAKU).Text) >= 0 Then
                            If Format(Text1(ptxURIAGE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                                ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                                        CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
                            Else
'                                ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
'                                        CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
                                '2019.10.04                             ���o�O
                                ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                        CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))

                            End If
                        Else
                            
                            wkKINGAKU = CLng(Text1(ptxKINGAKU).Text) * -1
                            
                            If Format(Text1(ptxURIAGE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                                ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                                        CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
                            Else
'                                ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
'                                        CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
                                '2019.10.04                             ���o�O
                                ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                        CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))
                            
                            
                            End If
                            ZEI = ZEI * -1
                        End If

                        Text1(ptxZEI_KIN).Text = ZEI
                    
                    
                    End If
                End If
            End If
    
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                  ���ޔ����ް��X�V
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer




    Update_Proc = True
                                        
    Call Input_Lock
                                        
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If

                                        
    If Trim(Text1(ptxURIAGE_NO).Text) = "" Then
                                        
                                        
                                            '�Ǘ��t�@�C����莑�ޔ���ԍ��̊l��
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
                        Update_Proc = True
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�Ǘ��}�X�^")
                    GoTo Abort_Tran
            
            End Select
        
        
        Loop
        
        '�����ް����{�P
        If CLng(StrConv(P_KANRIREC.URIAGE_NO, vbUnicode)) = 99999 Then
            Call UniCode_Conv(P_KANRIREC.URIAGE_NO, "00001")
        Else
            Call UniCode_Conv(P_KANRIREC.URIAGE_NO, Format(CLng(StrConv(P_KANRIREC.URIAGE_NO, vbUnicode)) + 1, "00000"))
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
    
        Call UniCode_Conv(P_SHURIAGE_REC.URIAGE_NO, StrConv(P_KANRIREC.URIAGE_NO, vbUnicode))
    
        com = BtOpInsert
    Else
        com = BtOpUpdate
    
        Do
            Call UniCode_Conv(K0_P_SHURIAGE.URIAGE_NO, Text1(ptxURIAGE_NO).Text)
            sts = BTRV(BtOpGetEqual + BtSNoWait, P_SHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), K0_P_SHURIAGE, Len(K0_P_SHURIAGE), 0)
            Select Case sts
                Case BtNoErr
                
                    Exit Do
                Case BtErrKeyNotFound
                    
                    MsgBox "���ޔ���f�[�^���ύX����Ă��܂��B"
                    Update_Proc = False
                    GoTo Abort_Tran
                
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SHURIAGE.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        GoTo Abort_Tran
                    End If
                        
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "���ޔ���f�[�^")
                    Exit Function
            
            End Select
        Loop
    
    
    
    End If
    '---------------------------------------------------    '���ޔ���f�[�^�X�V
    Call UniCode_Conv(P_SHURIAGE_REC.URIAGE_DT, Format(Text1(ptxURIAGE_DT).Text, "YYYYMMDD"))   '�����
    Call UniCode_Conv(P_SHURIAGE_REC.KEIJYO_YM, Mid(Text1(ptxKEIJYO_YM), 1, 4) & Mid(Text1(ptxKEIJYO_YM), 6, 2))  '�v��N��
    
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxTOKUI_CODE).Text)
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
            
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            '���o�^�͈�ʈ����i�����ɂ͂��Ȃ��͂��j
            Call UniCode_Conv(P_UKEHARAIREC.TORI_KBN, P_TORI_GENERAL)
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�󕥐�}�X�^")
            Exit Function
        
    End Select

    
    
    
    Call UniCode_Conv(P_SHURIAGE_REC.TORI_KBN, StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode))      '�����敪
    Call UniCode_Conv(P_SHURIAGE_REC.TOKUI_CODE, Text1(ptxTOKUI_CODE).Text)                     '���Ӑ溰��
    Call UniCode_Conv(P_SHURIAGE_REC.JGYOBU, SHIZAI)                                            '���ƕ�
    Call UniCode_Conv(P_SHURIAGE_REC.NAIGAI, NAIGAI_NAI)                                        '�����O
    Call UniCode_Conv(P_SHURIAGE_REC.HIN_GAI, Text1(ptxHIN_GAI).Text)                           '�i��
    Call UniCode_Conv(P_SHURIAGE_REC.G_SYUSHI, Text1(ptxG_SYUSHI).Text)                         '���x�P��
    Call UniCode_Conv(P_SHURIAGE_REC.G_HANBAI_KBN, Text1(ptxG_HANBAI_KBN).Text)                 '�̔��敪
                                                                                                '����
    
    If CDbl(Text1(ptxURIAGE_QTY).Text) < 0 Then
        Call UniCode_Conv(P_SHURIAGE_REC.URIAGE_QTY, Format(CDbl(Text1(ptxURIAGE_QTY).Text), "0000000.00"))
    Else
        Call UniCode_Conv(P_SHURIAGE_REC.URIAGE_QTY, Format(CDbl(Text1(ptxURIAGE_QTY).Text), "00000000.00"))
    End If
                                                                                                '�P��
    Call UniCode_Conv(P_SHURIAGE_REC.TANKA, Format(CDbl(Text1(ptxTANKA).Text), "00000000.00"))
                                                                                                '���z
    
    If CLng(Text1(ptxKINGAKU).Text) < 0 Then
        Call UniCode_Conv(P_SHURIAGE_REC.KINGAKU, Format(CLng(Text1(ptxKINGAKU).Text), "00000000"))
    Else
        Call UniCode_Conv(P_SHURIAGE_REC.KINGAKU, Format(CLng(Text1(ptxKINGAKU).Text), "000000000"))
    End If
    
    If CLng(Text1(ptxZEI_KIN).Text) < 0 Then
        Call UniCode_Conv(P_SHURIAGE_REC.ZEI_KIN, Format(CLng(Text1(ptxZEI_KIN).Text), "00000000"))
    Else
        Call UniCode_Conv(P_SHURIAGE_REC.ZEI_KIN, Format(CLng(Text1(ptxZEI_KIN).Text), "000000000"))
    End If
    
    
    Call UniCode_Conv(P_SHURIAGE_REC.SEIKU_F, P_SEIKYU_NON)                       '�����׸�
    
    Call UniCode_Conv(P_SHURIAGE_REC.FILLER, "")
    
                                                                                    '�X�V����
    Call UniCode_Conv(P_SHURIAGE_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    Do
        
        DoEvents
        
        sts = BTRV(com, P_SHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), K0_P_SHURIAGE, Len(K0_P_SHURIAGE), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SHURIAGE.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, com, "���ޔ����ް�")
                GoTo Abort_Tran
        End Select
    
    Loop
    
    If com = BtOpInsert Then
        If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_ON Then
            If Not POS_UMU Then
                'POS���тȂ��́A�W���I�Ԃ��݌Ɉ������Ƃ�
            
                If CLng(Text1(ptxURIAGE_QTY).Text) > 0 Then
                    sts = Syuko_Update_Proc(SHIZAI, _
                                            NAIGAI_NAI, _
                                            Text1(ptxHIN_GAI).Text, _
                                            "", _
                                            (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)), _
                                            YOIN, _
                                            0, _
                                            CLng(Text1(ptxURIAGE_QTY).Text), _
                                            0, _
                                            WS_NO, _
                                            TANTO)
            
                End If
                Select Case sts
                    Case False
                    Case Else
                        Update_Proc = sts
                        GoTo Abort_Tran
                End Select
            
            
            
                        
            
            
            
            End If
        End If
    End If
End_Tran:
                                        '�g�����U�N�V�����I��
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
Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                  ���ޔ����ް��폜(��ݾ�)
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer




    Delete_Proc = True
                                        
    Call Input_Lock
                                        
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If

    
    Do
        Call UniCode_Conv(K0_P_SHURIAGE.URIAGE_NO, Text1(ptxURIAGE_NO).Text)
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_SHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), K0_P_SHURIAGE, Len(K0_P_SHURIAGE), 0)
        Select Case sts
            Case BtNoErr
            
                Exit Do
            Case BtErrKeyNotFound
                
                MsgBox "���ޔ���f�[�^���ύX����Ă��܂��B"
                Delete_Proc = False
                GoTo Abort_Tran
            
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SHURIAGE.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
                    
            Case Else
                Call File_Error(sts, BtOpGetEqual, "���ޔ���f�[�^")
                Exit Function
        
        End Select
    Loop
    
    
    
    '---------------------------------------------------    '���ޔ���f�[�^�X�V
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpDelete, P_SHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), K0_P_SHURIAGE, Len(K0_P_SHURIAGE), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SHURIAGE.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpDelete, "���ޔ����ް�")
                GoTo Abort_Tran
        End Select
    
    Loop
    

End_Tran:
                                        '�g�����U�N�V�����I��
    sts = BTRV(BtOpEndTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
    Call Input_UnLock
    
    Delete_Proc = False
    
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
        Case pcmbTOKUI          '���Ӑ�
            Text1(ptxTOKUI_CODE).Text = Trim(Right(Combo1(pcmbTOKUI).Text, 5))
        Case pcmbG_SYUSHI       '���x�P��
            Text1(ptxG_SYUSHI).Text = Trim(Right(Combo1(pcmbG_SYUSHI).Text, 3))
        Case pcmbG_HANBAI_KBN   '�̔��敪
            Text1(ptxG_HANBAI_KBN).Text = Trim(Left(Right(Combo1(pcmbG_HANBAI_KBN).Text, 3), 2))
    End Select
    
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    
    Select Case Index
        Case pcmbTOKUI          '���Ӑ�
            Text1(ptxTOKUI_CODE).Text = Trim(Right(Combo1(pcmbTOKUI).Text, 5))
        Case pcmbG_SYUSHI       '���x�P��
            Text1(ptxG_SYUSHI).Text = Trim(Right(Combo1(pcmbG_SYUSHI).Text, 3))
        Case pcmbG_HANBAI_KBN   '�̔��敪
            Text1(ptxG_HANBAI_KBN).Text = Trim(Left(Right(Combo1(pcmbG_HANBAI_KBN).Text, 3), 2))
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans         As Integer
Dim i           As Integer


Dim sts         As Integer

    Select Case Index
        Case P_CMD_Upd        '�X�V
            
            
            For i = ptxURIAGE_DT To ptxG_ST_SHITAN
            
                If Error_Check_Proc(i) Then     '�G���[�`�F�b�N
                    Exit Sub
                End If
            
            Next i
            
            Beep
            ans = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
                
                
                'LIST�\��
                If List_Disp_Proc() Then
                    Unload Me
                End If
            
                Call Init_Proc(1)
            
            
            End If
            
            
            
            Text1(ptxURIAGE_DT).SetFocus
        
        Case P_CMD_DEL                      '�폜
    
            Beep
            ans = MsgBox("�L�����Z�����܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Delete_Proc() Then
                    Unload Me
                End If
                
                
                'LIST�\��
                If List_Disp_Proc() Then
                    Unload Me
                End If
            
                Call Init_Proc(1)
            
            
            End If
    
    
        Case P_CMD_DSP                      '����/�\��
        
        
            'LIST�\��
            If List_Disp_Proc() Then
                Unload Me
            End If
        
            Call Init_Proc(1)
        
        
        Case P_CMD_OUT                      '�ް��o��
        
        Case P_CMD_PRT                      '���
            
        Case P_CMD_End                      '�I��
    
            Unload Me
    
    End Select

End Sub


Private Sub Form_DblClick()
'    PrintForm      2017.11.20
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �j���� �c������ �O����
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

Dim sBuffer As String * 255
Dim com     As String


'    If App.PrevInstance Then                       '2017.11.20
'        Beep                                       '2017.11.20
'        MsgBox "����v���O�������s���ł��B"        '2017.11.20
'        End                                        '2017.11.20
'    End If                                         '2017.11.20

                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                
    PI000551.Caption = PI000551.Caption & LAST_UPDATE_DAY   '2017.11.20
                                
                                    
                                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>   P_SYS.INI-->   PI00055.INI 2017.11.20
                                'POS���їL���̎�荞��
    If GetIni(StrConv(App.EXEName, vbUpperCase), "POS_UMU", StrConv(App.EXEName, vbUpperCase), c) Then
        POS_UMU = False
    Else
        If RTrim(c) = "0" Then
            POS_UMU = False
        Else
            POS_UMU = True
        End If
    End If
                                
    If Not POS_UMU Then
                                'POS���і����A�o�ɗv��
        If GetIni(StrConv(App.EXEName, vbUpperCase), "YOIN", StrConv(App.EXEName, vbUpperCase), c) Then
            Beep
            MsgBox "�o�ɗv���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
        End If
        YOIN = Trim(c)
    
                                'POS���і����A�S���Һ���
    
        If GetIni(StrConv(App.EXEName, vbUpperCase), "TANTO", StrConv(App.EXEName, vbUpperCase), c) Then
            TANTO = ""
        End If
        TANTO = Trim(c)
    
    
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>   P_SYS.INI-->   PI00055.INI 2017.11.20
                                
                                
'�\������   2017.11.20
    If GetIni(StrConv(App.EXEName, vbUpperCase), "LOAD_LIMIT", StrConv(App.EXEName, vbUpperCase), c) Then
        txtLOAD_LIMIT.Text = "0"
    Else
        txtLOAD_LIMIT.Text = Val(Trim(c))
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    �㉺���ݒ�  ������@2017.11.17
    If GetIni(App.EXEName, "UKEIRE_DT", App.EXEName, c) Then
        UKEIRE_DT = 0
    Else
        UKEIRE_DT = Val(RTrim(c))
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    �㉺���ݒ�  �v��N���@2017.11.17
    If GetIni(App.EXEName, "KEIJYO_YM", App.EXEName, c) Then
        KEIJYO_YM = 0
    Else
        KEIJYO_YM = Val(RTrim(c))
    End If
                                
                                
                                '���ԃ}�X�^�n�o�d�m
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�ړ����n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '������n�o�d�m
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�o�ח\��n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�v���n�o�d�m
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���۸ނn�o�d�m
    If P_SAGYO_LOG_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�i�ڃ}�X�^�n�o�d�m
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�Ǘ��}�X�^�n�o�d�m
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�󕥐�}�X�^�n�o�d�m
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ޔ����ް��n�o�d�m
    If P_SHURIAGE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '����Ͻ��n�o�d�m
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌��ް��n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    '�Ǘ��}�X�^�̓ǂݍ���
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)

    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            If P_KANRI_MAKE_Proc() Then
                Unload Me
            End If
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�Ǘ��}�X�^")
            Unload Me
    End Select
        
    
    '����Ͻ���`
    Call P_CODE_TBL_Proc

    
    '���Ӑ�
    If Ukeharai_Set_Proc(pcmbTOKUI) Then
        Unload Me
    End If
    
    '���x�P�ʂ̃Z�b�g
    If Code_Set_Proc(pcmbG_SYUSHI, P_KBN03_CD, 0) Then
        Unload Me
    End If
    
    '�̔��敪�̃Z�b�g
    If Code_Set_Proc(pcmbG_HANBAI_KBN, P_KBN02_CD, 0) Then
        Unload Me
    End If
    
                                'ܰ��ð��ݔԍ���荞��
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)
    

    'LIST�\��
    Text1(ptxSEL_KEIJYO_YM).Text = Left(Format(Now, "YYYY/MM/DD"), 7)
    If List_Disp_Proc() Then
        Unload Me
    End If
    '��ʏ����ݒ�
    Call Init_Proc


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
                                            
                                            
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
    
    
                                            '�Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�Ǘ��}�X�^")
        End If
    End If
                                            '�R�[�h�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�R�[�h�}�X�^")
        End If
    End If
                                            '�󕥐�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�󕥐�}�X�^")
        End If
    End If
                                            '���ޔ����ް��b�k�n�r�d
    sts = BTRV(BtOpClose, P_SHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), K0_P_SHURIAGE, Len(K0_P_SHURIAGE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ޔ����ް�")
        End If
    End If
                                            '�݌��ް��b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌��ް�")
        End If
    End If
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PI000551 = Nothing

    End
End Sub



Private Sub TDBGrid1_DblClick()
Dim sts As Integer
    
    '���ޒ����f�[�^�̃`�F�b�N
    Text1(ptxURIAGE_NO).Text = URIAGE(TDBGrid1.Bookmark, ColURIAGE_NO)
        
    sts = Item_Disp_Proc()
    Select Case sts
        Case False, BtNoErr
        
        Case BtErrKeyNotFound
            MsgBox "���[���ŏ����������Ă��܂��B"
            TDBGrid1.SetFocus
            Exit Sub
        Case Else
            Exit Sub
    End Select
    
    Text1(ptxURIAGE_NO).SetFocus

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
                    
        URIAGE.QuickSort Min_Row, URIAGE.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
        
        Set TDBGrid1.Array = URIAGE
        
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
    
        Case ptxURIAGE_QTY
            wkQTY = Trim(Text1(ColURIAGE_QTY).Text)
    
    End Select





End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    If Error_Check_Proc(Index) Then     '�G���[�`�F�b�N
        
        Text1(Index).SetFocus
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '�ړ�
End Sub
Private Sub Init_Proc(Optional Mode As Integer = 0)
'----------------------------------------------------------------------------
'                   ���͉�ʂ̏����ݒ�
'----------------------------------------------------------------------------
Dim i       As Integer
Dim st_i    As Integer

Dim sts     As Integer


    
    
    
    For i = ColURIAGE_NO To ptxZEI_KIN
        
        If Mode = 1 Then
            If i = ptxURIAGE_DT Or i = ptxKEIJYO_YM Or i = ptxTOKUI_CODE Then
            Else
                Text1(i).Text = ""
            End If
        Else
            Text1(i).Text = ""
        End If
    Next i
    '���ぁ����
    If Trim(Text1(ptxURIAGE_DT).Text) = "" Then
        Text1(ptxURIAGE_DT).Text = Format(Now, "YYYY/MM/DD")
    End If
    '�v�㌎
    If Mode = 0 Then
        Text1(ptxKEIJYO_YM).Text = Mid(Format(Now, "YYYY/MM/DD"), 1, 7)
    End If

    If Mode = 0 Then
        st_i = pcmbTOKUI
    Else
        st_i = pcmbG_SYUSHI
    End If
        
    For i = st_i To pcmbG_HANBAI_KBN
        
        Combo1(i).ListIndex = -1
    
    Next i




    Call UniCode_Conv(ITEMREC.JGYOBU, "")
    Call UniCode_Conv(ITEMREC.NAIGAI, "")
    Call UniCode_Conv(ITEMREC.HIN_GAI, "")

End Sub
Private Function Ukeharai_Set_Proc(Index As Integer) As Integer
'----------------------------------------------------------------------------
'                   �󕥐�}�X�^���R���{�ɃZ�b�g����B
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
                Call File_Error(sts, com, "�󕥐�}�X�^")
                Exit Function
        
        End Select

        
        
        Combo1(Index).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode) & "            " & _
                                StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        
        com = BtOpGetNext
    
    Loop

    Ukeharai_Set_Proc = False
    



End Function
Private Function Code_Set_Proc(Index As Integer, KBN As String, Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   �R�[�h�}�X�^���R���{�ɃZ�b�g����B
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
                Call File_Error(sts, com, "�R�[�h�}�X�^")
                Exit Function
        
        End Select

        wkOption = ""
        If P_KBN_TBL(i).KBN_OP1 Then
            wkOption = Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
        End If
        If P_KBN_TBL(i).KBN_OP2 Then
            wkOption = wkOption & Trim(StrConv(P_CODEREC.OPTION2, vbUnicode))
        End If
        
        
        
'        Combo1(Index).AddItem StrConv(P_CODEREC.C_RNAME, vbUnicode) & "            " &
        Combo1(Index).AddItem StrConv(P_CODEREC.C_NAME, vbUnicode) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), Key_Len) & wkOption
        
        
        com = BtOpGetNext
    
    Loop

    Code_Set_Proc = False
    



End Function



Private Function Item_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   ����f�[�^�̓ǂݍ��݁����e�\��
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim i           As Integer

Dim SUMI_QTY    As Long
Dim MI_QTY      As Long

    Item_Disp_Proc = True


    If Trim(Text1(ptxURIAGE_NO).Text) <> "" Then
        If Text1(ptxURIAGE_NO).Text = StrConv(P_SHURIAGE_REC.URIAGE_NO, vbUnicode) Then
            Item_Disp_Proc = False
            Exit Function
        End If
    End If
    Call UniCode_Conv(K0_P_SHURIAGE.URIAGE_NO, Text1(ptxURIAGE_NO).Text)
    sts = BTRV(BtOpGetEqual, P_SHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), K0_P_SHURIAGE, Len(K0_P_SHURIAGE), 0)
        
    Select Case sts
        Case BtNoErr
        
        
        Case BtErrKeyNotFound
                    
            For i = ptxURIAGE_DT To ptxZEI_KIN
            
                Text1(i).Text = ""
            
            Next i
        
            For i = pcmbTOKUI To pcmbG_HANBAI_KBN
                Combo1(i).ListIndex = -1
            Next i
            
                        
            Call UniCode_Conv(P_SHURIAGE_REC.URIAGE_NO, "")
            
            MsgBox "���ޔ���o�^����Ă��܂���B"
                    
                    
            Item_Disp_Proc = sts
            Exit Function
                    
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "���ޔ���f�[�^")
            Exit Function
    
    End Select

    Text1(ptxURIAGE_NO).Text = StrConv(P_SHURIAGE_REC.URIAGE_NO, vbUnicode)


    '����N����
    Text1(ptxURIAGE_DT).Text = Mid(StrConv(P_SHURIAGE_REC.URIAGE_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SHURIAGE_REC.URIAGE_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SHURIAGE_REC.URIAGE_DT, vbUnicode), 7, 2)

    '�����N��
    Text1(ptxKEIJYO_YM).Text = Mid(StrConv(P_SHURIAGE_REC.KEIJYO_YM, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SHURIAGE_REC.KEIJYO_YM, vbUnicode), 5, 2)

    '���Ӑ�
    Text1(ptxTOKUI_CODE).Text = StrConv(P_SHURIAGE_REC.TOKUI_CODE, vbUnicode)
    For i = 0 To Combo1(pcmbTOKUI).ListCount - 1
        If Trim(Text1(ptxTOKUI_CODE).Text) = Trim(Right(Combo1(pcmbTOKUI).List(i), 5)) Then
            Combo1(pcmbTOKUI).ListIndex = i
            Exit For
        End If
    Next i

    '���ޕi��
    
    
    Text1(ptxHIN_GAI).Text = StrConv(P_SHURIAGE_REC.HIN_GAI, vbUnicode)
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            
            Text1(ptxHIN_NAME).Text = ""
            Text1(ptxZAIKO_QTY).Text = ""
            Text1(ptxG_ST_URITAN).Text = ""
            Text1(ptxG_ST_SHITAN).Text = ""
            
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    
    End Select
    
    Text1(ptxHIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)

    '���x�P��
    Text1(ptxG_SYUSHI).Text = StrConv(P_SHURIAGE_REC.G_SYUSHI, vbUnicode)
    Combo1(pcmbG_SYUSHI).ListIndex = -1
    For i = 0 To Combo1(pcmbG_SYUSHI).ListCount - 1
        If Text1(ptxG_SYUSHI).Text = Trim(Right(Combo1(pcmbG_SYUSHI).List(i), 3)) Then
            Combo1(pcmbG_SYUSHI).ListIndex = i
            Exit For
        End If
    
    Next i
    '�̔��敪
    Text1(ptxG_HANBAI_KBN).Text = StrConv(P_SHURIAGE_REC.G_HANBAI_KBN, vbUnicode)
    Combo1(pcmbG_HANBAI_KBN).ListIndex = -1
    For i = 0 To Combo1(pcmbG_HANBAI_KBN).ListCount - 1
        If Text1(ptxG_HANBAI_KBN).Text = Trim(Left(Right(Combo1(pcmbG_HANBAI_KBN).List(i), 3), 2)) Then
            Combo1(pcmbG_HANBAI_KBN).ListIndex = i
            Exit For
        End If
    
    Next i
                
    '�݌ɐ���
    If Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                               StrConv(ITEMREC.NAIGAI, vbUnicode), _
                               StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
        Exit Function
    
    End If
    Text1(ptxZAIKO_QTY).Text = Format(SUMI_QTY + MI_QTY, "#,##0")
                
    '�W������
    If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
        Text1(ptxG_ST_URITAN).Text = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#,##0.00")
    Else
        Text1(ptxG_ST_URITAN).Text = ""
    End If
    '�W������
    If IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
        Text1(ptxG_ST_SHITAN).Text = Format(CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)), "#,##0.00")
    Else
        Text1(ptxG_ST_SHITAN).Text = ""
    End If
    
    '����
    If IsNumeric(StrConv(P_SHURIAGE_REC.URIAGE_QTY, vbUnicode)) Then
        Text1(ptxURIAGE_QTY).Text = Format(CDbl(StrConv(P_SHURIAGE_REC.URIAGE_QTY, vbUnicode)), "#,##0.00")
    Else
        Text1(ptxURIAGE_QTY).Text = ""
    End If
    '�P��
    If IsNumeric(StrConv(P_SHURIAGE_REC.TANKA, vbUnicode)) Then
        Text1(ptxTANKA).Text = Format(CDbl(StrConv(P_SHURIAGE_REC.TANKA, vbUnicode)), "#,##0.00")
    Else
        Text1(ptxTANKA).Text = ""
    End If
    '���z
    If IsNumeric(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)) Then
        Text1(ptxKINGAKU).Text = Format(CDbl(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)), "#,##0")
    Else
        Text1(ptxKINGAKU).Text = ""
    End If
    '�����
    If IsNumeric(StrConv(P_SHURIAGE_REC.ZEI_KIN, vbUnicode)) Then
        Text1(ptxZEI_KIN).Text = Format(CDbl(StrConv(P_SHURIAGE_REC.ZEI_KIN, vbUnicode)), "#,##0")
    Else
        Text1(ptxZEI_KIN).Text = "0"
    End If

    Item_Disp_Proc = False


End Function

Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   ����f�[�^�̃��X�g�\��
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim com     As Integer
    
Dim Row     As Long
    
    List_Disp_Proc = True

    If Len(Trim(Text1(ptxSEL_KEIJYO_YM).Text)) >= 7 Then
        Call UniCode_Conv(K1_P_SHURIAGE.KEIJYO_YM, Mid(Text1(ptxSEL_KEIJYO_YM).Text, 1, 4) & _
                                                    Mid(Text1(ptxSEL_KEIJYO_YM).Text, 6, 2))
    Else
        Call UniCode_Conv(K1_P_SHURIAGE.KEIJYO_YM, "")
    End If

    Call UniCode_Conv(K1_P_SHURIAGE.G_SYUSHI, "")
    Call UniCode_Conv(K1_P_SHURIAGE.TOKUI_CODE, "")
    Call UniCode_Conv(K1_P_SHURIAGE.URIAGE_DT, "")
    Call UniCode_Conv(K1_P_SHURIAGE.URIAGE_NO, "")

    com = BtOpGetGreater

                                    '�e�[�u�����Z�b�g
    Set URIAGE = Nothing

    Row = 0

    Do
        DoEvents
        sts = BTRV(com, P_SHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), K1_P_SHURIAGE, Len(K1_P_SHURIAGE), 1)
            
        Select Case sts
            Case BtNoErr
            
            
                If Len(Trim(Text1(ptxSEL_KEIJYO_YM).Text)) >= 7 Then
                    If StrConv(P_SHURIAGE_REC.KEIJYO_YM, vbUnicode) <> Mid(Text1(ptxSEL_KEIJYO_YM).Text, 1, 4) & _
                                                               Mid(Text1(ptxSEL_KEIJYO_YM).Text, 6, 2) Then
                        Exit Do
                    End If
                End If
            
            
            Case BtErrEOF
                        
                Exit Do
                        
            
            Case Else
                Call File_Error(sts, com, "���ޔ���f�[�^")
                Exit Function
        
        End Select
    
    
        Row = Row + 1
        
        If Row > Val(txtLOAD_LIMIT.Text) Then       '2017.11.20
            Exit Do
        End If
        
        
        
        URIAGE.ReDim Min_Row, Row, Min_Col, Max_Col
    
        '���㇂
        URIAGE(Row, ColURIAGE_NO) = StrConv(P_SHURIAGE_REC.URIAGE_NO, vbUnicode)
        '����N����
        URIAGE(Row, ColURIAGE_DT) = Mid(StrConv(P_SHURIAGE_REC.URIAGE_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SHURIAGE_REC.URIAGE_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_SHURIAGE_REC.URIAGE_DT, vbUnicode), 7, 2)
        '�����N��
        URIAGE(Row, ColKEIJYO_YM) = Mid(StrConv(P_SHURIAGE_REC.KEIJYO_YM, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_SHURIAGE_REC.KEIJYO_YM, vbUnicode), 5, 2)
    
        '���Ӑ�
        Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHURIAGE_REC.TOKUI_CODE, vbUnicode))
        sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
            
        Select Case sts
            Case BtNoErr
            
            Case BtErrKeyNotFound
                Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_NAME, "")
                        
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�󕥐�Ͻ�")
                Exit Function
        
        End Select
        URIAGE(Row, ColTOKUI_CODE) = Trim(StrConv(P_SHURIAGE_REC.TOKUI_CODE, vbUnicode)) & " " & Trim(StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode))
        '���ޕi��
        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SHURIAGE_REC.HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                        
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�󕥐�Ͻ�")
                Exit Function
        
        End Select
        URIAGE(Row, ColHIN_GAI) = Trim(StrConv(P_SHURIAGE_REC.HIN_GAI, vbUnicode)) & " " & Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
        '���x�P��
        URIAGE(Row, ColG_SYUSHI) = Trim(StrConv(P_SHURIAGE_REC.G_SYUSHI, vbUnicode))
        '�̔��P��
        URIAGE(Row, ColG_HANBAI_KBN) = Trim(StrConv(P_SHURIAGE_REC.G_HANBAI_KBN, vbUnicode))
        '����
        URIAGE(Row, ColURIAGE_QTY) = Format(Val(StrConv(P_SHURIAGE_REC.URIAGE_QTY, vbUnicode)), "#,##0.00")
        '�P��
        URIAGE(Row, ColTANKA) = Format(Val(StrConv(P_SHURIAGE_REC.TANKA, vbUnicode)), "#,##0.00")
        '���z
        URIAGE(Row, ColKINGAKU) = Format(Val(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)), "#,##0")
        '�����
        URIAGE(Row, ColZEI_KIN) = Format(Val(StrConv(P_SHURIAGE_REC.ZEI_KIN, vbUnicode)), "#,##0")
            
    
    
    
        com = BtOpGetNext
    Loop

    Set TDBGrid1.Array = URIAGE
    TDBGrid1.ReBind

    TDBGrid1.Update
    TDBGrid1.MoveFirst

    Call UniCode_Conv(P_SHURIAGE_REC.URIAGE_NO, "")

    List_Disp_Proc = False

End Function

Private Sub Text1_LostFocus(Index As Integer)
Dim ZEI         As Long
Dim wkKINGAKU   As Long

    Select Case Index
    
        Case ptxTOKUI_CODE
        
           Text1(ptxTOKUI_CODE).Text = StrConv(Text1(ptxTOKUI_CODE).Text, vbUpperCase)      '2017.11.20
        
        
        Case ptxTANKA
        
        
            If wkTANKA = Trim(ptxTANKA) Then
                Exit Sub
            End If
                    
            If IsNumeric(Text1(ptxTANKA).Text) And IsNumeric(Text1(ptxURIAGE_QTY).Text) Then
                        
                    
            
                Text1(ptxKINGAKU).Text = Format(CDbl(Text1(ptxTANKA).Text) * CLng(Text1(ptxURIAGE_QTY).Text), "#,##0")
                    
                    
                    
                If CLng(Text1(ptxKINGAKU).Text) >= 0 Then
                    If Format(Text1(ptxURIAGE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                        ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                                CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
                    Else
'                        ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
'                                CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
                        '2019.10.03                             ���o�O
                        ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))
                    End If
                Else
                    
                    wkKINGAKU = CLng(Text1(ptxKINGAKU).Text) * -1
                    
                    If Format(Text1(ptxURIAGE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                        ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                                CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
                    Else
'                        ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
'                                CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
                        '2019.10.03                             ���o�O
                        ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))
                                
                    End If
                    ZEI = ZEI * -1
                End If

                Text1(ptxZEI_KIN).Text = Format(ZEI, "#,##0")
        
        
            End If
        Case ptxURIAGE_QTY
    
    
            If wkQTY = Trim(ptxURIAGE_QTY) Then
                Exit Sub
            End If
                    
            If IsNumeric(Text1(ptxTANKA).Text) And IsNumeric(Text1(ptxURIAGE_QTY).Text) Then
    
                Text1(ptxKINGAKU).Text = Format(CDbl(Text1(ptxTANKA).Text) * CLng(Text1(ptxURIAGE_QTY).Text), "#,##0")
                    
                    
                    
                If CLng(Text1(ptxKINGAKU).Text) >= 0 Then
                    If Format(Text1(ptxURIAGE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                        ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                                CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
                    Else
'                        ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
'                                CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
                        '2019.10.04                             ���o�O
                        ZEI = Int(CDbl(CLng(Text1(ptxKINGAKU).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))
                    End If
                Else
                    
                    wkKINGAKU = CLng(Text1(ptxKINGAKU).Text) * -1
                    
                    If Format(Text1(ptxURIAGE_DT).Text, "YYYYMMDD") < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
                        ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
                                CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
                    Else
'                        ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
'                                CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
                        '2019.10.04                             ���o�O
                        ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                CDbl(CDbl(StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)) / 10))
                    End If
                    ZEI = ZEI * -1
                End If

                Text1(ptxZEI_KIN).Text = Format(ZEI, "#,##0")
    
    
            End If
    End Select

End Sub
