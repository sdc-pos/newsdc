VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form SEI00251 
   Caption         =   "[�����V�X�e��]���i�����ѐ������쐬����"
   ClientHeight    =   11145
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   14580
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
   ScaleHeight     =   11145
   ScaleWidth      =   14580
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton Command1 
      Caption         =   "���i������"
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
      Left            =   4620
      TabIndex        =   24
      Top             =   120
      Width           =   2010
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   7
      Left            =   6090
      TabIndex        =   23
      Top             =   2760
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   6
      Left            =   6090
      TabIndex        =   21
      Top             =   2400
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   5
      Left            =   6090
      TabIndex        =   19
      Top             =   2040
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   4
      Left            =   1995
      TabIndex        =   17
      Top             =   3120
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   3
      Left            =   1995
      TabIndex        =   15
      Top             =   2760
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   2
      Left            =   1995
      TabIndex        =   13
      Top             =   2400
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   1
      Left            =   2940
      TabIndex        =   10
      Top             =   1320
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   0
      Left            =   1365
      TabIndex        =   8
      Top             =   1320
      Width           =   1380
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   1365
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   5
      Top             =   840
      Width           =   2220
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
      Index           =   3
      Left            =   6825
      TabIndex        =   4
      Top             =   120
      Width           =   2010
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   2
      Top             =   10320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�\  ��"
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
      Left            =   2415
      TabIndex        =   1
      Top             =   120
      Width           =   2010
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�W�@�v"
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
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   2010
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   5655
      Left            =   525
      TabIndex        =   3
      Top             =   3840
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   9975
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "���ѓ��t"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "�`�[��"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "�o�א�"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "�i��"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "�i��"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "����"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "���i���H���P��"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "���i���H�����z"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "���i������P��"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "���i��������z"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2328"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2196"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=1561"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1429"
      Splits(0)._ColumnProps(8)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=2170"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=2037"
      Splits(0)._ColumnProps(13)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=2858"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2725"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=2778"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(23)=   "Column(5).Width=1561"
      Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=1429"
      Splits(0)._ColumnProps(26)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(27)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(28)=   "Column(6).Width=3254"
      Splits(0)._ColumnProps(29)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(6)._WidthInPix=3122"
      Splits(0)._ColumnProps(31)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(33)=   "Column(7).Width=3254"
      Splits(0)._ColumnProps(34)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(7)._WidthInPix=3122"
      Splits(0)._ColumnProps(36)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(37)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(38)=   "Column(8).Width=3254"
      Splits(0)._ColumnProps(39)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(8)._WidthInPix=3122"
      Splits(0)._ColumnProps(41)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(42)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(43)=   "Column(9).Width=3254"
      Splits(0)._ColumnProps(44)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(9)._WidthInPix=3122"
      Splits(0)._ColumnProps(46)=   "Column(9)._ColStyle=2"
      Splits(0)._ColumnProps(47)=   "Column(9).Order=10"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �o�S�V�b�N"
      PrintInfos(0).PageFooterFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �o�S�V�b�N"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "��������"
      AllowArrows     =   0   'False
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=�l�r �o�S�V�b�N"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=102,.parent=87"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=99,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=100,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=101,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=106,.parent=87"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=103,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=104,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=105,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=110,.parent=87"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=107,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=108,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=109,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=87"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=87"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=114,.parent=87,.alignment=1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=111,.parent=88"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=112,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=113,.parent=91"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=16,.parent=87,.alignment=1"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=13,.parent=88"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=14,.parent=89"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=15,.parent=91"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=28,.parent=87,.alignment=1"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=88"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=89"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=91"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=20,.parent=87,.alignment=1"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=17,.parent=88"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=18,.parent=89"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=19,.parent=91"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=32,.parent=87,.alignment=1"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=29,.parent=88"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=30,.parent=89"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=31,.parent=91"
      _StyleDefs(76)  =   "Named:id=33:Normal"
      _StyleDefs(77)  =   ":id=33,.parent=0"
      _StyleDefs(78)  =   "Named:id=34:Heading"
      _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(80)  =   ":id=34,.wraptext=-1"
      _StyleDefs(81)  =   "Named:id=35:Footing"
      _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(83)  =   "Named:id=36:Selected"
      _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(85)  =   "Named:id=37:Caption"
      _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(87)  =   "Named:id=38:HighlightRow"
      _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(89)  =   "Named:id=39:EvenRow"
      _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(91)  =   "Named:id=40:OddRow"
      _StyleDefs(92)  =   ":id=40,.parent=33"
      _StyleDefs(93)  =   "Named:id=41:RecordSelector"
      _StyleDefs(94)  =   ":id=41,.parent=34"
      _StyleDefs(95)  =   "Named:id=42:FilterBar"
      _StyleDefs(96)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '����
      Caption         =   "�������z"
      Height          =   375
      Index           =   12
      Left            =   4515
      TabIndex        =   22
      Top             =   2760
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '����
      Caption         =   "����Ŋz"
      Height          =   375
      Index           =   11
      Left            =   4515
      TabIndex        =   20
      Top             =   2400
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '����
      Caption         =   "�������v"
      Height          =   375
      Index           =   10
      Left            =   4515
      TabIndex        =   18
      Top             =   2040
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '����
      Caption         =   "���v"
      Height          =   375
      Index           =   9
      Left            =   630
      TabIndex        =   16
      Top             =   3120
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '����
      Caption         =   "����"
      Height          =   375
      Index           =   8
      Left            =   630
      TabIndex        =   14
      Top             =   2760
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '����
      Caption         =   "�H��"
      Height          =   375
      Index           =   7
      Left            =   630
      TabIndex        =   12
      Top             =   2400
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  '����
      Caption         =   "���i��"
      Height          =   375
      Index           =   6
      Left            =   630
      TabIndex        =   11
      Top             =   2040
      Width           =   3690
   End
   Begin VB.Label Label1 
      Caption         =   "�`"
      Height          =   255
      Index           =   2
      Left            =   2730
      TabIndex        =   9
      Top             =   1440
      Width           =   225
   End
   Begin VB.Label Label1 
      Caption         =   "���t�͈�"
      Height          =   255
      Index           =   1
      Left            =   315
      TabIndex        =   7
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "�d������"
      Height          =   255
      Index           =   0
      Left            =   315
      TabIndex        =   6
      Top             =   960
      Width           =   1065
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "�����I��"
      Index           =   0
      Begin VB.Menu SHORI 
         Caption         =   "�W�v"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "EXCEL(�\��)"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "EXCEL(�o�ז���)"
         Index           =   2
      End
      Begin VB.Menu SHORI 
         Caption         =   "EXCEL(���ɖ���)"
         Index           =   3
      End
      Begin VB.Menu SHORI 
         Caption         =   "EXCEL(�Ǖi�ԕi����)"
         Index           =   4
      End
      Begin VB.Menu SHORI 
         Caption         =   "�I��"
         Index           =   5
      End
      Begin VB.Menu SHORI 
         Caption         =   "��ʈ��"
         Index           =   6
      End
   End
End
Attribute VB_Name = "SEI00251"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const pcmbSHIMUKE% = 0          '�d������

Private Const ptxS_Date% = 0            '���t�͈́@�J�n
Private Const ptxE_Date% = 1            '���t�͈́@�I��


Private Const ptxGK_SYOHIN_KOURYO% = 2  '���i���@�H��
Private Const ptxGK_SYOHIN_SHIZAI% = 3  '���i���@����
Private Const ptxGK_SYOHIN% = 4         '���i���@���v

Private Const ptxGK_SEIKYU% = 5        '�������v
Private Const ptxGK_ZEI_KIN% = 6       '����Ŋz
Private Const ptxGK_SEIKYU_KIN% = 7    '�������z





Dim SEIKYU  As New XArrayDB

Private Const Min_Row% = 1              '�ŏ��s��

Dim Max_Row    As Integer               '�O���b�h�ő�\������


Private Const Min_Col% = 0              '�ŏ���
Private Const Max_Col% = 9             '�ő��

Private Const ColSYUKA_YMD% = 0         '�`�[���t

Private Const ColHIN_GAI% = 3           '�i��
Private Const ColHIN_NAME% = 4          '�i��



Private Const ColSURYO% = 5             '����
Private Const ColSYOHIN_KOURYO_T% = 6   '���i���@�H��
Private Const ColSYOHIN_KOURYO_K% = 7   '���i���@�H��
Private Const ColSYOHIN_SHIZAI_T% = 8   '���i���@����
Private Const ColSYOHIN_SHIZAI_K% = 9   '���i���@����

Private GK_SYOHIN_KOURYO    As Long     '���i���@�H��
Private GK_SYOHIN_SHIZAI    As Long     '���i���@����


Dim Name1               As String
Dim Name2               As String
    
Dim ITEM                As String

Dim ADDR1               As String
Dim ADDR2               As String

Dim SYAMEI              As String

Dim BIKOU1              As String
Dim BIKOU2              As String
Dim BIKOU3              As String

Dim SHIMEBI             As String


Private Type MEISAI_TBL_tag
    HIN_NAME    As String               '�������i�\���j �i��
    TEKIYO      As String               '�������i�\���j �E�v
    KINGAKU     As Long                 '�������i�\���j ���z
End Type
Private MEISAI_TBL()    As MEISAI_TBL_tag








Private Sub Command1_Click(Index As Integer)
Dim ans As Integer

    Select Case Index
        Case 0                          '�W�v
        
            If Update_Proc() Then
                Unload Me
            End If
        
        Case 1                          'EXCEL�o��(�\��)
        
            If COVER_Proc() Then
                Unload Me
            End If
        
        Case 2                          'EXCEL�o��
        
            If DETAIL_Proc() Then
                Unload Me
            End If
        
        Case 3                          '�I��
            Unload Me
        Case Else
            Beep
    End Select

End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
Private Sub Form_Load()

Dim c           As String * 128
Dim sts         As Integer

Dim S_DATE      As String
Dim E_DATE      As String
Dim S_YY        As String * 4
Dim S_MM        As String * 2
Dim S_DD        As String * 2
    
Dim i           As Integer
Dim j           As Integer
    
    
    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If


    
    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[�����V�X�e��]���i�����ѐ������쐬����", Me.hwnd, 0)
    '�y�C���������
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SETPARTS, 0, -1)


    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                


    Max_Row = 9999
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�R�[�h�}�X�^�n�o�d�m
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�Ǘ��}�X�^�n�o�d�m
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '�w�}�\�f�[�^(�e)�n�o�d�m
    If P_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�w�}�\��������n�o�d�m
    If P_SUKEIRE_Open(BtOpenNomal) Then
        Unload Me
    End If



    '�Ǘ��}�X�^�ǂݍ���
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
            
        Case Else
            Unload Me
    End Select

    If JGYOB_TB_Set(1) Then      '���ƕ��̊l��
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    '����Ͻ���`
    Call P_CODE_TBL_Proc
    '�d������̃Z�b�g
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD, 0) Then
        Unload Me
    End If
    '�������i�\���j�̐ݒ荀�ڊl��
    If GetIni(App.EXEName, "Name1", App.EXEName, c) Then
        Name1 = ""
    Else
        Name1 = Trim(c)
    End If
    
    If GetIni(App.EXEName, "Name2", App.EXEName, c) Then
        Name2 = ""
    Else
        Name2 = Trim(c)
    End If
    If GetIni(App.EXEName, "Item", App.EXEName, c) Then
        ITEM = ""
    Else
        ITEM = Trim(c)
    End If
    If GetIni(App.EXEName, "ADDR1", App.EXEName, c) Then
        ADDR1 = ""
    Else
        ADDR1 = Trim(c)
    End If
    If GetIni(App.EXEName, "ADDR2", App.EXEName, c) Then
        ADDR2 = ""
    Else
        ADDR2 = Trim(c)
    End If
    If GetIni(App.EXEName, "SYAMEI", App.EXEName, c) Then
        SYAMEI = ""
    Else
        SYAMEI = Trim(c)
    End If
    If GetIni(App.EXEName, "BIKOU1", App.EXEName, c) Then
        BIKOU1 = ""
    Else
        BIKOU1 = Trim(c)
    End If
    If GetIni(App.EXEName, "BIKOU2", App.EXEName, c) Then
        BIKOU2 = ""
    Else
        BIKOU2 = Trim(c)
    End If
    If GetIni(App.EXEName, "BIKOU3", App.EXEName, c) Then
        BIKOU3 = ""
    Else
        BIKOU3 = Trim(c)
    End If
    If GetIni(App.EXEName, "SHIMEBI", App.EXEName, c) Then
        SHIMEBI = ""
    Else
        SHIMEBI = Trim(c)
    End If
    
        
    i = -1
    j = 1
    Do
        If GetIni(App.EXEName, "HIN_NAME" & Format(j, "00"), App.EXEName, c) Then
            Exit Do
        End If
              
        i = i + 1
        ReDim Preserve MEISAI_TBL(0 To i)
        MEISAI_TBL(i).HIN_NAME = Trim(c)
        
        If GetIni(App.EXEName, "TEKIYO" & Format(j, "00"), App.EXEName, c) Then
            
                        
            
            MEISAI_TBL(i).TEKIYO = ""
        Else
            MEISAI_TBL(i).TEKIYO = Trim(c)
        End If
    
        j = j + 1
    Loop
    
    
    Combo1(pcmbSHIMUKE).ListIndex = 0

    
    E_DATE = Format(Now, "YYYY/MM/DD")
    S_DATE = DateAdd("m", -1, Left(E_DATE, 8) & SHIMEBI)
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


    Text1(ptxS_Date).Text = S_DATE
    Text1(ptxE_Date).Text = E_DATE
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
                                            '�R�[�h�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�R�[�h�}�X�^")
        End If
    End If
                                            '�Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�Ǘ��}�X�^")
        End If
    End If
                                            '�w�}�\�f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�w�}�\�f�[�^")
        End If
    End If
                                            '��������b�k�n�r�d
    sts = BTRV(BtOpClose, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "��������f�[�^")
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
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    SEI00251.MousePointer = vbHourglass


    TDBGrid1.Enabled = False


    Call Ctrl_Lock(SEI00251)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(SEI00251)

    TDBGrid1.Enabled = True

    SEI00251.MousePointer = vbDefault

End Sub


Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   �����W�v����
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim ans         As Integer
    
Dim Row         As Long
    
Dim com         As Integer
    
Dim GK_ZEI_KIN  As Long
    
    
Dim Skip_F      As Boolean
    
Dim i           As Integer
Dim j           As Integer
    
    
    Update_Proc = True
    
    Call Input_Lock
                                    
                        '�W�v�l�@�N���A�[

    GK_SYOHIN_KOURYO = 0
    GK_SYOHIN_SHIZAI = 0
                                    
                                    
                        '�e�[�u�����Z�b�g
    Set SEIKYU = Nothing
    Row = Min_Row - 1
    
    
    
    '------------------------------------------------------------------------   '��������̓ǂݍ���
    Call UniCode_Conv(K1_P_SUKEIRE.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(K1_P_SUKEIRE.UKEIRE_DT, Format(Text1(ptxS_Date), "YYYYMMDD"))
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        sts = BTRV(com, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K1_P_SUKEIRE, Len(K1_P_SUKEIRE), 1)
        Select Case sts
            Case BtNoErr
            
                If Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) <> StrConv(P_SUKEIRE_REC.SHIMUKE_CODE, vbUnicode) Then
                    Exit Do
                End If
            
                If Format(Text1(ptxE_Date), "YYYYMMDD") < StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "�������")
                Exit Function
        End Select

        Skip_F = False
    
    
        Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_NO, StrConv(P_SUKEIRE_REC.SHIJI_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
        Select Case sts
            Case BtNoErr
            
            
                If StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
            
                    Skip_F = True
                
                End If
            
            Case BtErrKeyNotFound
                Skip_F = True
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "���i���w�}�\(�e)")
                Exit Function
        End Select
    
        If Not Skip_F Then
    
            Row = Row + 1
    
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
        
        
        End If
        
        com = BtOpGetNext
    Loop




'    SEIKYU.QuickSort 1, SEIKYU.UpperBound(1), ColSYUKA_YMD, 0, XTYPE_STRING
        


    Set TDBGrid1.Array = SEIKYU
    
    
'    TDBGrid1.Bookmark = Null
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.ScrollBars = dbgAutomatic









    Text1(ptxGK_SYOHIN_KOURYO).Text = Format(GK_SYOHIN_KOURYO, "#,##0")
    Text1(ptxGK_SYOHIN_SHIZAI).Text = Format(GK_SYOHIN_SHIZAI, "#,##0")
    Text1(ptxGK_SYOHIN).Text = Format(GK_SYOHIN_KOURYO + GK_SYOHIN_SHIZAI, "#,##0")

        Text1(ptxGK_SEIKYU).Text = Format(GK_SYOHIN_KOURYO + _
                                        GK_SYOHIN_SHIZAI, "#,##0")



    GK_ZEI_KIN = Fix((CDbl(Text1(ptxGK_SEIKYU).Text) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                            CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10)


    Text1(ptxGK_ZEI_KIN).Text = Format(GK_ZEI_KIN, "#,##0")

    Text1(ptxGK_SEIKYU_KIN).Text = Format(CDbl(Text1(ptxGK_SEIKYU).Text) + GK_ZEI_KIN, "#,##0")

    Call Input_UnLock




    Update_Proc = False


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
        
        
        
        Combo1(Index).AddItem StrConv(P_CODEREC.C_RNAME, vbUnicode) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), Key_Len) & wkOption
        
        
        com = BtOpGetNext
    
    Loop

    Code_Set_Proc = False
    



End Function


Private Sub SHORI_MENU_Click(Index As Integer)

    Select Case Index
        Case 0 To 3
            Command1(Index).Value = True

        Case 2      '��ʈ��
        
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
Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           ���i���f�[�^--��Grid
'----------------------------------------------------------------------------

Dim sts As Integer

    
    Grid_Set_Proc = True

    

    SEIKYU.ReDim Min_Row, Row, Min_Col, Max_Col
    
    '���ѓ��t
    SEIKYU(Row, ColSYUKA_YMD) = Mid(StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode), 7, 2)
    
    
    
    '�i��
    SEIKYU(Row, ColHIN_GAI) = StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)
    
    '����
    SEIKYU(Row, ColSURYO) = Format(CLng(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode)), "#,##0")
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        
        
        Case BtErrKeyNotFound
            
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
            
        
            Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
            Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    
    End Select
    '�i��
    SEIKYU(Row, ColHIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
    
    '���i���@�H��
    If IsNumeric(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) Then
        SEIKYU(Row, ColSYOHIN_KOURYO_T) = Format(CDbl(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)), "#,##0.00")
        SEIKYU(Row, ColSYOHIN_KOURYO_K) = Format(Fix(CDbl(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode)) * CDbl(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) + 0.9), "#,##0")
        GK_SYOHIN_KOURYO = GK_SYOHIN_KOURYO + Fix(CDbl(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode)) * CDbl(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) + 0.9)
    Else
        SEIKYU(Row, ColSYOHIN_KOURYO_T) = Format(0, "0.00")
        SEIKYU(Row, ColSYOHIN_KOURYO_K) = 0
    End If
    '���i���@����
    If IsNumeric(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)) Then
        SEIKYU(Row, ColSYOHIN_SHIZAI_T) = Format(CDbl(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)), "#,##0.00")
        SEIKYU(Row, ColSYOHIN_SHIZAI_K) = Format(Fix(CDbl(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode)) * CDbl(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)) + 0.9), "#,##0")
        GK_SYOHIN_SHIZAI = GK_SYOHIN_SHIZAI + Fix(CDbl(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode)) * CDbl(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)) + 0.9)
    Else
        SEIKYU(Row, ColSYOHIN_SHIZAI_T) = Format(0, "0.00")
        SEIKYU(Row, ColSYOHIN_SHIZAI_K) = 0
    End If
    
    
    
    
    Grid_Set_Proc = False
End Function




Private Function COVER_Proc() As Integer
'----------------------------------------------------------------------------
'                   �d�w�b�d�k�i�\���j�o��
'----------------------------------------------------------------------------
Dim i                   As Integer
Dim j                   As Integer
Dim Row                 As Integer
    
    
Dim sts                 As Integer
Dim com                 As Integer
    
Dim End_Date            As String


Dim GK_KINGAKU          As Long
Dim WK_TANKA            As Double
Dim ZEI_KIN             As Long

    
Dim Skip_F              As Boolean


Dim excelApplication    As excel.Application
'Dim excelWorkBooks      As excel.Workbooks
Dim excelWorkBook       As excel.Workbook
Dim excelSheet          As excel.Worksheet

    

    COVER_Proc = True
    
    Call Input_Lock
    



    
    Set excelApplication = CreateObject("Excel.Application")
    excelApplication.Visible = True


    
    Set excelWorkBook = excelApplication.Workbooks.Add
'    Set excelSheet = excelWorkBook.Worksheets.Add
    Set excelSheet = excelWorkBook.Worksheets(1)
    

    
    excelApplication.StandardFontSize = 11
    
    excelApplication.StandardFont = "�l�r�@�S�V�b�N"

    
    
'    excelSheet.Application.Select
'    With excelSheet.Application.Selection.Font
'        .NAME = "�l�r�@�S�V�b�N"
'        .FontStyle = "�W��"
'        .Size = 11
'    End With
    
    '�y�[�W�ݒ�
    With excelSheet.Application.ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    
    '��̕�
    excelSheet.Application.Columns(1).Select
    excelSheet.Application.Selection.ColumnWidth = 7.25
    excelSheet.Application.Columns(2).Select
    excelSheet.Application.Selection.ColumnWidth = 36.13
    excelSheet.Application.Columns(3).Select
    excelSheet.Application.Selection.ColumnWidth = 5.38
    excelSheet.Application.Columns(4).Select
    excelSheet.Application.Selection.ColumnWidth = 12.13
    excelSheet.Application.Columns(5).Select
    excelSheet.Application.Selection.ColumnWidth = 13.38
    excelSheet.Application.Columns(6).Select
    excelSheet.Application.Selection.ColumnWidth = 15
    
    '�s�̕�
    excelSheet.Application.Rows(1).Select
    excelSheet.Application.Selection.RowHeight = 24
    excelSheet.Application.Rows("3:4").Select
    excelSheet.Application.Selection.RowHeight = 14.25
    excelSheet.Application.Rows(12).Select
    excelSheet.Application.Selection.RowHeight = 27
    excelSheet.Application.Rows("14:31").Select
    excelSheet.Application.Selection.RowHeight = 27
    
    '�Z���̌���
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(1, 6)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(1, 6)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(1, 6)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(1, 6)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "�l�r�@�S�V�b�N"
        .FontStyle = "����"
        .Size = 20
        .Underline = xlUnderlineStyleSingle
    End With
    excelSheet.Application.Cells(1, 1).Value = "�� �� ��"
    
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 3)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 3)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 3)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 3)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "�l�r�@�S�V�b�N"
        .FontStyle = "�W��"
        .Size = 14
    End With
    excelSheet.Application.Cells(12, 1).Value = "�� �v �� �z"
    
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 4), excelSheet.Application.Cells(12, 6)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 4), excelSheet.Application.Cells(12, 6)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 4), excelSheet.Application.Cells(12, 6)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 4), excelSheet.Application.Cells(12, 6)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "�l�r�@�S�V�b�N"
        .FontStyle = "�W��"
        .Size = 14
    End With
    excelSheet.Application.Cells(12, 4).Value = ""
    
    excelSheet.Application.Range(excelSheet.Application.Cells(29, 1), excelSheet.Application.Cells(29, 4)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(29, 1), excelSheet.Application.Cells(29, 4)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(29, 1), excelSheet.Application.Cells(29, 4)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(29, 1), excelSheet.Application.Cells(29, 4)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "�l�r�@�S�V�b�N"
        .FontStyle = "�W��"
        .Size = 11
    End With
    excelSheet.Application.Cells(29, 1).Value = "�� �� �� �� �z"
    
    excelSheet.Application.Range(excelSheet.Application.Cells(30, 1), excelSheet.Application.Cells(30, 4)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(30, 1), excelSheet.Application.Cells(30, 4)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(30, 1), excelSheet.Application.Cells(30, 4)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(30, 1), excelSheet.Application.Cells(30, 4)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "�l�r�@�S�V�b�N"
        .FontStyle = "�W��"
        .Size = 11
    End With
    excelSheet.Application.Cells(30, 1).Value = "��    ��    ��"
    
    excelSheet.Application.Range(excelSheet.Application.Cells(31, 1), excelSheet.Application.Cells(31, 4)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(31, 1), excelSheet.Application.Cells(31, 4)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(31, 1), excelSheet.Application.Cells(31, 4)).MergeCells = True
    excelSheet.Application.Range(excelSheet.Application.Cells(31, 1), excelSheet.Application.Cells(31, 4)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "�l�r�@�S�V�b�N"
        .FontStyle = "�W��"
        .Size = 11
    End With
    excelSheet.Application.Cells(31, 1).Value = "�� �� �� �� �z"
    
    
    
    '�r��
    excelSheet.Application.Range(excelSheet.Application.Cells(12, 1), excelSheet.Application.Cells(12, 6)).Select
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

    excelSheet.Application.Range(excelSheet.Application.Cells(14, 1), excelSheet.Application.Cells(31, 6)).Select
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
    With excelSheet.Application.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With excelSheet.Application.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    '�Œ荀�ځi���o���j
    excelSheet.Application.Range(excelSheet.Application.Cells(14, 1), excelSheet.Application.Cells(14, 6)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(14, 1), excelSheet.Application.Cells(14, 6)).VerticalAlignment = xlCenter
    excelSheet.Application.Cells(14, 1).Value = "��/��"
    excelSheet.Application.Cells(14, 2).Value = "�i     ��"
    excelSheet.Application.Cells(14, 3).Value = "�� ��"
    excelSheet.Application.Cells(14, 4).Value = "�P  ��"
    excelSheet.Application.Cells(14, 5).Value = "���@�z"
    excelSheet.Application.Cells(14, 6).Value = "�E�@�v"
    '�Œ荀�ځiINI�j
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 6), excelSheet.Application.Cells(2, 6)).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(2, 6).Value = Left(Format(Now, "YYYY�NMM��DD��"), 8) & SHIMEBI & "��"
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(3, 1)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "�l�r�@�S�V�b�N"
        .FontStyle = "����"
        .Size = 11
    End With
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 6), excelSheet.Application.Cells(9, 6)).HorizontalAlignment = xlRight
    
    excelSheet.Application.Cells(3, 1).Value = Name1
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 1), excelSheet.Application.Cells(4, 1)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "�l�r�@�S�V�b�N"
        .FontStyle = "����"
        .Size = 11
        .Underline = xlUnderlineStyleSingle
    End With
    excelSheet.Application.Cells(4, 1).Value = Name2
    
    excelSheet.Application.Range(excelSheet.Application.Cells(8, 2), excelSheet.Application.Cells(8, 2)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "�l�r�@�S�V�b�N"
        .FontStyle = "����"
        .Size = 11
        .Underline = xlUnderlineStyleSingle
    End With
    excelSheet.Application.Cells(8, 2).Value = ITEM
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(4, 6), excelSheet.Application.Cells(7, 6)).Select
    With excelSheet.Application.Selection.Font
        .NAME = "�l�r�@�S�V�b�N"
        .FontStyle = "����"
        .Size = 9
    End With
    excelSheet.Application.Cells(4, 6).Value = ADDR1
    excelSheet.Application.Cells(5, 6).Value = ADDR2
    excelSheet.Application.Cells(6, 6).Value = SYAMEI
    excelSheet.Application.Cells(7, 6).Value = BIKOU1
    excelSheet.Application.Cells(8, 6).Value = BIKOU2
    excelSheet.Application.Cells(9, 6).Value = BIKOU3
    
    
    
    For i = 0 To UBound(MEISAI_TBL)
        MEISAI_TBL(i).KINGAKU = 0
    Next i
    
    
    
    
    
     '------------------------------------------------------------------------   '��������̓ǂݍ���
    Call UniCode_Conv(K1_P_SUKEIRE.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(K1_P_SUKEIRE.UKEIRE_DT, Format(Text1(ptxS_Date), "YYYYMMDD"))
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        sts = BTRV(com, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K1_P_SUKEIRE, Len(K1_P_SUKEIRE), 1)
        Select Case sts
            Case BtNoErr
            
            
                If Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) <> StrConv(P_SUKEIRE_REC.SHIMUKE_CODE, vbUnicode) Then
                    Exit Do
                End If
            
            
            
                If Format(Text1(ptxE_Date), "YYYYMMDD") < StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "�������")
                Exit Function
        End Select

        Skip_F = False
    
    
        Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_NO, StrConv(P_SUKEIRE_REC.SHIJI_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
        Select Case sts
            Case BtNoErr
            
            
                If StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
            
                    Skip_F = True
                
                End If
            
            Case BtErrKeyNotFound
                Skip_F = True
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "���i���w�}�\(�e)")
                Exit Function
        End Select
    
    
                
    
    
        If Not Skip_F Then
    
            If Cover_Total_Proc(1) Then
                Exit Function
        
            End If
        
        
        End If
        
        com = BtOpGetNext
    Loop
   
    
    
    
    
    
    
    
    
    For i = 0 To UBound(MEISAI_TBL)
        '���^��
        excelSheet.Application.Range(excelSheet.Application.Cells(15 + i, 1), excelSheet.Application.Cells(15 + i, 1)).HorizontalAlignment = xlCenter
        excelSheet.Application.Range(excelSheet.Application.Cells(15 + i, 1), excelSheet.Application.Cells(15 + i, 1)).Select
        excelSheet.Application.Selection.NumberFormatLocal = "@"
        excelSheet.Application.Selection.HorizontalAlignment = xlCenter
        excelSheet.Application.Cells(15 + i, 1).Value = Format(CInt(Mid(Format(Now, "YYYYMMDD"), 5, 2)), "#") & "/" & SHIMEBI
        '�i��
        excelSheet.Application.Cells(15 + i, 2).Value = Trim(MEISAI_TBL(i).HIN_NAME) & "(" & Text1(ptxS_Date).Text & "�`" & Text1(ptxS_Date).Text & ")"
        '����
        excelSheet.Application.Range(excelSheet.Application.Cells(15 + i, 3), excelSheet.Application.Cells(15 + i, 3)).Select
        excelSheet.Application.Selection.NumberFormatLocal = "#,##0"
        excelSheet.Application.Cells(15 + i, 3).Value = 1
        '�P���`���z
        excelSheet.Application.Range(excelSheet.Application.Cells(15 + i, 4), excelSheet.Application.Cells(15 + i, 5)).Select
        excelSheet.Application.Selection.NumberFormatLocal = "#,##0"
        excelSheet.Application.Cells(15 + i, 4).Value = MEISAI_TBL(i).KINGAKU
        excelSheet.Application.Cells(15 + i, 5).Value = MEISAI_TBL(i).KINGAKU
        '�E�v
        excelSheet.Application.Cells(15 + i, 6).Value = Trim(MEISAI_TBL(i).TEKIYO)
    
    Next i
    
    
    GK_KINGAKU = 0
    For i = 0 To UBound(MEISAI_TBL)
        GK_KINGAKU = GK_KINGAKU + MEISAI_TBL(i).KINGAKU
    Next i
    
    
    
    '�Ŕ������z
    excelSheet.Application.Range(excelSheet.Application.Cells(29, 5), excelSheet.Application.Cells(31, 5)).Select
    excelSheet.Application.Selection.NumberFormatLocal = "#,##0;""�� ""#,##0"
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[-14]C:R[-1]C)"
    '�����
    ZEI_KIN = Fix((GK_KINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
                                CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10)
    excelSheet.Application.Cells(30, 5).Value = ZEI_KIN
    '�ō��݋��z
    excelSheet.Application.Cells(31, 5).Value = GK_KINGAKU + ZEI_KIN
    '���v���z
    excelSheet.Application.Cells(12, 4).Value = Format(GK_KINGAKU + ZEI_KIN, "\\#,##0")





'    excelApplication.Quit

    Set excelSheet = Nothing
    Set excelWorkBook = Nothing
'    Set excelWorkBooks = Nothing
    Set excelApplication = Nothing


    
    Call Input_UnLock
    COVER_Proc = False
    

End Function


Private Function DETAIL_Proc() As Integer
'----------------------------------------------------------------------------
'                   �d�w�b�d�k�o��
'----------------------------------------------------------------------------


Dim Row                 As Integer
    
    
Dim sts                 As Integer
Dim com                 As Integer
    
Dim End_Date            As String

Dim s_test_now          As String

Dim Skip_F              As Boolean


Dim excelApplication    As excel.Application
'Dim excelWorkBooks      As excel.Workbooks
Dim excelWorkBook       As excel.Workbook
Dim excelSheet          As excel.Worksheet
    
    
    
s_test_now = Format(Now, "YYYY/MM/DD HH:MM:SS")
    
    DETAIL_Proc = True
    
    Call Input_Lock
    
    Set excelApplication = CreateObject("Excel.Application")
    excelApplication.Visible = True

        
    
    
    Set excelWorkBook = excelApplication.Workbooks.Add
    
    
    Set excelSheet = excelWorkBook.Worksheets(1)
    
    excelApplication.StandardFontSize = 11
    
    excelApplication.StandardFont = "�l�r�@�S�V�b�N"
    
    

    excelSheet.Application.Calculation = xlManual
    excelSheet.Application.MaxChange = 0.001

    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(1, 4)).Select
    With excelSheet.Application.Selection.Font
        .Size = 16
    End With
    excelSheet.Application.Cells(1, 1).Value = "���i�����і��ו\" & _
                                    Trim(StrConv(P_KANRIREC.CENTER_NAME, vbUnicode)) & _
                                    "�i" & StrConv(Text1(ptxS_Date).Text, vbWide) & "�`" & _
                                    StrConv(Text1(ptxE_Date).Text, vbWide) & "�j"
    
    
    
    '��̕�
    excelSheet.Application.Columns(1).Select
    excelSheet.Application.Selection.ColumnWidth = 4.88
    '�Z���̌���
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 6), excelSheet.Application.Cells(2, 7)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 6), excelSheet.Application.Cells(2, 7)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 6), excelSheet.Application.Cells(2, 7)).MergeCells = True
   
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 8), excelSheet.Application.Cells(2, 9)).HorizontalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 8), excelSheet.Application.Cells(2, 9)).VerticalAlignment = xlCenter
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 8), excelSheet.Application.Cells(2, 9)).MergeCells = True
    
    '�Q�s�ڌ��o���ݒ�
    excelSheet.Application.Cells(2, 6).Value = "���H��"
    excelSheet.Application.Cells(2, 8).Value = "������"
    '�R�s�ڌ��o���ݒ�
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 1), excelSheet.Application.Cells(3, 1)).HorizontalAlignment = xlRight
    excelSheet.Application.Cells(3, 1).Value = "��"
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 2), excelSheet.Application.Cells(3, 9)).HorizontalAlignment = xlCenter
    excelSheet.Application.Cells(3, 2).Value = "���ѓ�"
    excelSheet.Application.Cells(3, 3).Value = "�i��"
    excelSheet.Application.Cells(3, 4).Value = "�i��"
    excelSheet.Application.Cells(3, 5).Value = "����"
    excelSheet.Application.Cells(3, 6).Value = "�P��"
    excelSheet.Application.Cells(3, 7).Value = "���z"
    excelSheet.Application.Cells(3, 8).Value = "�P��"
    excelSheet.Application.Cells(3, 9).Value = "���z"
    '���o���@�r��
    excelSheet.Application.Range(excelSheet.Application.Cells(2, 1), excelSheet.Application.Cells(3, 9)).Select
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
    
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 6), excelSheet.Application.Cells(3, 7)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
   
    excelSheet.Application.Range(excelSheet.Application.Cells(3, 8), excelSheet.Application.Cells(3, 9)).Select
    excelSheet.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excelSheet.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With excelSheet.Application.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    Row = 3
        
    
    
    '------------------------------------------------------------------------   '��������̓ǂݍ���
    Call UniCode_Conv(K1_P_SUKEIRE.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
    Call UniCode_Conv(K1_P_SUKEIRE.UKEIRE_DT, Format(Text1(ptxS_Date), "YYYYMMDD"))
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        sts = BTRV(com, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K1_P_SUKEIRE, Len(K1_P_SUKEIRE), 1)
        Select Case sts
            Case BtNoErr
            
                If Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2) <> StrConv(P_SUKEIRE_REC.SHIMUKE_CODE, vbUnicode) Then
                    Exit Do
                End If
            
            
            
                If Format(Text1(ptxE_Date), "YYYYMMDD") < StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "�������")
                Exit Function
        End Select

        Skip_F = False
    
    
        Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_NO, StrConv(P_SUKEIRE_REC.SHIJI_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
        Select Case sts
            Case BtNoErr
            
            
                If StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
            
                    Skip_F = True
                
                End If
            
            Case BtErrKeyNotFound
                Skip_F = True
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "���i���w�}�\(�e)")
                Exit Function
        End Select
    
        If Not Skip_F Then
    
            Row = Row + 1
    
    

    
            If Excel_Set_Proc(Row, excelApplication, excelWorkBook, excelSheet) Then
                Exit Function
            End If
        
        
        End If
        
        com = BtOpGetNext
    Loop
    
    
    Row = Row + 1
    
    '���v
    excelSheet.Application.Cells(Row, 1).Value = "���v"
    
    '���ʍ��v
    excelSheet.Application.Cells(Row, 5).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 5), excelSheet.Application.Cells(Row, 5)).Select
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C)"
    
    
    '���H���@���z���v
    excelSheet.Application.Cells(Row, 7).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 7), excelSheet.Application.Cells(Row, 7)).Select
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C)"
    
    '������@���z���v
    excelSheet.Application.Cells(Row, 9).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 9), excelSheet.Application.Cells(Row, 9)).Select
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=SUM(R[" & ((Row - 1) * -1) + 3 & "]C:R[-1]C)"
    '�r��

    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row, 9)).Select
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
    
    
    excelSheet.Application.Range(excelSheet.Application.Cells(1, 1), excelSheet.Application.Cells(Row, 9)).Select
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
    With excelSheet.Application.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    
    
    
    excelSheet.Application.Columns("B:U").EntireColumn.AutoFit
    
    
    
        
    
    



    Set excelSheet = Nothing
    
    Set excelWorkBook = Nothing
    

    
    
    Set excelApplication = Nothing


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        s_test_now & " " & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
    
    Call Input_UnLock
    DETAIL_Proc = False
    

End Function
Private Function Excel_Set_Proc(Row As Integer, excelApplication As excel.Application, excelWorkBook As excel.Workbook, excelSheet As excel.Worksheet) As Integer
'----------------------------------------------------------------------------
'           ���уf�[�^--��EXCEL
'----------------------------------------------------------------------------
Dim sts     As Integer
    
    Excel_Set_Proc = True
        
    '�Z���̏����ݒ�
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 1), excelSheet.Application.Cells(Row, 1)).Select
    excelSheet.Application.ActiveCell.FormulaR1C1 = "=ROW()-3"
    
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 2), excelSheet.Application.Cells(Row, 5)).Select
    excelSheet.Application.Selection.NumberFormatLocal = "@"
    excelSheet.Application.Range(excelSheet.Application.Cells(Row, 10), excelSheet.Application.Cells(Row, 10)).Select
    excelSheet.Application.Selection.NumberFormatLocal = "@"
    
    
    excelSheet.Application.Cells(Row, 5).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Cells(Row, 6).NumberFormatLocal = "#,##0.00_ "
    
    excelSheet.Application.Cells(Row, 7).NumberFormatLocal = "#,##0_ "
    excelSheet.Application.Cells(Row, 8).NumberFormatLocal = "#,##0.00_ "
    excelSheet.Application.Cells(Row, 9).NumberFormatLocal = "#,##0_ "
    '���ѓ�
    excelSheet.Application.Cells(Row, 2).Value = Mid(StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode), 7, 2)
    '�i��
    excelSheet.Application.Cells(Row, 3).Value = Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))
    
    '�i��
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
        
        
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        
            Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
            Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    End Select
    
    
    excelSheet.Application.Cells(Row, 4).Value = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
    '����
    excelSheet.Application.Cells(Row, 5).Value = CLng(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode))
    '���H���@�P��
    If IsNumeric(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) Then
        excelSheet.Application.Cells(Row, 6).Value = CDbl(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode))
    Else
        excelSheet.Application.Cells(Row, 6).Value = 0
    End If
    '���H���@���z
    excelSheet.Application.Cells(Row, 7).Value = Fix(CDbl(excelSheet.Application.Cells(Row, 5).Value) * CDbl(excelSheet.Application.Cells(Row, 6).Value) + 0.9)
    '������@�P��
    If IsNumeric(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)) Then
        excelSheet.Application.Cells(Row, 8).Value = CDbl(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode))
    Else
        excelSheet.Application.Cells(Row, 8).Value = 0
    End If
    '������@���z
    excelSheet.Application.Cells(Row, 9).Value = Fix(CDbl(excelSheet.Application.Cells(Row, 5).Value) * CDbl(excelSheet.Application.Cells(Row, 8).Value) + 0.9)
    '�r��
'    excelSheet.Range(Cells(Row, 1), Cells(Row, 21)).Select
'    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'    With Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With Selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With Selection.Borders(xlEdgeBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With Selection.Borders(xlInsideVertical)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With























    Excel_Set_Proc = False

End Function

Private Function Cover_Total_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'           ���i�����уf�[�^��苾�p�̋��z�W�v
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim INV_F  As Boolean
    
    
    Cover_Total_Proc = True
    
    
    Select Case Mode
        Case 1
            
            
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))
            
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                
                
                    Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                
                    Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "00000000.00")
                    Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "00000000.00")
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Function
            End Select
            
            
            '���i���@�H��
            If IsNumeric(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) Then
                MEISAI_TBL(0).KINGAKU = MEISAI_TBL(0).KINGAKU + Fix(CDbl(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode)) * CDbl(StrConv(ITEMREC.S_KOUSU_BAIKA, vbUnicode)) + 0.9)
            Else
            End If
            '���i���@����
            If IsNumeric(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)) Then
                MEISAI_TBL(1).KINGAKU = MEISAI_TBL(1).KINGAKU + Fix(CDbl(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode)) * CDbl(StrConv(ITEMREC.S_SHIZAI_BAIKA, vbUnicode)) + 0.9)
            Else
            End If
    
    
    End Select
    
    Cover_Total_Proc = False

End Function

