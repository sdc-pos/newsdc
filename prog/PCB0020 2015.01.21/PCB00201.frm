VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PCB00201 
   Caption         =   "�i�ڃ}�X�^�����e�i���X(���W���[�����i�p)"
   ClientHeight    =   10590
   ClientLeft      =   2025
   ClientTop       =   -3210
   ClientWidth     =   15270
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
   ScaleHeight     =   10590
   ScaleWidth      =   15270
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000A&
      Height          =   375
      Index           =   13
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000A&
      Height          =   375
      Index           =   1
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   720
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000A&
      Height          =   375
      Index           =   12
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      Height          =   375
      Index           =   11
      Left            =   3480
      TabIndex        =   29
      Top             =   6000
      Width           =   600
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      Height          =   375
      Index           =   10
      Left            =   2280
      TabIndex        =   27
      Top             =   6000
      Width           =   720
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   375
      Index           =   9
      Left            =   2280
      TabIndex        =   24
      Top             =   4920
      Width           =   240
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   375
      Index           =   8
      Left            =   2280
      TabIndex        =   21
      Top             =   4440
      Width           =   240
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   375
      Index           =   5
      Left            =   2280
      TabIndex        =   11
      Top             =   2640
      Width           =   240
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000A&
      Height          =   375
      Index           =   7
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3960
      Width           =   240
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   375
      Index           =   6
      Left            =   2280
      TabIndex        =   13
      Top             =   3480
      Width           =   240
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000A&
      Height          =   375
      Index           =   4
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000A&
      Height          =   375
      Index           =   3
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1440
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   7
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���@��"
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
      Left            =   1800
      TabIndex        =   1
      ToolTipText     =   "���i���\����ǂݍ��݂܂��i�e5�j"
      Top             =   0
      Width           =   1380
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   2535
      Left            =   240
      TabIndex        =   5
      Top             =   7200
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   4471
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "�Ǘ���"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "���@�t"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "�݌v�ύX��"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "���޽�i��"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "�H��i��"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "��"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "���޽�i��"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "�H��i��"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "�ݕώ��{"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "�ύX���i"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "�ύX���e(�ύX/�ǉ�)"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "�����ꏊ"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "�ݕό����ۊ�"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "��       �l�P"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "��       �l�Q"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "��       �l�R"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "��       �l�S"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   17
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   714
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=17"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=794"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=688"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2011"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1905"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1826"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1720"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2381"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2275"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=1588"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1482"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=609"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=503"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=2831"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2725"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=1958"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=1852"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=847"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=741"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=1958"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=1852"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(41)=   "Column(10).Width=3334"
      Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=3228"
      Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(45)=   "Column(11).Width=2593"
      Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=2487"
      Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(49)=   "Column(12).Width=3281"
      Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=3175"
      Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(53)=   "Column(13).Width=12250"
      Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=12144"
      Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(57)=   "Column(14).Width=3281"
      Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=3175"
      Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(61)=   "Column(15).Width=3281"
      Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=3175"
      Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(65)=   "Column(16).Width=3281"
      Splits(0)._ColumnProps(66)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(16)._WidthInPix=3175"
      Splits(0)._ColumnProps(68)=   "Column(16).Order=17"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �S�V�b�N"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �S�V�b�N"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
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
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
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
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFF00&"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=110,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=107,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=108,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=109,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=3"
      _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=86,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=83,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=84,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=85,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=66,.parent=13,.alignment=3"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=70,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=67,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=68,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=69,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=74,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=71,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=72,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=73,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(12).Style:id=78,.parent=13"
      _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=75,.parent=14"
      _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=76,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=77,.parent=17"
      _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=82,.parent=13"
      _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=79,.parent=14"
      _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=80,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=81,.parent=17"
      _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=90,.parent=13"
      _StyleDefs(95)  =   "Splits(0).Columns(14).HeadingStyle:id=87,.parent=14"
      _StyleDefs(96)  =   "Splits(0).Columns(14).FooterStyle:id=88,.parent=15"
      _StyleDefs(97)  =   "Splits(0).Columns(14).EditorStyle:id=89,.parent=17"
      _StyleDefs(98)  =   "Splits(0).Columns(15).Style:id=94,.parent=13"
      _StyleDefs(99)  =   "Splits(0).Columns(15).HeadingStyle:id=91,.parent=14"
      _StyleDefs(100) =   "Splits(0).Columns(15).FooterStyle:id=92,.parent=15"
      _StyleDefs(101) =   "Splits(0).Columns(15).EditorStyle:id=93,.parent=17"
      _StyleDefs(102) =   "Splits(0).Columns(16).Style:id=98,.parent=13"
      _StyleDefs(103) =   "Splits(0).Columns(16).HeadingStyle:id=95,.parent=14"
      _StyleDefs(104) =   "Splits(0).Columns(16).FooterStyle:id=96,.parent=15"
      _StyleDefs(105) =   "Splits(0).Columns(16).EditorStyle:id=97,.parent=17"
      _StyleDefs(106) =   "Named:id=33:Normal"
      _StyleDefs(107) =   ":id=33,.parent=0"
      _StyleDefs(108) =   "Named:id=34:Heading"
      _StyleDefs(109) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(110) =   ":id=34,.wraptext=-1"
      _StyleDefs(111) =   "Named:id=35:Footing"
      _StyleDefs(112) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(113) =   "Named:id=36:Selected"
      _StyleDefs(114) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(115) =   "Named:id=37:Caption"
      _StyleDefs(116) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(117) =   "Named:id=38:HighlightRow"
      _StyleDefs(118) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(119) =   "Named:id=39:EvenRow"
      _StyleDefs(120) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(121) =   "Named:id=40:OddRow"
      _StyleDefs(122) =   ":id=40,.parent=33"
      _StyleDefs(123) =   "Named:id=41:RecordSelector"
      _StyleDefs(124) =   ":id=41,.parent=34"
      _StyleDefs(125) =   "Named:id=42:FilterBar"
      _StyleDefs(126) =   ":id=42,.parent=33"
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
      Index           =   2
      Left            =   3360
      TabIndex        =   2
      ToolTipText     =   "���i���\����ۑ����܂�"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�o �^"
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
      Left            =   420
      TabIndex        =   0
      Top             =   0
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      Caption         =   "��"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   19
      Left            =   6240
      TabIndex        =   38
      Top             =   6600
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      Caption         =   "���݌�"
      Height          =   240
      Index           =   13
      Left            =   1200
      TabIndex        =   36
      Top             =   6600
      Width           =   960
   End
   Begin VB.Label LblHantei_MARK 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   3240
      TabIndex        =   35
      Top             =   5640
      Width           =   120
   End
   Begin VB.Label Label1 
      Caption         =   "�S����"
      Height          =   255
      Index           =   18
      Left            =   360
      TabIndex        =   34
      Top             =   840
      Width           =   855
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      TabIndex        =   33
      Top             =   9840
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      Caption         =   "��"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   17
      Left            =   6240
      TabIndex        =   32
      Top             =   6120
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      Caption         =   "������"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   16
      Left            =   4200
      TabIndex        =   30
      Top             =   6120
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      Caption         =   "�~"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   15
      Left            =   3000
      TabIndex        =   28
      Top             =   6120
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      Caption         =   "����݌�"
      Height          =   240
      Index           =   14
      Left            =   1200
      TabIndex        =   26
      Top             =   6120
      Width           =   960
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      Caption         =   "0:��Ώ�/1:�Ώ�"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   11
      Left            =   2760
      TabIndex        =   25
      Top             =   5040
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      Caption         =   "�݌v�ύX�Ώ�"
      Height          =   240
      Index           =   10
      Left            =   720
      TabIndex        =   23
      Top             =   5040
      Width           =   1440
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      Caption         =   "1:����Ȃ�"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   9
      Left            =   2760
      TabIndex        =   22
      Top             =   4560
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   240
      Index           =   8
      Left            =   480
      TabIndex        =   20
      Top             =   4560
      Width           =   1680
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      Caption         =   "0:�P�i/1�ƯĐe/2:�ƯĎq/3:�P�i�Ư�"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   7
      Left            =   2760
      TabIndex        =   19
      Top             =   2760
      Width           =   4080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ӽޭ���Ưċ敪"
      Height          =   240
      Index           =   6
      Left            =   480
      TabIndex        =   18
      Top             =   2760
      Width           =   1680
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      Caption         =   "0:��Ώ�/1:�Ώ�/2:�Ő؈ē���/3:�Ő�"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   5
      Left            =   2760
      TabIndex        =   17
      Top             =   4080
      Width           =   4200
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���������敪"
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   15
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      Caption         =   "0:��Ώ�/1:�Ώ�"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   2760
      TabIndex        =   14
      Top             =   3600
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "Ӽޭ�ّΏ�"
      Height          =   255
      Index           =   2
      Left            =   585
      TabIndex        =   12
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "��\�@��(����)"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   9
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "�i��(�O��)"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "PCB00201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim DEF_NAIGAI          As String * 1



Private Const ptxTANTO_CODE% = 0        '�S���҃R�[�h
Private Const ptxTANTO_NAME% = 1        '�S���Җ�


Private Const ptxHIN_GAI% = 2           '�i��(�O��)
Private Const ptxHIN_NAME% = 3          '�i��
Private Const ptxL_KISHU1% = 4          '��\�@��(����)

Private Const ptxMODULE_UNIT_KBN% = 5   'Ӽޭ���Ưċ敪


Private Const ptxMODULE_KBN% = 6        'Ӽޭ�ّΏ�
Private Const ptxNAI_BUHIN% = 7         '���������敪

Private Const ptxKENSA_JIGU% = 8        '��������
Private Const ptxSETUHEN_KBN% = 9       '�݌v�ύX�Ώ�


Private Const ptxHITUYO_SU% = 10        '�K�v���@��
Private Const ptxHITUYO_TUKI% = 11      '�K�v���@��

Private Const ptxHITUYO_QTY% = 12       '�K�v���@�~��
Private Const ptxZAIKO_NOW% = 13        '���݌�


'--------------------------------------------<��د��>
Dim PCB_U      As New XArrayDB

Private Const Min_Row% = 1              '�ŏ��s��

Dim Max_Row    As Integer               '�O���b�h�ő�\������


Private Const Min_Col% = 0              '�ŏ���
Private Const Max_Col% = 16             '�ő��

Private Const colKANRI_NO% = 0          '�Ǘ���
Private Const colEX_DATE% = 1           '���t
Private Const colSETUHEN_NO% = 2        '�ݕϊǗ���
Private Const colBEF_HIN_GAI% = 3       '�ύX�O�@���޽�i��
Private Const colBEF_HIN_NAI% = 4       '�ύX�O�@�H��i��

Private Const colDummy% = 5             '��

Private Const colAFT_HIN_GAI% = 6       '�ύX��@���޽�i��
Private Const colAFT_HIN_NAI% = 7       '�ύX��@�H��i��

Private Const colSETUHEN_JITSU% = 8     '�ݕώ��{


Private Const colHEN_BUHIN% = 9         '�ύX���i
Private Const colHEN_NAIYO% = 10         '�ύX���e
Private Const colHEN_BASHO% = 11        '�����ꏊ
Private Const colSETUHEN_HOKAN% = 12    '�ݕό����ۊ�
Private Const colBIKOU1% = 13           '���l1
Private Const colBIKOU2% = 14           '���l2
Private Const colBIKOU3% = 15           '���l3
Private Const colBIKOU4% = 16           '���l4
'--------------------------------------------<EXCEL>
Private Const selKANRI_NO% = 2          '�Ǘ���
Private Const selEX_DATE% = 3           '���t
Private Const selSETUHEN_NO% = 4        '�ݕϊǗ���
Private Const selBEF_HIN_GAI% = 5       '�ύX�O�@���޽�i��
Private Const selBEF_HIN_NAI% = 6       '�ύX�O�@�H��i��

Private Const selDummy% = 7             '��

Private Const selAFT_HIN_GAI% = 8       '�ύX��@���޽�i��
Private Const selAFT_HIN_NAI% = 9       '�ύX��@�H��i��

Private Const selSETUHEN_JITSU% = 10    '�ݕώ��{


Private Const selHEN_BUHIN% = 11        '�ύX���i
Private Const selHEN_NAIYO% = 12        '�ύX���e
Private Const selHEN_BASHO% = 15        '�����ꏊ
Private Const selSETUHEN_HOKAN% = 16    '�ݕό����ۊ�
Private Const selBIKOU1% = 17           '���l1
Private Const selBIKOU2% = 18           '���l2
Private Const selBIKOU3% = 19           '���l3
Private Const selBIKOU4% = 20           '���l4

    
        


'------------------------------------------------------------------ �ޗǃ��W���[���p    2014.07.03
Private Nara_Soko_T              As Variant




'Private Const LAST_UPDATE_DAY$ = "[PCB0020] <���x����e�X�g�p>�@2015.01.21 08:00"
Private Const LAST_UPDATE_DAY$ = "[PCB0020] <���x����e�X�g�p>�@2015.01.21 13:00"

Private Sub Command1_Click(Index As Integer)

Dim sWk         As String
Dim i           As Integer

Dim yn          As Integer

Dim Cl_Now      As String * 8   '2015.01.20
Dim Zaiko_Qty   As Long         '2015.01.21
Dim Location    As String * 8   '2015.01.21

Dim Sumi_Qty    As Long         '2015.01.21
Dim Mi_Qty      As Long         '2015.01.21

    Select Case Index



        Case 0          '�o�^
            
            Cl_Now = Format(Now, "hh:mm:ss")        '2015.01.20
            
            
            
'            If Trim(LblHantei_MARK.Caption) = "Ӽޭ�ٕi�ږ��o�^" Then      2014.09.17
'                MsgBox "���̉�ʂŁAӼޭ�ٕi�ڂ̒ǉ��͏o���܂���B"        2014.09.17
'                Exit Sub                                                   2014.09.17
'            End If                                                         2014.09.17
            
            
            
            
            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2015.01.20
    
'            If RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) = RTrim(Text1(ptxHIN_GAI).Text) Then
                
                Zaiko_Qty = 0
                For i = 0 To UBound(Nara_Soko_T)
                
                    Location = Nara_Soko_T(i)
                
                    If SOKO_Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode), Location) Then
                        Exit Sub
                    End If
                    Zaiko_Qty = Zaiko_Qty + (Sumi_Qty + Mi_Qty)
                Next i
                Text1(ptxZAIKO_NOW).Text = Format(Zaiko_Qty, "#")
            
'            End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2015.01.20
            
            
            
            
            
            For i = ptxTANTO_CODE To ptxHITUYO_QTY
            
                If Error_Check_Proc(i, Zaiko_Qty) Then      '�����ǉ� 2015.01.21
                    Exit Sub
                End If
            
            Next i
            
            
            yn = MsgBox("�o�^���܂����H", vbYesNo, "�m�F����")

            If yn = vbYes Then
                If Update_Proc(Cl_Now) Then
                    Unload Me
                End If
            
            
                Call Clear_Field_Proc
            
            End If



        Case 1          '����
                
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2015.01.20
    
'            If RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) = RTrim(Text1(ptxHIN_GAI).Text) Then
                
                Zaiko_Qty = 0
                For i = 0 To UBound(Nara_Soko_T)
                
                    Location = Nara_Soko_T(i)
                
                    If SOKO_Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode), Location) Then
                        Exit Sub
                    End If
                    Zaiko_Qty = Zaiko_Qty + (Sumi_Qty + Mi_Qty)
                Next i
                Text1(ptxZAIKO_NOW).Text = Format(Zaiko_Qty, "#")
            
'            End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2015.01.20
                
                
            For i = ptxTANTO_CODE To ptxHIN_GAI
                If Error_Check_Proc(i, Zaiko_Qty) Then  '�����ǉ� 2015.01.21
                    Exit Sub
                End If
            Next i

        Case 2          '�I��

            yn = MsgBox("�I�����܂����H", vbYesNo, "�m�F����")
            If yn = vbYes Then
                Unload Me
            End If
    End Select



    Command1(Index).SetFocus


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �j���� �c������ �O����
'----------------------------------------------------------------------------
    
    
'    If Shift = vbAltMask Then
'
'        If TDBGrid1.AllowUpdate Then
'
'            TDBGrid1.AllowUpdate = False
'            TDBGrid1.AllowAddNew = False
'            TDBGrid1.AllowDelete = False
'
'
'            TDBGrid1.Columns(colTEI_LABELID).Visible = False
'            TDBGrid1.Columns(colHAKO_NO).Visible = False
'
'            TDBGrid1.Columns(colTEI_LABELID).Locked = True
'            TDBGrid1.Columns(colHAKO_NO).Locked = True
'
'
'
'        Else
'
'
'            TDBGrid1.AllowUpdate = True
'            TDBGrid1.AllowAddNew = True
'            TDBGrid1.AllowDelete = True
'
'
'            TDBGrid1.Columns(colTEI_LABELID).Visible = True
'            TDBGrid1.Columns(colHAKO_NO).Visible = True
'
'
'            TDBGrid1.Columns(colTEI_LABELID).Locked = False
'            TDBGrid1.Columns(colHAKO_NO).Locked = False
'
'        End If
'
'    End If
    
    
    Select Case KeyCode
        Case vbKeyF12
            Unload Me
    End Select
    
    
    

End Sub

Private Sub Form_Load()


Dim c       As String * 128
Dim i       As Integer


    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�i�ڃ}�X�^�����e�i���X(���W���[�����i�p)", Me.hwnd, 0)
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



    If GetIni(App.EXEName, "NAIGAI", App.EXEName, c) Then
        Beep
        MsgBox "�����O[NAIGAI]�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    DEF_NAIGAI = Trim(c)


'---------------------------------------------- '�ޗǃ��W���[���@�Ώۑq�� 2014.07.03
    If GetIni(App.EXEName, "MODULE_SOKO", App.EXEName, c) Then
        c = "**"
        Nara_Soko_T = Split(Trim(c), ",", -1)
    Else
        Nara_Soko_T = Split(Trim(c), ",", -1)
    End If


'---------------------------------------------- '�ޗǃ��W���[���@�Ώۑq�� 2014.07.03




                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            PCB00201.Caption = "�i�ڃ}�X�^�����e�i���X(���W���[�����i�p)(" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)

'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i

    Unload SubMenu(i)




    PCB00201.Caption = PCB00201.Caption & " " & LAST_UPDATE_DAY

                                '���W���[���i�ڃ}�X�^ �n�o�d�m
    If M_ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^ �n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�S���҃}�X�^ �n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�݌Ƀf�[�^ �n�o�d�m    2014.07.03
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                'PCB.U�ݕ� �n�o�d�m
    If PCB_U_Open(BtOpenNomal) Then
        Unload Me
    End If


    Call Clear_Field_Proc


    Text1(ptxTANTO_CODE).SetFocus

End Sub



Private Sub Form_Unload(Cancel As Integer)
    
    
Dim sts     As Integer
    
    
    sts = BTRV(BtOpClose, M_ITEM_POS, M_ITEM_REC, Len(M_ITEM_REC), K0_M_ITEM, Len(K0_M_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���W���[���i�ڃ}�X�^")
        End If
    End If
    
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
    
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�S���҃}�X�^")
        End If
    End If
    
    sts = BTRV(BtOpClose, PCB_U_POS, PCB_U_REC, Len(PCB_U_REC), K0_PCB_U, Len(K0_PCB_U), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "PCB.U�ݕϊǗ��䒠")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, M_ITEM_POS, M_ITEM_REC, Len(M_ITEM_REC), K0_M_ITEM, Len(K0_M_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If


    Set PCB00201 = Nothing



    End

End Sub


Private Sub SubMenu_Click(Index As Integer)
Dim i   As Integer
    
    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    PCB00201.Caption = "�i�ڃ}�X�^�����e�i���X(���W���[�����i�p)�i" + RTrim(JGYOBU_T(Index).NAME) + "�j"
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub




Private Function Update_Proc(Cl_Now As String) As Integer
'----------------------------------------------------------------------------
'                   �uPCB.U�ݕρv�o�^����
'----------------------------------------------------------------------------

Dim sts             As Integer
Dim ans             As Integer
    
Dim bt_update       As Integer  '2014.09.17


Dim St_Now          As String * 8   '2015.01.20


    Update_Proc = True
    
    Call Input_Lock


St_Now = Format(Now, "hh:mm:ss")        '2015.01.20

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�i�ڃ}�X�^�����e�i���X(���W���[�����i�p)�@[�o�^]�����J�n�I�I", Me.hwnd, 0)


        
'-------------------------------------  <���W���[���i�ڃ}�X�^����>
    Call UniCode_Conv(K0_M_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_M_ITEM.NAIGAI, DEF_NAIGAI)
    Call UniCode_Conv(K0_M_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
    
    sts = BTRV(BtOpGetEqual, M_ITEM_POS, M_ITEM_REC, Len(M_ITEM_REC), K0_M_ITEM, Len(K0_M_ITEM), 0)
    Select Case sts
        Case BtNoErr
        
            bt_update = BtOpUpdate      '2014.09.17
        
        Case BtErrKeyNotFound
        
        
        
 '           MsgBox "�i�ڃ}�X�^�i���W���[���j�̓��e���ύX����Ă��܂��B�i�ԁi�O���j�̓��e���m�F���ĉ������B"    2014.09.17
 '
 '
 '           Update_Proc = False                                                                                2014.09.17
 '           Exit Function                                                                                      2014.09.17
        
            bt_update = BtOpInsert      '2014.09.17
           
        
        
        
        Case Else
            Call File_Error(sts, BtOpInsert, "���W���[���i�ڃ}�X�^")
            Call Input_UnLock
            Exit Function
    End Select
            
    If bt_update = BtOpInsert Then          '2014.09.17
        Call UniCode_Conv(M_ITEM_REC.JGYOBU, Last_JGYOBU)                       '���ƕ�
        Call UniCode_Conv(M_ITEM_REC.NAIGAI, DEF_NAIGAI)                        '�����O
        Call UniCode_Conv(M_ITEM_REC.HIN_GAI, Text1(ptxHIN_GAI).Text)           '�i�ԁi�O���j

        Call UniCode_Conv(M_ITEM_REC.MODULE_KBN, "")                            '���W���[���Ώۋ敪
        Call UniCode_Conv(M_ITEM_REC.MODULE_UNIT_KBN, "")                       '���W���[�����j�b�g�敪
        
        Call UniCode_Conv(M_ITEM_REC.KENSA_JIGU, "")                            '��������
        
        Call UniCode_Conv(M_ITEM_REC.SETUHEN_KBN, "")                           '�݌v�ύX�Ώۋ敪
        
        Call UniCode_Conv(M_ITEM_REC.SENDO_LAST_DATE, "")                       '�N�x�Ǘ��ŏI��
        
        Call UniCode_Conv(M_ITEM_REC.HITUYO_SU, "")                             '�K�v���@��
        Call UniCode_Conv(M_ITEM_REC.HITUYO_TUKI, "")                           '�K�v���@��

    
    
        Call UniCode_Conv(M_ITEM_REC.FILLER, "")
    
        Call UniCode_Conv(M_ITEM_REC.INS_TANTO, App.EXEName)                    '�ǉ��S����
        Call UniCode_Conv(M_ITEM_REC.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))                     '�ǉ�����
        Call UniCode_Conv(M_ITEM_REC.INS_PROG_ID, App.EXEName)                  '�ǉ��v���O����ID
    End If
            

    Call UniCode_Conv(M_ITEM_REC.MODULE_KBN, Text1(ptxMODULE_KBN).Text)             '���W���[���Ώۋ敪
    Call UniCode_Conv(M_ITEM_REC.MODULE_UNIT_KBN, Text1(ptxMODULE_UNIT_KBN).Text)   '���W���[�����j�b�g�敪
            
    Call UniCode_Conv(M_ITEM_REC.KENSA_JIGU, Text1(ptxKENSA_JIGU).Text)             '��������
            
    Call UniCode_Conv(M_ITEM_REC.SETUHEN_KBN, Text1(ptxSETUHEN_KBN).Text)           '�݌v�ύX�Ώۋ敪
            
    
'    Call UniCode_Conv(M_ITEM_REC.SETUHEN_LAST_DATE, Format(Text1(ptxSETUHEN_LAST_DATE).Text, "YYYYMMDD"))       '�݌v�ύX�ŏI��
'
'    If IsDate(Text1(ptxSENDO_LAST_DATE).Text) Then
'        Call UniCode_Conv(M_ITEM_REC.SENDO_LAST_DATE, Format(Text1(ptxSENDO_LAST_DATE).Text, "YYYYMMDD"))       '�N�x�Ǘ��ŏI��
'    Else
'        Call UniCode_Conv(M_ITEM_REC.SENDO_LAST_DATE, "")
'    End If
            
            
    Call UniCode_Conv(M_ITEM_REC.HITUYO_SU, Format(Val(Text1(ptxHITUYO_SU).Text), "00000"))                     '�K�v���@��
    Call UniCode_Conv(M_ITEM_REC.HITUYO_TUKI, Format(Val(Text1(ptxHITUYO_TUKI).Text), "00"))                    '�K�v���@��

        
        
        
    If bt_update = BtOpUpdate Then      '2014.09.17
        Call UniCode_Conv(M_ITEM_REC.UPD_TANTO, Text1(ptxTANTO_CODE).Text)              '�X�V�S����
        Call UniCode_Conv(M_ITEM_REC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))       '�X�V����
        Call UniCode_Conv(M_ITEM_REC.UPD_PROG_ID, App.EXEName)                          '�X�V�v���O����
    End If
            
            
                    
    Do
'        sts = BTRV(BtOpUpdate, M_ITEM_POS, M_ITEM_REC, Len(M_ITEM_REC), K0_M_ITEM, Len(K0_M_ITEM), 0)      '2014.09.17
        
        sts = BTRV(bt_update, M_ITEM_POS, M_ITEM_REC, Len(M_ITEM_REC), K0_M_ITEM, Len(K0_M_ITEM), 0)        '2014.09.17
        
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                
                Beep
                ans = MsgBox("�u�i�ڃ}�X�^(���W���[��)�v���[���Ńf�[�^�g�p���ł��B<M_ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Call Input_UnLock
                    Exit Function
                End If
            
            Case Else
                Call Input_UnLock
                Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^(���W���[��)")
                Exit Function
        End Select
        
    Loop


    Call UniCode_Conv(ITEMREC.HIN_GAI, "")


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�i�ڃ}�X�^�����e�i���X(���W���[�����i�p)�@[�o�^]�����I���I�I" & " Cl_TIME=" & Cl_Now & " ST_TIME=" & St_Now & " END_TIME=" & Format(Now, "hh:mm:ss"), Me.hwnd, 0)



    Call Input_UnLock


    Update_Proc = False

End Function

Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   �uOCB.U �ݕϊǗ��䒠�v�Ǎ��ݏ���
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
    

Dim Row             As Long
Dim i               As Long



Dim com             As Integer

Dim wkDATE          As String * 10



    List_Disp_Proc = True

    Call Input_Lock

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�i�ڃ}�X�^�����e�i���X(���W���[�����i�p)�@[����]�����J�n�I�I", Me.hwnd, 0)


    Set PCB_U = Nothing
    
    Row = Min_Row - 1




    Call UniCode_Conv(K0_PCB_U.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_PCB_U.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_PCB_U.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))

    Call UniCode_Conv(K0_PCB_U.EX_DATE, "")

    com = BtOpGetGreaterEqual

    Do
        DoEvents
    
        sts = BTRV(com, PCB_U_POS, PCB_U_REC, Len(PCB_U_REC), K0_PCB_U, Len(K0_PCB_U), 0)
        Select Case sts
            Case BtNoErr
            
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> StrConv(PCB_U_REC.JGYOBU, vbUnicode) Or _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) <> StrConv(PCB_U_REC.NAIGAI, vbUnicode) Or _
                    StrConv(ITEMREC.HIN_GAI, vbUnicode) <> StrConv(PCB_U_REC.HIN_GAI, vbUnicode) Then
                    Exit Do
                End If

'            Case BtErrKeyNotFound  '2014.09.17
            Case BtErrEOF           '2014.09.17
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "PCB.U�ݕ�")
                Call Input_UnLock
                Exit Function
        End Select
    
        Row = Row + 1
        PCB_U.ReDim Min_Row, Row, Min_Col, Max_Col
            
            
        PCB_U(Row, colKANRI_NO) = RTrim(StrConv(PCB_U_REC.KANRI_NO, vbUnicode))
        
        
        wkDATE = RTrim(StrConv(PCB_U_REC.EX_DATE, vbUnicode))
        If Trim(wkDATE) <> "" Then
            wkDATE = Mid(wkDATE, 1, 4) & "/" & Mid(wkDATE, 5, 2) & "/" & Mid(wkDATE, 7, 2)
        End If
        PCB_U(Row, colEX_DATE) = wkDATE
        
        
        PCB_U(Row, colSETUHEN_NO) = RTrim(StrConv(PCB_U_REC.SETUHEN_NO, vbUnicode))
        PCB_U(Row, colBEF_HIN_GAI) = RTrim(StrConv(PCB_U_REC.BEF_HIN_GAI, vbUnicode))
        PCB_U(Row, colBEF_HIN_NAI) = RTrim(StrConv(PCB_U_REC.BEF_HIN_NAI, vbUnicode))
        PCB_U(Row, colAFT_HIN_GAI) = RTrim(StrConv(PCB_U_REC.AFT_HIN_GAI, vbUnicode))
        PCB_U(Row, colAFT_HIN_NAI) = RTrim(StrConv(PCB_U_REC.AFT_HIN_NAI, vbUnicode))
        
        PCB_U(Row, colSETUHEN_JITSU) = RTrim(StrConv(PCB_U_REC.SETUHEN_JITSU, vbUnicode))
        
        
        PCB_U(Row, colHEN_BUHIN) = RTrim(StrConv(PCB_U_REC.HEN_BUHIN, vbUnicode))
        PCB_U(Row, colHEN_NAIYO) = RTrim(StrConv(PCB_U_REC.HEN_NAIYO, vbUnicode))
        PCB_U(Row, colHEN_BASHO) = RTrim(StrConv(PCB_U_REC.HEN_BASHO, vbUnicode))
        PCB_U(Row, colSETUHEN_HOKAN) = RTrim(StrConv(PCB_U_REC.SETUHEN_HOKAN, vbUnicode))
        PCB_U(Row, colBIKOU1) = RTrim(StrConv(PCB_U_REC.BIKOU1, vbUnicode))
        PCB_U(Row, colBIKOU2) = RTrim(StrConv(PCB_U_REC.BIKOU2, vbUnicode))
        PCB_U(Row, colBIKOU3) = RTrim(StrConv(PCB_U_REC.BIKOU3, vbUnicode))
        PCB_U(Row, colBIKOU4) = RTrim(StrConv(PCB_U_REC.BIKOU4, vbUnicode))
    
    
        com = BtOpGetNext
    
    
    Loop


    Set TDBGrid1.Array = PCB_U
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�i�ڃ}�X�^�����e�i���X(���W���[�����i�p)�@[����]�����I���I�I", Me.hwnd, 0)



    Call Input_UnLock


    List_Disp_Proc = False
    Exit Function


End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------
Dim i   As Integer


    PCB00201.MousePointer = vbHourglass

    Call Ctrl_Lock(PCB00201)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(PCB00201)


    PCB00201.MousePointer = vbDefault

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
Dim i   As Integer
    
Dim Zaiko_Qty   As Long         '2015.01.21
Dim Sumi_Qty    As Long         '2015.01.21
Dim Mi_Qty      As Long         '2015.01.21
Dim Location    As String * 8   '2015.01.21
    
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If


'>>>>>>>>>>>>>  2015.01.21

'    If Trim(Text1(ptxZAIKO_NOW).Text) = "" Then

        Zaiko_Qty = 0
        For i = 0 To UBound(Nara_Soko_T)
        
            Location = Nara_Soko_T(i)
        
            If SOKO_Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode), Location) Then
                Exit Sub
            End If
            Zaiko_Qty = Zaiko_Qty + (Sumi_Qty + Mi_Qty)
        Next i
        Text1(ptxZAIKO_NOW).Text = Format(Zaiko_Qty, "#")

'    End If
'>>>>>>>>>>>>>  2015.01.21


    If Error_Check_Proc(Index, Zaiko_Qty) Then  '�G���[�`�F�b�N         '2015.01.21 �����ǉ�
        Exit Sub
    End If
        
        
'    Call Tab_Ctrl(Shift)        '�ړ�



    For i = Index + 1 To ptxHITUYO_QTY
    
        If Text1(i).TabStop Then
            
            Text1(i).SetFocus
            Exit Sub
        
        End If
    
    Next i


End Sub
Private Function Error_Check_Proc(Mode As Integer, Optional Zaiko_Qty As Long = 0) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'   2015.01.21 �����ǉ��@Zaiko_qty
'----------------------------------------------------------------------------
    
Dim sts         As Integer
    
Dim wkDATE      As String * 8
    
    
Dim USE_QTY     As Long     '2014.07.03
'Dim Zaiko_qty   As Long     '2014.07.03        -->2015.01.21 DEL
Dim Sumi_Qty    As Long     '2014.07.03
Dim Mi_Qty      As Long     '2014.07.03
    
Dim i           As Integer      '2014.07.03
Dim Location    As String * 8   '2014.07.03
Dim HANTEI_MARK As String   '2014.07.03
    
    Error_Check_Proc = True
    
    
    
    
    
    
    
    
    Select Case Mode
    
    
        Case ptxTANTO_CODE     '�S����
        
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTANTO_CODE).Text)

            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    Text1(ptxTANTO_NAME).Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    Text1(ptxTANTO_NAME).Text = ""
            
                    MsgBox "���͂������ڂ̓G���[�ł��B(�S����)"
                    Text1(Mode).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
                    Exit Function
                
            
            
            End Select
        
        Case ptxHIN_GAI         '�i��
    
            
            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2015.01.20
            If RTrim(Text1(ptxHIN_GAI).Text) = "" Then
            
                Text1(ptxHIN_NAME).Text = ""
                Text1(ptxL_KISHU1).Text = ""
                
                Text1(ptxNAI_BUHIN).Text = ""

                MsgBox "���͂������ڂ̓G���[�ł��B(�i�ڃ}�X�^ ���o�^)"
                Text1(Mode).SetFocus
                Exit Function
            End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2015.01.20
            
            
            Text1(ptxHIN_GAI).Text = Trim(StrConv(Text1(ptxHIN_GAI).Text, vbUpperCase))
            
            
            If RTrim(Text1(ptxHIN_GAI).Text) = RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                Error_Check_Proc = False
                Exit Function
            End If
            
            
            Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, DEF_NAIGAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)


            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                
                    Text1(ptxHIN_NAME).Text = RTrim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                    Text1(ptxL_KISHU1).Text = RTrim(StrConv(ITEMREC.L_KISHU1, vbUnicode))
                    Text1(ptxNAI_BUHIN).Text = RTrim(StrConv(ITEMREC.NAI_BUHIN, vbUnicode))
                
                                
                    Call UniCode_Conv(K0_M_ITEM.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_M_ITEM.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_M_ITEM.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                
                    sts = BTRV(BtOpGetEqual, M_ITEM_POS, M_ITEM_REC, Len(M_ITEM_REC), K0_M_ITEM, Len(K0_M_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                            
                            
                            
                                            
                            Text1(ptxMODULE_KBN).Text = StrConv(M_ITEM_REC.MODULE_KBN, vbUnicode)
                            Text1(ptxMODULE_UNIT_KBN).Text = StrConv(M_ITEM_REC.MODULE_UNIT_KBN, vbUnicode)
                            Text1(ptxKENSA_JIGU).Text = StrConv(M_ITEM_REC.KENSA_JIGU, vbUnicode)
                            Text1(ptxSETUHEN_KBN).Text = StrConv(M_ITEM_REC.SETUHEN_KBN, vbUnicode)
                        
                        
                            
                            'If Trim(StrConv(M_ITEM_REC.SETUHEN_LAST_DATE, vbUnicode)) = "" Then
                            'Else
                            '    wkDATE = StrConv(M_ITEM_REC.SETUHEN_LAST_DATE, vbUnicode)
                            '    Text1(ptxSETUHEN_LAST_DATE).Text = Mid(wkDATE, 1, 4) & "/" & Mid(wkDATE, 5, 2) & "/" & Mid(wkDATE, 7, 2)
                            'End If
                            
                            'If Trim(StrConv(M_ITEM_REC.SENDO_LAST_DATE, vbUnicode)) = "" Then
                            'Else
                            '    wkDATE = StrConv(M_ITEM_REC.SENDO_LAST_DATE, vbUnicode)
                            '    Text1(ptxSENDO_LAST_DATE).Text = Mid(wkDATE, 1, 4) & "/" & Mid(wkDATE, 5, 2) & "/" & Mid(wkDATE, 7, 2)
                            'End If
                            
                            
                            'lblTUKI.Caption = Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode)), "#") 2014.07.17
                            
                            
                            Text1(ptxHITUYO_SU).Text = Format(Val(StrConv(M_ITEM_REC.HITUYO_SU, vbUnicode)), "#")
                            Text1(ptxHITUYO_TUKI).Text = Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode)), "#")
                            Text1(ptxHITUYO_QTY).Text = Format(Val(Text1(ptxHITUYO_SU).Text) * Val(Text1(ptxHITUYO_TUKI).Text), "#")
                                                           
                            USE_QTY = Val(Text1(ptxHITUYO_QTY).Text)
                                                           
                            Zaiko_Qty = 0
                            For i = 0 To UBound(Nara_Soko_T)
                            
                                Location = Nara_Soko_T(i)
                            
                                If SOKO_Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode), Location) Then
                                    Exit Function
                                End If
                                Zaiko_Qty = Zaiko_Qty + (Sumi_Qty + Mi_Qty)
                            Next i
                            Text1(ptxZAIKO_NOW).Text = Format(Zaiko_Qty, "#")
                                                           
                            Call HANTEI_Proc(HANTEI_MARK, Zaiko_Qty)    '2015.01.20 �����ǉ�
                                                           
                                                           
                                                           
                            LblHantei_MARK.Caption = HANTEI_MARK

                            If List_Disp_Proc() Then
                                Exit Function
                            End If
                        
                        
                        Case BtErrKeyNotFound
                        
                        
                        
                            Text1(ptxMODULE_KBN).Text = ""
                            Text1(ptxMODULE_UNIT_KBN).Text = ""
                            Text1(ptxKENSA_JIGU).Text = ""
                            Text1(ptxSETUHEN_KBN).Text = ""
                        
'                            Text1(ptxSENDO_LAST_DATE).Text = ""
'                            Text1(ptxHITUYO_SU).Text = ""
                        
                            'lblTUKI.Caption = ""   2014.07.17
                            
                            
                            Text1(ptxHITUYO_SU).Text = ""
                            Text1(ptxHITUYO_TUKI).Text = ""
                            Text1(ptxHITUYO_QTY).Text = ""
                        
                        
                            Text1(ptxZAIKO_NOW).Text = ""   '2014.11.18
                        
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2015.01.21
                            USE_QTY = Val(Text1(ptxHITUYO_QTY).Text)
                                                           
                            Zaiko_Qty = 0
                            For i = 0 To UBound(Nara_Soko_T)
                            
                                Location = Nara_Soko_T(i)
                            
                                If SOKO_Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode), Location) Then
                                    Exit Function
                                End If
                                Zaiko_Qty = Zaiko_Qty + (Sumi_Qty + Mi_Qty)
                            Next i
                            Text1(ptxZAIKO_NOW).Text = Format(Zaiko_Qty, "#")
                                                           
                            Call HANTEI_Proc(HANTEI_MARK, Zaiko_Qty)    '2015.01.20 �����ǉ�
                                                           
                                                           
                                                           
                            LblHantei_MARK.Caption = HANTEI_MARK
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2015.01.21
                        
                        
                        
                            If List_Disp_Proc() Then        '2014.09.17
                                Exit Function
                            End If
                        
                        
'                            MsgBox "���͂������ڂ̓G���[�ł��B(�i�ڃ}�X�^�i���W���[���j ���o�^)"   2014.09.17
'                            Text1(Mode).SetFocus                                                   2014.09.17
                            
                            LblHantei_MARK.Caption = "Ӽޭ�ٕi�ږ��o�^"        '2014.09.17
                        
                                        
                            'Exit Function                                      2014.09.17
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^�i���W���[���j")
                            Exit Function
                    
                    End Select
                
                Case BtErrKeyNotFound

                    Text1(ptxHIN_NAME).Text = ""
                    Text1(ptxL_KISHU1).Text = ""
                    
                    Text1(ptxNAI_BUHIN).Text = ""

                    MsgBox "���͂������ڂ̓G���[�ł��B(�i�ڃ}�X�^ ���o�^)"
                    Text1(Mode).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Function

            End Select
        
        Case ptxMODULE_UNIT_KBN         'Ӽޭ���Ưċ敪

            If Text1(Mode).Text = "0" Or Text1(Mode).Text = "1" Or Text1(Mode).Text = "2" Or Text1(Mode).Text = "3" Then
            Else
                MsgBox "���͂������ڂ̓G���[�ł��B(Ӽޭ���Ưċ敪)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        
            USE_QTY = Val(Text1(ptxHITUYO_QTY).Text)
                                           
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �ړ� 2015.01.20
'            ZAIKO_QTY = 0
'            For i = 0 To UBound(Nara_Soko_T)
'
'                Location = Nara_Soko_T(i)
'
'                If SOKO_Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode), Location) Then
'                    Exit Function
'                End If
'                ZAIKO_QTY = ZAIKO_QTY + (SUMI_QTY + MI_QTY)
'            Next i
'            Text1(ptxZAIKO_NOW).Text = Format(ZAIKO_QTY, "#")
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �ړ� 2015.01.20
        
        
            Call HANTEI_Proc(HANTEI_MARK, Zaiko_Qty)    '2015.01.20 �����ǉ�
                                                           
                                                           
                                                           
            LblHantei_MARK.Caption = HANTEI_MARK
        
        
        Case ptxMODULE_KBN              'Ӽޭ�ّΏ�
        
            If Text1(Mode).Text = "0" Or Text1(Mode).Text = "1" Then
            Else
                MsgBox "���͂������ڂ̓G���[�ł��B(Ӽޭ�ّΏ�)"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        
            USE_QTY = Val(Text1(ptxHITUYO_QTY).Text)
                                           
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �ړ� 2015.01.20
'            ZAIKO_QTY = 0
'            For i = 0 To UBound(Nara_Soko_T)
'
'                Location = Nara_Soko_T(i)
'
'                If SOKO_Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode), Location) Then
'                    Exit Function
'                End If
'                ZAIKO_QTY = ZAIKO_QTY + (SUMI_QTY + MI_QTY)
'            Next i
'            Text1(ptxZAIKO_NOW).Text = Format(ZAIKO_QTY, "#")
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �ړ� 2015.01.20
        
            Call HANTEI_Proc(HANTEI_MARK, Zaiko_Qty)    '2015.01.20 �����ǉ�
                                                           
                                                           
                                                           
            LblHantei_MARK.Caption = HANTEI_MARK
        
        
        Case ptxNAI_BUHIN               '���������敪
        

        Case ptxKENSA_JIGU              '��������
        
            USE_QTY = Val(Text1(ptxHITUYO_QTY).Text)
                                           
 '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �ړ� 2015.01.20
'           ZAIKO_QTY = 0
'            For i = 0 To UBound(Nara_Soko_T)
'
'                Location = Nara_Soko_T(i)
'
'                If SOKO_Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode), Location) Then
'                    Exit Function
'                End If
'                ZAIKO_QTY = ZAIKO_QTY + (SUMI_QTY + MI_QTY)
'            Next i
'            Text1(ptxZAIKO_NOW).Text = Format(ZAIKO_QTY, "#")
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �ړ� 2015.01.20
                                           
            Call HANTEI_Proc(HANTEI_MARK, Zaiko_Qty)    '2015.01.20 �����ǉ�
                                           
                                           
                                           
            LblHantei_MARK.Caption = HANTEI_MARK
        
        
        
        Case ptxSETUHEN_KBN             '�݌v�ύX�Ώ�

            If Text1(Mode).Text = "0" Or Text1(Mode).Text = "1" Then
            Else
                MsgBox "���͂������ڂ̓G���[�ł��B(�݌v�ύX�Ώ�)"
                Text1(Mode).SetFocus
                Exit Function
            
            End If
            USE_QTY = Val(Text1(ptxHITUYO_QTY).Text)
                                           
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �ړ� 2015.01.20
'            ZAIKO_QTY = 0
'            For i = 0 To UBound(Nara_Soko_T)
'
'                Location = Nara_Soko_T(i)
'
'                If SOKO_Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode), Location) Then
'                    Exit Function
'                End If
'                ZAIKO_QTY = ZAIKO_QTY + (SUMI_QTY + MI_QTY)
'            Next i
'            Text1(ptxZAIKO_NOW).Text = Format(ZAIKO_QTY, "#")
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �ړ� 2015.01.20

            Call HANTEI_Proc(HANTEI_MARK, Zaiko_Qty)    '2015.01.20 �����ǉ�
                                                           
                                                           
                                                           
            LblHantei_MARK.Caption = HANTEI_MARK


'        Case ptxSETUHEN_LAST_DATE       '�݌v�ύX�ŏI��
'
'            If IsDate(Text1(Mode).Text) Then
'            Else
'                MsgBox "���͂������ڂ̓G���[�ł��B(�݌v�ύX�ŏI��)"
'                Text1(Mode).SetFocus
'                Exit Function
'            End If
'
'
'        Case ptxSENDO_LAST_DATE         '�N�x�Ǘ��ŏI��
'
'            If Trim(Text1(Mode).Text) = "" Then
'            Else
'                If IsDate(Text1(Mode).Text) Then
'                Else
'                    MsgBox "���͂������ڂ̓G���[�ł��B(�N�x�Ǘ��ŏI��)"
'                    Text1(Mode).SetFocus
'                    Exit Function
'                End If
'            End If

        Case ptxHITUYO_SU               '�K�v���@��
        
            If IsNumeric(Text1(Mode).Text) Then
            Else
                MsgBox "���͂������ڂ̓G���[�ł��B(�K�v���@(��))"
                Text1(Mode).SetFocus
                Exit Function
            End If
            
        
            If IsNumeric(Text1(ptxHITUYO_SU).Text) And IsNumeric(Text1(ptxHITUYO_TUKI).Text) Then
                Text1(ptxHITUYO_QTY).Text = Val((Text1(ptxHITUYO_SU).Text)) * Val(Text1(ptxHITUYO_TUKI).Text)
            End If
            
            USE_QTY = Val(Text1(ptxHITUYO_QTY).Text)
                                           
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �ړ� 2015.01.20
'            ZAIKO_QTY = 0
'            For i = 0 To UBound(Nara_Soko_T)
'
'                Location = Nara_Soko_T(i)
'
'                If SOKO_Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode), Location) Then
'                    Exit Function
'                End If
'                ZAIKO_QTY = ZAIKO_QTY + (SUMI_QTY + MI_QTY)
'            Next i
'            Text1(ptxZAIKO_NOW).Text = Format(ZAIKO_QTY, "#")
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �ړ� 2015.01.20
        
            Call HANTEI_Proc(HANTEI_MARK, Zaiko_Qty)    '2015.01.20 �����ǉ�
                                                           
                                                           
                                                           
            LblHantei_MARK.Caption = HANTEI_MARK
     
        
        Case ptxHITUYO_TUKI             '�K�v���@��
        
            If IsNumeric(Text1(Mode).Text) Then
            Else
                MsgBox "���͂������ڂ̓G���[�ł��B(�K�v���@(��))"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
            If IsNumeric(Text1(ptxHITUYO_SU).Text) And IsNumeric(Text1(ptxHITUYO_TUKI).Text) Then
                Text1(ptxHITUYO_QTY).Text = Val((Text1(ptxHITUYO_SU).Text)) * Val(Text1(ptxHITUYO_TUKI).Text)
            End If
        
            'lblTUKI.Caption = Val(Text1(Mode).Text)    2014.07.17
        
            USE_QTY = Val(Text1(ptxHITUYO_QTY).Text)
                                           
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �ړ� 2015.01.20
'            ZAIKO_QTY = 0
'            For i = 0 To UBound(Nara_Soko_T)
'
'                Location = Nara_Soko_T(i)
'
'                If SOKO_Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode), Location) Then
'                    Exit Function
'                End If
'                ZAIKO_QTY = ZAIKO_QTY + (SUMI_QTY + MI_QTY)
'            Next i
'            Text1(ptxZAIKO_NOW).Text = Format(ZAIKO_QTY, "#")
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �ړ� 2015.01.20
        
        
            Call HANTEI_Proc(HANTEI_MARK, Zaiko_Qty)    '2015.01.20 �����ǉ�
                                                           
                                                           
                                                           
            LblHantei_MARK.Caption = HANTEI_MARK
        
        Case ptxHITUYO_QTY              '�K�v���@�~��
        
        
    
    
    End Select
        
        
        
        
        
    Error_Check_Proc = False
    

End Function



Public Sub Clear_Field_Proc()
'----------------------------------------------------------------------------
'                   ��ʏ���
'----------------------------------------------------------------------------
Dim i       As Integer


    For i = ptxTANTO_CODE To ptxZAIKO_NOW
    
        Text1(i).Text = ""
    
    Next

'    lblTUKI.Caption = ""   2014.07.17

    LblHantei_MARK = ""     '2015.01.20
    
    Set PCB_U = Nothing

    Set TDBGrid1.Array = PCB_U
    TDBGrid1.ReBind
    
    TDBGrid1.Update


End Sub

Private Sub HANTEI_Proc(HANTEI_MARK As String, Zaiko_Qty As Long)
'----------------------------------------------------------------------------
'                   ����
'       2014.07.03
'       2015.01.20 Zaiko_Qty As Long �ǉ�
'----------------------------------------------------------------------------
Dim USE_QTY     As Long
'Dim Zaiko_Qty   As Long    2015.01.20
Dim Sumi_Qty    As Long
Dim Mi_Qty      As Long


Dim Location    As String * 8
    
Dim i           As Integer
    
    HANTEI_MARK = ""
    If Text1(ptxMODULE_KBN).Text = "0" Then
        HANTEI_MARK = "Ӽޭ�ّΏۊO"
    End If
                                   
    If Trim(HANTEI_MARK) = "" Then
        If Text1(ptxNAI_BUHIN).Text = "0" Then
            HANTEI_MARK = "0��Ώ�"
        End If
    End If
                                   
    If Trim(HANTEI_MARK) = "" Then
        If Text1(ptxNAI_BUHIN).Text = "3" Then
            HANTEI_MARK = "3�Ő؂�"
        End If
        
    End If
                                   
    If Trim(HANTEI_MARK) = "" Then
        If Text1(ptxKENSA_JIGU).Text = "1" Then
            HANTEI_MARK = "����Ȃ�"
        End If
    End If

    If Trim(HANTEI_MARK) = "" Then
        If Text1(ptxSETUHEN_KBN).Text = "1" Then
            HANTEI_MARK = "�~�ݕϗL��"
        End If
    End If

    USE_QTY = Val(Text1(ptxHITUYO_QTY).Text)
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �ړ� 2015.01.20
'    ZAIKO_QTY = 0
'    For i = 0 To UBound(Nara_Soko_T)
'
'        Location = Nara_Soko_T(i)
'
'        If SOKO_Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, Last_JGYOBU, DEF_NAIGAI, Text1(ptxHIN_GAI).Text, Location) Then
'            Exit Sub
'        End If
'        ZAIKO_QTY = ZAIKO_QTY + (SUMI_QTY + MI_QTY)
'    Next i
'    Text1(ptxZAIKO_NOW).Text = Format(ZAIKO_QTY, "#")
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �ړ� 2015.01.20
    
    
    If Trim(HANTEI_MARK) = "" Then
        If USE_QTY >= 200 Then
            If Zaiko_Qty >= USE_QTY Then
                HANTEI_MARK = "����`�@ �N�x�m�F"
            Else
                HANTEI_MARK = "����a�@ �Đ����"
            End If
        Else
            If Zaiko_Qty >= USE_QTY Then
                HANTEI_MARK = "����b�@ �N�x�m�F"
            Else
                HANTEI_MARK = "����c�@ �Đ����"
            End If
        End If
    End If

End Sub
