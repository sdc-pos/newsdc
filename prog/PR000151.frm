VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PR000151 
   Caption         =   "������ь����E����W�v�\���s [PR00015] 2013.06.04 15:00"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14985
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
   ScaleHeight     =   10305
   ScaleWidth      =   14985
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   2
      Left            =   2280
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   5
      Top             =   960
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "�o�͑Ώ�"
      Height          =   855
      Left            =   9000
      TabIndex        =   25
      Top             =   240
      Width           =   5295
      Begin VB.CheckBox Check1 
         Caption         =   "���㖾�ו\"
         Height          =   375
         Index           =   1
         Left            =   3240
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "���Ӑ�ʔ���W�v�\"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   6720
      MaxLength       =   7
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   1
      Left            =   2280
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   2280
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   5655
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   9975
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "�N����"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "���Ӑ�"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "�i��"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "�i��"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "�̔��敪"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "���x�P��"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "����"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "�P��"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "���z"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1984"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=5345"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5239"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=2090"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1984"
      Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=0"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=3175"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=3069"
      Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=0"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=1879"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=1773"
      Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=0"
      Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(25)=   "Column(5).Width=1879"
      Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=1773"
      Splits(0)._ColumnProps(28)=   "Column(5)._ColStyle=0"
      Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(30)=   "Column(6).Width=2699"
      Splits(0)._ColumnProps(31)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(6)._WidthInPix=2593"
      Splits(0)._ColumnProps(33)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(34)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(35)=   "Column(7).Width=2699"
      Splits(0)._ColumnProps(36)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(7)._WidthInPix=2593"
      Splits(0)._ColumnProps(38)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(39)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(40)=   "Column(8).Width=2699"
      Splits(0)._ColumnProps(41)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(8)._WidthInPix=2593"
      Splits(0)._ColumnProps(43)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(44)=   "Column(8).Order=9"
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
      Caption         =   "���㖾�ו\"
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
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
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
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=16,.parent=43"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=13,.parent=44"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=14,.parent=45"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=15,.parent=47"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=28,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(49)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(50)  =   ":id=28,.fontname=�l�r �S�V�b�N"
      _StyleDefs(51)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=44"
      _StyleDefs(52)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=45"
      _StyleDefs(53)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=47"
      _StyleDefs(54)  =   "Splits(0).Columns(3).Style:id=66,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(55)  =   ":id=66,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(56)  =   ":id=66,.fontname=�l�r �S�V�b�N"
      _StyleDefs(57)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=44"
      _StyleDefs(58)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=45"
      _StyleDefs(59)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=47"
      _StyleDefs(60)  =   "Splits(0).Columns(4).Style:id=32,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(61)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(62)  =   ":id=32,.fontname=�l�r �S�V�b�N"
      _StyleDefs(63)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=44"
      _StyleDefs(64)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=45"
      _StyleDefs(65)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=47"
      _StyleDefs(66)  =   "Splits(0).Columns(5).Style:id=74,.parent=43,.alignment=0"
      _StyleDefs(67)  =   "Splits(0).Columns(5).HeadingStyle:id=71,.parent=44"
      _StyleDefs(68)  =   "Splits(0).Columns(5).FooterStyle:id=72,.parent=45"
      _StyleDefs(69)  =   "Splits(0).Columns(5).EditorStyle:id=73,.parent=47"
      _StyleDefs(70)  =   "Splits(0).Columns(6).Style:id=20,.parent=43,.alignment=1"
      _StyleDefs(71)  =   "Splits(0).Columns(6).HeadingStyle:id=17,.parent=44"
      _StyleDefs(72)  =   "Splits(0).Columns(6).FooterStyle:id=18,.parent=45"
      _StyleDefs(73)  =   "Splits(0).Columns(6).EditorStyle:id=19,.parent=47"
      _StyleDefs(74)  =   "Splits(0).Columns(7).Style:id=24,.parent=43,.alignment=1"
      _StyleDefs(75)  =   "Splits(0).Columns(7).HeadingStyle:id=21,.parent=44"
      _StyleDefs(76)  =   "Splits(0).Columns(7).FooterStyle:id=22,.parent=45"
      _StyleDefs(77)  =   "Splits(0).Columns(7).EditorStyle:id=23,.parent=47"
      _StyleDefs(78)  =   "Splits(0).Columns(8).Style:id=78,.parent=43,.alignment=1"
      _StyleDefs(79)  =   "Splits(0).Columns(8).HeadingStyle:id=75,.parent=44"
      _StyleDefs(80)  =   "Splits(0).Columns(8).FooterStyle:id=76,.parent=45"
      _StyleDefs(81)  =   "Splits(0).Columns(8).EditorStyle:id=77,.parent=47"
      _StyleDefs(82)  =   "Named:id=33:Normal"
      _StyleDefs(83)  =   ":id=33,.parent=0"
      _StyleDefs(84)  =   "Named:id=34:Heading"
      _StyleDefs(85)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(86)  =   ":id=34,.wraptext=-1"
      _StyleDefs(87)  =   "Named:id=35:Footing"
      _StyleDefs(88)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(89)  =   "Named:id=36:Selected"
      _StyleDefs(90)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(91)  =   "Named:id=37:Caption"
      _StyleDefs(92)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(93)  =   "Named:id=38:HighlightRow"
      _StyleDefs(94)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(95)  =   "Named:id=39:EvenRow"
      _StyleDefs(96)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(97)  =   "Named:id=40:OddRow"
      _StyleDefs(98)  =   ":id=40,.parent=33"
      _StyleDefs(99)  =   "Named:id=41:RecordSelector"
      _StyleDefs(100) =   ":id=41,.parent=34"
      _StyleDefs(101) =   "Named:id=42:FilterBar"
      _StyleDefs(102) =   ":id=42,.parent=33"
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
      Left            =   10440
      TabIndex        =   21
      Top             =   9720
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
      Left            =   9600
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   9720
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
      Left            =   8760
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   9720
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
      Index           =   8
      Left            =   7920
      TabIndex        =   18
      Top             =   9720
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
      Left            =   6600
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   9720
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
      Left            =   5760
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   9720
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
      Left            =   4920
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   9720
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
      Left            =   4080
      TabIndex        =   14
      Top             =   9720
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
      Index           =   3
      Left            =   2760
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9720
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
      Left            =   1920
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9720
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
      Left            =   1080
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�� ��"
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
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   9720
      Width           =   855
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   1575
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   1800
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   2778
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "���x�P��"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "���Ӑ�"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "�̔�"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "����"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "�ƒ�"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "���̑�"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "���v"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "�h��"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "���v"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1984"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=4207"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4101"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=2090"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1984"
      Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=2699"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2593"
      Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=2699"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=2593"
      Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(25)=   "Column(5).Width=2699"
      Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=2593"
      Splits(0)._ColumnProps(28)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(30)=   "Column(6).Width=2699"
      Splits(0)._ColumnProps(31)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(6)._WidthInPix=2593"
      Splits(0)._ColumnProps(33)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(34)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(35)=   "Column(7).Width=2699"
      Splits(0)._ColumnProps(36)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(7)._WidthInPix=2593"
      Splits(0)._ColumnProps(38)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(39)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(40)=   "Column(8).Width=2699"
      Splits(0)._ColumnProps(41)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(8)._WidthInPix=2593"
      Splits(0)._ColumnProps(43)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(44)=   "Column(8).Order=9"
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
      Caption         =   "���Ӑ�ʔ���W�v�\�@���x�P�ʁi�����j�ʏW�v"
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
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
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
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=16,.parent=43"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=13,.parent=44"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=14,.parent=45"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=15,.parent=47"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=28,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(49)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(50)  =   ":id=28,.fontname=�l�r �S�V�b�N"
      _StyleDefs(51)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=44"
      _StyleDefs(52)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=45"
      _StyleDefs(53)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=47"
      _StyleDefs(54)  =   "Splits(0).Columns(3).Style:id=66,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(55)  =   ":id=66,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(56)  =   ":id=66,.fontname=�l�r �S�V�b�N"
      _StyleDefs(57)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=44"
      _StyleDefs(58)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=45"
      _StyleDefs(59)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=47"
      _StyleDefs(60)  =   "Splits(0).Columns(4).Style:id=32,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(61)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(62)  =   ":id=32,.fontname=�l�r �S�V�b�N"
      _StyleDefs(63)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=44"
      _StyleDefs(64)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=45"
      _StyleDefs(65)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=47"
      _StyleDefs(66)  =   "Splits(0).Columns(5).Style:id=74,.parent=43,.alignment=1"
      _StyleDefs(67)  =   "Splits(0).Columns(5).HeadingStyle:id=71,.parent=44"
      _StyleDefs(68)  =   "Splits(0).Columns(5).FooterStyle:id=72,.parent=45"
      _StyleDefs(69)  =   "Splits(0).Columns(5).EditorStyle:id=73,.parent=47"
      _StyleDefs(70)  =   "Splits(0).Columns(6).Style:id=20,.parent=43,.alignment=1"
      _StyleDefs(71)  =   "Splits(0).Columns(6).HeadingStyle:id=17,.parent=44"
      _StyleDefs(72)  =   "Splits(0).Columns(6).FooterStyle:id=18,.parent=45"
      _StyleDefs(73)  =   "Splits(0).Columns(6).EditorStyle:id=19,.parent=47"
      _StyleDefs(74)  =   "Splits(0).Columns(7).Style:id=24,.parent=43,.alignment=1"
      _StyleDefs(75)  =   "Splits(0).Columns(7).HeadingStyle:id=21,.parent=44"
      _StyleDefs(76)  =   "Splits(0).Columns(7).FooterStyle:id=22,.parent=45"
      _StyleDefs(77)  =   "Splits(0).Columns(7).EditorStyle:id=23,.parent=47"
      _StyleDefs(78)  =   "Splits(0).Columns(8).Style:id=78,.parent=43,.alignment=1"
      _StyleDefs(79)  =   "Splits(0).Columns(8).HeadingStyle:id=75,.parent=44"
      _StyleDefs(80)  =   "Splits(0).Columns(8).FooterStyle:id=76,.parent=45"
      _StyleDefs(81)  =   "Splits(0).Columns(8).EditorStyle:id=77,.parent=47"
      _StyleDefs(82)  =   "Named:id=33:Normal"
      _StyleDefs(83)  =   ":id=33,.parent=0"
      _StyleDefs(84)  =   "Named:id=34:Heading"
      _StyleDefs(85)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(86)  =   ":id=34,.wraptext=-1"
      _StyleDefs(87)  =   "Named:id=35:Footing"
      _StyleDefs(88)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(89)  =   "Named:id=36:Selected"
      _StyleDefs(90)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(91)  =   "Named:id=37:Caption"
      _StyleDefs(92)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(93)  =   "Named:id=38:HighlightRow"
      _StyleDefs(94)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(95)  =   "Named:id=39:EvenRow"
      _StyleDefs(96)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(97)  =   "Named:id=40:OddRow"
      _StyleDefs(98)  =   ":id=40,.parent=33"
      _StyleDefs(99)  =   "Named:id=41:RecordSelector"
      _StyleDefs(100) =   ":id=41,.parent=34"
      _StyleDefs(101) =   "Named:id=42:FilterBar"
      _StyleDefs(102) =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�̔��敪"
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   26
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "����N���x"
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   24
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���Ӑ�"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   23
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���x�P��"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   22
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "PR000151"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SAVE_SYUSHI         As String * 3
Dim SAVE_TOKUI          As String * 5

Dim SUM_URIAGE(0 To 6)  As Long



'�e�L�X�g�p�Y��
Private Const ptxG_SYUSHI_CODE% = 0         '���x����
Private Const ptxTOKUI_CODE% = 1            '���Ӑ溰��
Private Const ptxG_HANBAI_KBN% = 2          '�̔��敪
Private Const ptxKEIJYO_YM% = 3             '�Ώ۔N��

'�R���{�p�Y��
Private Const pcmbG_SYUSHI% = 0             '���x�P��
Private Const pcmbTOKUI% = 1                '���Ӑ�
Private Const pcmbG_HANBAI_KBN% = 2         '�̔��敪

'�`�F�b�N�{�b�N�X�p�Y��
Private Const pchkG_SHURIAGE% = 0           '���Ӑ�ʔ���W�v�\
Private Const pchkD_SHURIAGE% = 1           '���㖾�ו\

'Glid�p��---------------------------------



'����W�v
Private Const pGridTOTAL% = 0

Private G_SHURIAGE    As New XArrayDB

Private Const G_Min_Row% = 1                '�ŏ��s��
Private Const G_Min_Col% = 0                '�ŏ���
Private Const G_Max_Col% = 8                '�ő��

Private Const colG_SYUSHI% = 0              '���x�P��
Private Const colG_TOKUI% = 1               '���Ӑ於��
Private Const colG_URIAGE01% = 2            '�̔�
Private Const colG_URIAGE02% = 3            '����
Private Const colG_URIAGE03% = 4            '�ƒ�
Private Const colG_URIAGE04% = 5            '���̑�
Private Const colG_SUBTOTAL% = 6            '���v
Private Const colG_URIAGE05% = 7            '�h��
Private Const colG_TOTAL% = 8               '���v


Private G_Sort_Tbl(G_Min_Col To G_Max_Col) _
                As Integer                  '��Ă̐��� 0:���� 1:�~��
Private G_Tbl_Set_F   As Boolean


'���㖾��
Private Const pGridDETAIL% = 1


Private D_SHURIAGE    As New XArrayDB


Private Const D_Min_Row% = 1                '�ŏ��s��
Private Const D_Min_Col% = 0                '�ŏ���
Private Const D_Max_Col% = 8                '�ő��

Private Const colD_URIAGE_DT% = 0           '�N�����i������t�j
Private Const colD_TOKUI% = 1               '���Ӑ於��
Private Const colD_HIN_GAI% = 2             '�i��
Private Const colD_HIN_NAME% = 3            '�i��
Private Const colD_HANBAI_KBN% = 4          '�̔��敪
Private Const colD_SYUSHI% = 5              '���x
Private Const colD_URIAGE_QTY% = 6          '����
Private Const colD_TANKA% = 7               '�P��
Private Const colD_KINGAKU% = 8             '���z

Private D_Sort_Tbl(D_Min_Col To D_Max_Col) _
                As Integer                  '��Ă̐��� 0:���� 1:�~��
Private D_Tbl_Set_F   As Boolean

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PR000151.MousePointer = vbHourglass

    Call Ctrl_Lock(PR000151)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PR000151)


    PR000151.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
    
Dim sts         As Integer
Dim com         As Integer
    
Dim wkdate      As String * 10

Dim i           As Integer
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        
        
        Case ptxG_SYUSHI_CODE       '���x�P��
            
            
           
           Combo1(pcmbG_SYUSHI).ListIndex = -1
           For i = 0 To Combo1(pcmbG_SYUSHI).ListCount - 1
               If Text1(ptxG_SYUSHI_CODE).Text = Trim(Right(Combo1(pcmbG_SYUSHI).List(i), 3)) Then
                   Combo1(pcmbG_SYUSHI).ListIndex = i
                   Exit For
               End If
           
           Next i
        
        Case ptxTOKUI_CODE   '���Ӑ�
        
            
            
            
            Combo1(pcmbTOKUI).ListIndex = -1
            For i = 0 To Combo1(pcmbTOKUI).ListCount - 1
                If Text1(ptxTOKUI_CODE).Text = Trim(Right(Combo1(pcmbTOKUI).List(i), 5)) Then
                    Combo1(pcmbTOKUI).ListIndex = i
                    Exit For
                End If
            
            Next i
        
        Case ptxG_HANBAI_KBN    '�̔��敪
        
            
            Combo1(pcmbG_HANBAI_KBN).ListIndex = -1
            For i = 0 To Combo1(pcmbTOKUI).ListCount - 1
                If Text1(ptxG_HANBAI_KBN).Text = Trim(Left(Right(Combo1(pcmbG_HANBAI_KBN).List(i), 3), 2)) Then
                    Combo1(pcmbG_HANBAI_KBN).ListIndex = i
                    Exit For
                End If
            
            Next i
        
        Case ptxKEIJYO_YM       '�Ώ۔N��
        
            wkdate = Text1(ptxKEIJYO_YM).Text & "/01"
            
            If Not IsDate(wkdate) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(Mode).SetFocus
                Exit Function
            Else
                
                wkdate = Format(CDate(Text1(ptxKEIJYO_YM).Text), "YYYY/MM/DD")
                
                Text1(ptxKEIJYO_YM).Text = Mid(wkdate, 1, 7)
            End If
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
    
    Select Case Index
        Case pcmbG_SYUSHI           '���x�P��
            Text1(ptxG_SYUSHI_CODE).Text = Trim(Right(Combo1(pcmbG_SYUSHI).Text, 3))
        Case pcmbTOKUI              '���Ӑ�
            Text1(ptxTOKUI_CODE).Text = Trim(Right(Combo1(pcmbTOKUI).Text, 5))
        Case pcmbG_HANBAI_KBN       '�̔��敪
            Text1(ptxG_HANBAI_KBN).Text = Trim(Left(Right(Combo1(pcmbG_HANBAI_KBN).Text, 3), 2))
    End Select
    
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    
    Select Case Index
        Case pcmbG_SYUSHI           '���x�P��
            Text1(ptxG_SYUSHI_CODE).Text = Trim(Right(Combo1(pcmbG_SYUSHI).Text, 3))
        Case pcmbTOKUI              '���Ӑ�
            Text1(ptxTOKUI_CODE).Text = Trim(Right(Combo1(pcmbTOKUI).Text, 5))
        Case pcmbG_HANBAI_KBN       '�̔��敪
            Text1(ptxG_HANBAI_KBN).Text = Trim(Left(Right(Combo1(pcmbG_HANBAI_KBN).Text, 3), 2))
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans         As Integer
Dim i           As Integer


Dim sts         As Integer

    Select Case Index
        Case P_CMD_Upd          '�X�V
        
            If Kingaku_Proc() Then
                Unload Me
            End If
        
        Case P_CMD_DEL          '�폜
        
        Case P_CMD_DSP                      '����/�\��
        
            For i = ptxG_SYUSHI_CODE To ptxKEIJYO_YM
            
                If Error_Check_Proc(i) Then     '�G���[�`�F�b�N
                    Exit Sub
                End If
            
            Next i
        
        
            If List_Disp_Proc() Then
                Unload Me
            End If
        
            Text1(ptxG_SYUSHI_CODE).SetFocus
        
        
        Case P_CMD_OUT                      '�ް��o��
        
        Case P_CMD_PRT                      '���
 
            For i = ptxG_SYUSHI_CODE To ptxKEIJYO_YM
            
                If Error_Check_Proc(i) Then     '�G���[�`�F�b�N
                    Exit Sub
                End If
            
            Next i
                
            ans = MsgBox("������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                
                If Check1(pchkG_SHURIAGE).Value = vbChecked Then
                    '����W�v�\
                    If G_Print_Proc() Then
                        Unload Me
                    End If
                End If
            
                If Check1(pchkD_SHURIAGE).Value = vbChecked Then
                    '���㖾�ו\
                    If D_Print_Proc() Then
                        Unload Me
                    End If
                End If
            
            End If
            
            Text1(ptxG_SYUSHI_CODE).SetFocus
            
            
        Case P_CMD_End                      '�I��
    
            Unload Me
    
    End Select

End Sub


Private Sub Form_DblClick()
    PrintForm
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

    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
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
                                '�Ǘ��}�X�^�n�o�d�m
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�󕥐�}�X�^�n�o�d�m
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�R�[�h�}�X�^�n�o�d�m
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ޔ����ް��n�o�d�m
    If P_SHURIAGE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ޔ���W�v�ް��n�o�d�m
    If P_SHURI_SUM_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '���ޔ����ް�(�ꎞ̧��)�n�o�d�m
    If P_tmpSHURIAGE_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    
    Load PR000152
    
    
    
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
    If Code_Set_Proc(pcmbG_SYUSHI, P_KBN03_CD, 1) Then
        Unload Me
    End If
    
    '�̔��敪�̃Z�b�g
    If Code_Set_Proc(pcmbG_HANBAI_KBN, P_KBN02_CD, 1) Then
        Unload Me
    End If
    
    
    '��ʏ����ݒ�
    If Init_Proc() Then
        Unload Me
    End If

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
                                            
                                            '���ޔ����ް�(�ꎞ̧��)�b�k�n�r�d
    sts = BTRV(BtOpClose, P_SHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), K0_P_SHURIAGE, Len(K0_P_SHURIAGE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ޔ����ް�")
        End If
    End If
                                            
                                            '���ޔ���W�v�ް�(1)�b�k�n�r�d
    sts = BTRV(BtOpClose, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K0_P_SHURI_SUM, Len(K0_P_SHURI_SUM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ޔ���W�v�ް�(1)")
        End If
    End If
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PR000151 = Nothing
    Set PR000152 = Nothing
    Set PR00015F1 = Nothing
    Set PR00015F2 = Nothing


    End
End Sub





Private Sub TDBGrid1_HeadClick(Index As Integer, ByVal ColIndex As Integer)



    Select Case Index
        Case pGridTOTAL         '����W�v
        
            If G_Sort_Tbl(ColIndex) = 0 Then
                G_Sort_Tbl(ColIndex) = 1
            Else
                If G_Sort_Tbl(ColIndex) = 1 Then
                    G_Sort_Tbl(ColIndex) = 0
                End If
            
            End If
            
            If G_Sort_Tbl(ColIndex) = 0 Or G_Sort_Tbl(ColIndex) = 1 Then
                            
                G_SHURIAGE.QuickSort G_Min_Row, G_SHURIAGE.UpperBound(1), ColIndex, G_Sort_Tbl(ColIndex), XTYPE_STRING
                
                Set TDBGrid1(Index).Array = G_SHURIAGE
                
                TDBGrid1(Index).ReBind
                TDBGrid1(Index).Update
                TDBGrid1(Index).MoveFirst
        
        
            End If
        
        
        Case pGridDETAIL        '���㖾��
    
    
            If D_Sort_Tbl(ColIndex) = 0 Then
                D_Sort_Tbl(ColIndex) = 1
            Else
                If D_Sort_Tbl(ColIndex) = 1 Then
                    D_Sort_Tbl(ColIndex) = 0
                End If
            
            End If
            
            If D_Sort_Tbl(ColIndex) = 0 Or D_Sort_Tbl(ColIndex) = 1 Then
                            
                D_SHURIAGE.QuickSort D_Min_Row, D_SHURIAGE.UpperBound(1), ColIndex, D_Sort_Tbl(ColIndex), XTYPE_STRING
                
                Set TDBGrid1(Index).Array = D_SHURIAGE
                
                TDBGrid1(Index).ReBind
                TDBGrid1(Index).Update
                TDBGrid1(Index).MoveFirst
        
        
            End If
    
    
    
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
        
        
    If Error_Check_Proc(Index) Then    '�G���[�`�F�b�N
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '�ړ�
End Sub
Private Function Init_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���͉�ʂ̏����ݒ�
'----------------------------------------------------------------------------
Dim i       As Integer
Dim sts     As Integer


    Init_Proc = True
    
    
    
    For i = ptxG_SYUSHI_CODE To ptxKEIJYO_YM
        Text1(i).Text = ""
    Next i
    '����N��������
    Text1(ptxKEIJYO_YM).Text = Mid(Format(Now, "YYYY/MM/DD"), 1, 7)



    For i = pcmbG_SYUSHI To pcmbG_HANBAI_KBN
        
        Combo1(i).ListIndex = -1
    
    Next i
    '��ď��̏�����
    
    '��ď��̏�����
    For i = 0 To UBound(G_Sort_Tbl)
        G_Sort_Tbl(i) = 0               '��̫�ď���
    Next i
    
    For i = 0 To UBound(D_Sort_Tbl)
        D_Sort_Tbl(i) = 0               '��̫�ď���
    Next i

    D_Sort_Tbl(colD_HIN_NAME) = 9       '��ď��O

    Init_Proc = False

End Function
Private Function Ukeharai_Set_Proc(Index As Integer) As Integer
'----------------------------------------------------------------------------
'                   �󕥐�}�X�^���R���{�ɃZ�b�g����B
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer




Dim i           As Integer
    
    Ukeharai_Set_Proc = True
    
    Combo1(Index).Clear
    
    
    Combo1(Index).AddItem Space(5)

    
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
Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'           ���ޒ����ް��̕\��
'----------------------------------------------------------------------------
Dim sts                 As Integer
Dim com                 As Integer

Dim G_Row               As Long
Dim D_Row               As Long


Dim wKEIJYO_YM          As String * 6
Dim Skip_Flg            As Boolean

Dim i                   As Integer


    List_Disp_Proc = True
    PR000151.MousePointer = vbHourglass
    
    Set G_SHURIAGE = Nothing
    Set D_SHURIAGE = Nothing
    
    G_Row = G_Min_Row - 1
    D_Row = D_Min_Row - 1
       
    SAVE_SYUSHI = ""
    
    For i = 0 To UBound(SUM_URIAGE)
        SUM_URIAGE(i) = 0
    Next i
    
    wKEIJYO_YM = Mid(Format(CDate(Text1(ptxKEIJYO_YM).Text & "/01"), "YYYYMMDD"), 1, 6)
    
    Call UniCode_Conv(K1_P_SHURIAGE.KEIJYO_YM, wKEIJYO_YM)                      '�v��N��
    Call UniCode_Conv(K1_P_SHURIAGE.G_SYUSHI, Text1(ptxG_SYUSHI_CODE).Text)     '���x�P��
    Call UniCode_Conv(K1_P_SHURIAGE.TOKUI_CODE, Text1(ptxTOKUI_CODE).Text)      '���Ӑ�
    Call UniCode_Conv(K1_P_SHURIAGE.URIAGE_DT, "")                              '����N����
    Call UniCode_Conv(K1_P_SHURIAGE.URIAGE_NO, "")                              '����ں��އ�
    
    
    com = BtOpGetGreaterEqual
    
       
       
    Do
    
        DoEvents
        
        Skip_Flg = False       '2013.06.04
    
        sts = BTRV(com, P_SHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), K1_P_SHURIAGE, Len(K1_P_SHURIAGE), 1)
            
        Select Case sts
            Case BtNoErr
            
                '�v��N������ڰ�
                If StrConv(P_SHURIAGE_REC.KEIJYO_YM, vbUnicode) <> wKEIJYO_YM Then
                    Exit Do
                End If
            
                '���x����ڰ�
                If Trim(Text1(ptxG_SYUSHI_CODE).Text) = "" Then
                Else
                    If Trim(Text1(ptxG_SYUSHI_CODE).Text) <> Trim(StrConv(P_SHURIAGE_REC.G_SYUSHI, vbUnicode)) Then
                        Exit Do
                    End If
                End If
                '���Ӑ����ڰ�
                If Trim(Text1(ptxTOKUI_CODE).Text) = "" Then
                Else
                    If Trim(Text1(ptxTOKUI_CODE).Text) <> Trim(StrConv(P_SHURIAGE_REC.TOKUI_CODE, vbUnicode)) Then
                        'Exit Do                2013.06.04
                        Skip_Flg = True         '2013.06.04
                    End If
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���ޔ����ް�")
                Exit Function
        End Select
    
    
        'Skip_Flg = False       2013.06.04
    
        If Trim(Text1(ptxG_HANBAI_KBN).Text) = "" Then
        Else
            If Trim(Text1(ptxG_HANBAI_KBN).Text) <> Trim(StrConv(P_SHURIAGE_REC.G_HANBAI_KBN, vbUnicode)) Then
                Skip_Flg = True
            End If
        End If
        
        If Not Skip_Flg Then
            '�Ώ��ް�
            If Trim(SAVE_SYUSHI) = "" Then
                SAVE_SYUSHI = StrConv(P_SHURIAGE_REC.G_SYUSHI, vbUnicode)
                SAVE_TOKUI = StrConv(P_SHURIAGE_REC.TOKUI_CODE, vbUnicode)
            End If
    
    
            If SAVE_SYUSHI <> StrConv(P_SHURIAGE_REC.G_SYUSHI, vbUnicode) Then
                G_Row = G_Row + 1
                If G_Grid_Set_Proc(G_Row) Then
                    Exit Function
                End If
    
                SAVE_SYUSHI = StrConv(P_SHURIAGE_REC.G_SYUSHI, vbUnicode)
                SAVE_TOKUI = StrConv(P_SHURIAGE_REC.TOKUI_CODE, vbUnicode)
    
                For i = 0 To UBound(SUM_URIAGE)
                    SUM_URIAGE(i) = 0
                Next i
    
    
            End If
    
            If SAVE_TOKUI <> StrConv(P_SHURIAGE_REC.TOKUI_CODE, vbUnicode) Then
                G_Row = G_Row + 1
                If G_Grid_Set_Proc(G_Row) Then
                    Exit Function
                End If
    
                SAVE_TOKUI = StrConv(P_SHURIAGE_REC.TOKUI_CODE, vbUnicode)
    
                For i = 0 To UBound(SUM_URIAGE)
                    SUM_URIAGE(i) = 0
                Next i
    
    
    
            End If
    
            D_Row = D_Row + 1
            If D_Grid_Set_Proc(D_Row) Then
                Exit Function
            End If
        
            '����Ͻ��ǂݍ���
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN02_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SHURIAGE_REC.G_HANBAI_KBN, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_CODEREC.OPTION1, "")
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "����Ͻ�")
                    Exit Function
            End Select
        
            Select Case Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
                Case P_HN_HANBAI            '�̔�
                    SUM_URIAGE(0) = SUM_URIAGE(0) + CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode))
                Case P_HN_SEIZOU            '����
                    SUM_URIAGE(1) = SUM_URIAGE(1) + CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode))
                Case P_HN_YATIN             '�ƒ�
                    SUM_URIAGE(2) = SUM_URIAGE(2) + CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode))
                Case P_HN_ETC               '���̑�
                    SUM_URIAGE(3) = SUM_URIAGE(3) + CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode))
                Case P_HN_HAKEN             '�h��
                    SUM_URIAGE(5) = SUM_URIAGE(5) + CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode))
                Case Else
                    SUM_URIAGE(3) = SUM_URIAGE(3) + CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode))
            End Select
        
        
        
        
        
        
        
        End If
        
        com = BtOpGetNext
    
    Loop
    
    If Trim(SAVE_SYUSHI) <> "" Then
        G_Row = G_Row + 1
        If G_Grid_Set_Proc(G_Row) Then
            Exit Function
        End If
    End If
    
    
    Set TDBGrid1(pGridTOTAL).Array = G_SHURIAGE
    TDBGrid1(pGridTOTAL).ReBind
    TDBGrid1(pGridTOTAL).Update
    TDBGrid1(pGridTOTAL).MoveFirst
    
    Set TDBGrid1(pGridDETAIL).Array = D_SHURIAGE
    TDBGrid1(pGridDETAIL).ReBind
    TDBGrid1(pGridDETAIL).Update
    TDBGrid1(pGridDETAIL).MoveFirst
    
    
    PR000151.MousePointer = vbDefault
    List_Disp_Proc = False
    


End Function

Private Function G_Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           ���ޔ����ް��i���x�^���Ӑ�ʁj�̓��e���د�ނɾ�Ă���
'----------------------------------------------------------------------------
Dim sts As Integer
Dim i   As Integer


    G_Grid_Set_Proc = True
    
    G_SHURIAGE.ReDim G_Min_Row, Row, G_Min_Col, G_Max_Col


    '���x�P��
    G_SHURIAGE(Row, colG_SYUSHI) = SAVE_SYUSHI

    '���Ӑ�
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, SAVE_TOKUI)
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�󕥐�}�X�^")
            Exit Function
    End Select
    G_SHURIAGE(Row, colG_TOKUI) = SAVE_TOKUI & " " & StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode)

    For i = 0 To 3
        SUM_URIAGE(4) = SUM_URIAGE(4) + SUM_URIAGE(i)
    Next i

    SUM_URIAGE(6) = SUM_URIAGE(4) + SUM_URIAGE(5)

    For i = 0 To UBound(SUM_URIAGE)
    
        G_SHURIAGE(Row, i + colG_URIAGE01) = Format(SUM_URIAGE(i), "#,##0")
    
    Next i

    G_Grid_Set_Proc = False

End Function

Private Function D_Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           ���ޔ����ް��i���㖾�ו\�j�̓��e���د�ނɾ�Ă���
'----------------------------------------------------------------------------
Dim sts As Integer
Dim i   As Integer


    D_Grid_Set_Proc = True
    
    D_SHURIAGE.ReDim D_Min_Row, Row, D_Min_Col, D_Max_Col


    '�N����
    D_SHURIAGE(Row, colD_URIAGE_DT) = Mid(StrConv(P_SHURIAGE_REC.URIAGE_DT, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(P_SHURIAGE_REC.URIAGE_DT, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(P_SHURIAGE_REC.URIAGE_DT, vbUnicode), 7, 2)

    '���Ӑ�
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHURIAGE_REC.TOKUI_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�󕥐�}�X�^")
            Exit Function
    End Select
    D_SHURIAGE(Row, colD_TOKUI) = StrConv(P_SHURIAGE_REC.TOKUI_CODE, vbUnicode) & " " & StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode)
    '�i��
    D_SHURIAGE(Row, colD_HIN_GAI) = StrConv(P_SHURIAGE_REC.HIN_GAI, vbUnicode)
    '�i��
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHURIAGE_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHURIAGE_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SHURIAGE_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    End Select
    D_SHURIAGE(Row, colD_HIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    
    '�̔��敪
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN02_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SHURIAGE_REC.G_HANBAI_KBN, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�R�[�h�}�X�^")
            Exit Function
    End Select
    D_SHURIAGE(Row, colD_HANBAI_KBN) = Trim(StrConv(P_SHURIAGE_REC.G_HANBAI_KBN, vbUnicode)) & " " & _
                StrConv(P_CODEREC.C_RNAME, vbUnicode)
    '���x�敪
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN03_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SHURIAGE_REC.G_SYUSHI, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�R�[�h�}�X�^")
            Exit Function
    End Select
    D_SHURIAGE(Row, colD_SYUSHI) = Trim(StrConv(P_SHURIAGE_REC.G_SYUSHI, vbUnicode)) & " " & _
                StrConv(P_CODEREC.C_RNAME, vbUnicode)
    '����
    D_SHURIAGE(Row, colD_URIAGE_QTY) = Format(CDbl(StrConv(P_SHURIAGE_REC.URIAGE_QTY, vbUnicode)), "#,##0.00")
    '�P��
    D_SHURIAGE(Row, colD_TANKA) = Format(CDbl(StrConv(P_SHURIAGE_REC.TANKA, vbUnicode)), "#,##0.00")
    '���z
    D_SHURIAGE(Row, colD_KINGAKU) = Format(CDbl(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)), "#,##0")
    
    D_Grid_Set_Proc = False

End Function


Private Function G_Print_Proc() As Integer
'----------------------------------------------------------------------------
'           �������
'----------------------------------------------------------------------------

Dim Data_Flg        As Boolean


Dim rpt             As New PR00015F1
Dim f               As New PR000152
            
    
    G_Print_Proc = True
    '�W�v�ް��쐬
    If SHURI_SUM_Make1_Proc(Data_Flg) Then
        Exit Function
    End If
    
    If Data_Flg Then
        
        Set rpt = New PR00015F1
    
        '���|�[�g��������܂��B�itrue�F����_�C�A���O���� false�F�Ȃ��j
        rpt.PrintReport False
    
        Set rpt = Nothing
        
'        f.RunReport rpt
'        f.Show
    End If

    G_Print_Proc = False


End Function
Private Function D_Print_Proc() As Integer
'----------------------------------------------------------------------------
'           �������
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim Data_Flg        As Boolean


Dim rpt             As New PR00015F2
Dim f               As New PR000152
            
    
    D_Print_Proc = True
    '�W�v�ް��쐬
    If SHURI_SUM_Make2_Proc(Data_Flg) Then
        Exit Function
    End If
    
    If Not Data_Flg Then
        D_Print_Proc = False
        Exit Function
    End If
            
        
    com = BtOpGetFirst
    
    Do
        
        DoEvents
        
        sts = BTRV(com, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K1_P_SHURI_SUM, Len(K1_P_SHURI_SUM), 1)
            
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���ޔ���W�v�ް�")
                Exit Function
        End Select
    
        
        Set rpt = New PR00015F2
    
        '���|�[�g��������܂��B�itrue�F����_�C�A���O���� false�F�Ȃ��j
        rpt.PrintReport False
    
        Set rpt = Nothing


'                    f.RunReport rpt
'                    f.Show
        
        
    
    
        com = BtOpGetNext
    
    Loop
        
        
 
 
 
 
 
    D_Print_Proc = False



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

Private Function SHURI_SUM_Make2_Proc(Data_Flg As Boolean) As Integer
'----------------------------------------------------------------------------
'                   ���ޔ���W�v�ް��쐬(2)
'----------------------------------------------------------------------------
Dim sts                     As Integer
Dim com                     As Integer

Dim upd_com                 As Integer
Dim Skip_Flg                As Boolean

Dim wKEIJYO_YM              As String * 6

Dim i                       As Integer
    
    
    
    SHURI_SUM_Make2_Proc = True
    PR000151.MousePointer = vbHourglass

    com = BtOpGetFirst

    Do
    
    
        sts = BTRV(com, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K1_P_SHURI_SUM, Len(K1_P_SHURI_SUM), 1)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
            
                Exit Do
            
            
            Case Else
                Call File_Error(sts, com, "���ޔ���W�v�ް�")
        End Select

        sts = BTRV(BtOpDelete, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K1_P_SHURI_SUM, Len(K1_P_SHURI_SUM), 1)
            
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "���ޔ���W�v�ް�")
        End Select

    
        com = BtOpGetNext
    
    Loop
    
    com = BtOpGetFirst

    Do
    
    
        sts = BTRV(com, P_tmpSHURIAGE_POS, P_tmpSHURIAGE_REC, Len(P_tmpSHURIAGE_REC), K0_P_tmpSHURIAGE, Len(K0_P_tmpSHURIAGE), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
            
                Exit Do
            
            
            Case Else
                Call File_Error(sts, com, "���ޔ����ް�(�ꎞ̧��)")
        End Select

        sts = BTRV(BtOpDelete, P_tmpSHURIAGE_POS, P_tmpSHURIAGE_REC, Len(P_tmpSHURIAGE_REC), K0_P_tmpSHURIAGE, Len(K0_P_tmpSHURIAGE), 0)
            
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "���ޔ����ް�(�ꎞ̧��)")
        End Select

    
        com = BtOpGetNext
    
    Loop
    
    
    
    wKEIJYO_YM = Mid(Format(CDate(Text1(ptxKEIJYO_YM).Text & "/01"), "YYYYMMDD"), 1, 6)
    
    Call UniCode_Conv(K1_P_SHURIAGE.KEIJYO_YM, wKEIJYO_YM)                      '�v��N��
    Call UniCode_Conv(K1_P_SHURIAGE.G_SYUSHI, Text1(ptxG_SYUSHI_CODE).Text)     '���x�P��
    Call UniCode_Conv(K1_P_SHURIAGE.TOKUI_CODE, Text1(ptxTOKUI_CODE).Text)      '���Ӑ�
    Call UniCode_Conv(K1_P_SHURIAGE.URIAGE_DT, "")                              '����N����
    Call UniCode_Conv(K1_P_SHURIAGE.URIAGE_NO, "")                              '����ں��އ�
    
    
    com = BtOpGetGreaterEqual
    
       
       
    Do
    
        DoEvents
    
        Skip_Flg = False   '2013.06.04
        sts = BTRV(com, P_SHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), K1_P_SHURIAGE, Len(K1_P_SHURIAGE), 1)
            
        Select Case sts
            Case BtNoErr
            
                '�v��N������ڰ�
                If StrConv(P_SHURIAGE_REC.KEIJYO_YM, vbUnicode) <> wKEIJYO_YM Then
                    Exit Do
                End If
            
                '���x����ڰ�
                If Trim(Text1(ptxG_SYUSHI_CODE).Text) = "" Then
                Else
                    If Trim(Text1(ptxG_SYUSHI_CODE).Text) <> Trim(StrConv(P_SHURIAGE_REC.G_SYUSHI, vbUnicode)) Then
                        Exit Do
                    End If
                End If
                '���Ӑ����ڰ�
                If Trim(Text1(ptxTOKUI_CODE).Text) = "" Then
                Else
                    If Trim(Text1(ptxTOKUI_CODE).Text) <> Trim(StrConv(P_SHURIAGE_REC.TOKUI_CODE, vbUnicode)) Then
                        'Exit Do                     '2013.06.04
                        Skip_Flg = True             '2013.06.04
                    End If
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���ޔ����ް�")
                Exit Function
        End Select
    
    
        'Skip_Flg = False   2013.06.04
    
        If Trim(Text1(ptxG_HANBAI_KBN).Text) = "" Then
        Else
            If Trim(Text1(ptxG_HANBAI_KBN).Text) <> Trim(StrConv(P_SHURIAGE_REC.G_HANBAI_KBN, vbUnicode)) Then
                Skip_Flg = True
            End If
        End If
        
        If Not Skip_Flg Then
            '�Ώ��ް�
            Data_Flg = True
                        
            '�󕥐�Ͻ��ǂݍ���
            Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHURIAGE_REC.TOKUI_CODE, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    '���o�^�͈��
                    Call UniCode_Conv(P_UKEHARAIREC.TORI_KBN, P_TORI_GENERAL)
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�󕥐�Ͻ�")
                    Exit Function
            End Select
            
            
            
            
            '����Ͻ��ǂݍ���
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN02_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SHURIAGE_REC.G_HANBAI_KBN, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    '���o�^�͂��̑�
                    Call UniCode_Conv(P_CODEREC.OPTION1, P_HN_ETC)
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "����Ͻ�")
                    Exit Function
            End Select
            '���ޔ���W�v�ް��ǂݍ���
            Call UniCode_Conv(K1_P_SHURI_SUM.G_SYUSHI, StrConv(P_SHURIAGE_REC.G_SYUSHI, vbUnicode))
            Call UniCode_Conv(K1_P_SHURI_SUM.TOKUI_CODE, StrConv(P_SHURIAGE_REC.TOKUI_CODE, vbUnicode))
        
            sts = BTRV(BtOpGetEqual, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K1_P_SHURI_SUM, Len(K1_P_SHURI_SUM), 1)
            Select Case sts
                Case BtNoErr
                    upd_com = BtOpUpdate
                Case BtErrKeyNotFound
                    upd_com = BtOpInsert
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "���ޔ���W�v�ް�)")
                    Exit Function
            End Select
        
        
            If upd_com = BtOpInsert Then
            
                Call UniCode_Conv(P_SHURI_SUM_REC.G_SYUSHI, StrConv(P_SHURIAGE_REC.G_SYUSHI, vbUnicode))
                
                
                
                Select Case StrConv(P_SHURIAGE_REC.TORI_KBN, vbUnicode)
                    Case P_TORI_SYANAI
                        Call UniCode_Conv(P_SHURI_SUM_REC.TORI_KBN, P_TORI_SYANAI)
                    Case Else
                        Call UniCode_Conv(P_SHURI_SUM_REC.TORI_KBN, P_TORI_GENERAL)
                End Select
                Call UniCode_Conv(P_SHURI_SUM_REC.TOKUI_CODE, StrConv(P_SHURIAGE_REC.TOKUI_CODE, vbUnicode))
            
                For i = 0 To 4
                    Call UniCode_Conv(P_SHURI_SUM_REC.URIAGE_TBL(i).URIAGE, "00000000")
                Next i
            
            End If
        
        
        
            Select Case Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
                Case P_HN_HANBAI            '�̔�
                    
                    Call UniCode_Conv(P_SHURI_SUM_REC.URIAGE_TBL(0).URIAGE, _
                                    Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(0).URIAGE, vbUnicode)) + _
                                    CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)), "00000000"))
                Case P_HN_SEIZOU            '����
                    Call UniCode_Conv(P_SHURI_SUM_REC.URIAGE_TBL(1).URIAGE, _
                                    Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(1).URIAGE, vbUnicode)) + _
                                    CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)), "00000000"))
                Case P_HN_YATIN             '�ƒ�
                    Call UniCode_Conv(P_SHURI_SUM_REC.URIAGE_TBL(2).URIAGE, _
                                    Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(2).URIAGE, vbUnicode)) + _
                                    CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)), "00000000"))
                Case P_HN_ETC               '���̑�
                    Call UniCode_Conv(P_SHURI_SUM_REC.URIAGE_TBL(3).URIAGE, _
                                    Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(3).URIAGE, vbUnicode)) + _
                                    CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)), "00000000"))
                Case P_HN_HAKEN             '�h��
                    Call UniCode_Conv(P_SHURI_SUM_REC.URIAGE_TBL(4).URIAGE, _
                                    Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(4).URIAGE, vbUnicode)) + _
                                    CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)), "00000000"))
                Case Else
                    Call UniCode_Conv(P_SHURI_SUM_REC.URIAGE_TBL(3).URIAGE, _
                                    Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(3).URIAGE, vbUnicode)) + _
                                    CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)), "00000000"))
            End Select
        
        
            sts = BTRV(upd_com, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K0_P_SHURI_SUM, Len(K0_P_SHURI_SUM), 0)
            Select Case sts
                Case BtNoErr
                Case Else
                    Call File_Error(sts, upd_com, "���ޔ���W�v�ް�")
                    Exit Function
            End Select
        
        
        
            sts = BTRV(BtOpInsert, P_tmpSHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), K0_P_tmpSHURIAGE, Len(K0_P_tmpSHURIAGE), 0)
            Select Case sts
                Case BtNoErr
                Case Else
                    Call File_Error(sts, upd_com, "���ޔ����ް�(�ꎞ̧��)")
                    Exit Function
            End Select
        
        
        
        
        End If
        
        com = BtOpGetNext
    
    Loop


    PR000151.MousePointer = vbDefault

    SHURI_SUM_Make2_Proc = False

End Function
Private Function SHURI_SUM_Make1_Proc(Data_Flg As Boolean) As Integer
'----------------------------------------------------------------------------
'                   ���ޔ���W�v�ް��쐬(1)
'----------------------------------------------------------------------------
Dim sts                     As Integer
Dim com                     As Integer

Dim upd_com                 As Integer

Dim wKEIJYO_YM              As String * 6
Dim Skip_Flg                As Boolean

Dim i                       As Integer
    
Dim TOTAL_URIKAKE(0 To 5)   As Long
Dim TOTAL_FURIKAE(0 To 5)   As Long
    
    
Dim YMD                     As String * 8
Dim ZEI                     As Long
    
Dim wkKINGAKU               As Long
    
    
    SHURI_SUM_Make1_Proc = True
    PR000151.MousePointer = vbHourglass

    com = BtOpGetFirst

    Do
    
    
        sts = BTRV(com, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K0_P_SHURI_SUM, Len(K0_P_SHURI_SUM), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
            
                Exit Do
            
            
            Case Else
                Call File_Error(sts, com, "���ޔ���W�v�ް�")
        End Select

        sts = BTRV(BtOpDelete, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K0_P_SHURI_SUM, Len(K0_P_SHURI_SUM), 0)
            
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "���ޔ���W�v�ް�")
        End Select

    
        com = BtOpGetNext
    
    Loop
    
    For i = 0 To UBound(TOTAL_URIKAKE)
        TOTAL_URIKAKE(i) = 0
        TOTAL_FURIKAE(i) = 0
    Next i
    
    
    wKEIJYO_YM = Mid(Format(CDate(Text1(ptxKEIJYO_YM).Text & "/01"), "YYYYMMDD"), 1, 6)
    
    Call UniCode_Conv(K1_P_SHURIAGE.KEIJYO_YM, wKEIJYO_YM)                      '�v��N��
    Call UniCode_Conv(K1_P_SHURIAGE.G_SYUSHI, Text1(ptxG_SYUSHI_CODE).Text)     '���x�P��
    Call UniCode_Conv(K1_P_SHURIAGE.TOKUI_CODE, Text1(ptxTOKUI_CODE).Text)      '���Ӑ�
    Call UniCode_Conv(K1_P_SHURIAGE.URIAGE_DT, "")                              '����N����
    Call UniCode_Conv(K1_P_SHURIAGE.URIAGE_NO, "")                              '����ں��އ�
    
    
    com = BtOpGetGreaterEqual
    
       
       
    Do
    
        DoEvents
        Skip_Flg = False                '2013.06.04
    
        sts = BTRV(com, P_SHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), K1_P_SHURIAGE, Len(K1_P_SHURIAGE), 1)
            
        Select Case sts
            Case BtNoErr
            
                '�v��N������ڰ�
                If StrConv(P_SHURIAGE_REC.KEIJYO_YM, vbUnicode) <> wKEIJYO_YM Then
                    Exit Do
                End If
            
                '���x����ڰ�
                If Trim(Text1(ptxG_SYUSHI_CODE).Text) = "" Then
                Else
                    If Trim(Text1(ptxG_SYUSHI_CODE).Text) <> Trim(StrConv(P_SHURIAGE_REC.G_SYUSHI, vbUnicode)) Then
                        Exit Do
                    End If
                End If
                '���Ӑ����ڰ�
                If Trim(Text1(ptxTOKUI_CODE).Text) = "" Then
                Else
                    If Trim(Text1(ptxTOKUI_CODE).Text) <> Trim(StrConv(P_SHURIAGE_REC.TOKUI_CODE, vbUnicode)) Then
                        'Exit Do                     '2013.06.04
                        Skip_Flg = True            '2013.06.04
                    End If
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���ޔ����ް�")
                Exit Function
        End Select
    
    
        'Skip_Flg = False       2013.06.04
    
        If Trim(Text1(ptxG_HANBAI_KBN).Text) = "" Then
        Else
            If Trim(Text1(ptxG_HANBAI_KBN).Text) <> Trim(StrConv(P_SHURIAGE_REC.G_HANBAI_KBN, vbUnicode)) Then
                Skip_Flg = True
            End If
        End If
        
        If Not Skip_Flg Then
            '�Ώ��ް�
            Data_Flg = True
                        
            
            
            
            
            '����Ͻ��ǂݍ���
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN02_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SHURIAGE_REC.G_HANBAI_KBN, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    '���o�^�͂��̑�
                    Call UniCode_Conv(P_CODEREC.OPTION1, P_HN_ETC)
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "����Ͻ�")
                    Exit Function
            End Select

            Select Case Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
                Case P_HN_HANBAI            '�̔�
                    
                    Select Case StrConv(P_SHURIAGE_REC.TORI_KBN, vbUnicode)
                        Case P_TORI_SYANAI, P_TORI_ANOTHER
                            TOTAL_FURIKAE(0) = TOTAL_FURIKAE(0) + CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode))
                        Case Else
                            TOTAL_URIKAKE(0) = TOTAL_URIKAKE(0) + CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode))
                    End Select
                
                Case P_HN_SEIZOU            '����
                    
                    Select Case StrConv(P_SHURIAGE_REC.TORI_KBN, vbUnicode)
                        Case P_TORI_SYANAI, P_TORI_ANOTHER
                            TOTAL_FURIKAE(1) = TOTAL_FURIKAE(1) + CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode))
                        Case Else
                            TOTAL_URIKAKE(1) = TOTAL_URIKAKE(1) + CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode))
                    End Select
                    
                Case P_HN_YATIN             '�ƒ�
                    
                    
                    Select Case StrConv(P_SHURIAGE_REC.TORI_KBN, vbUnicode)
                        Case P_TORI_SYANAI, P_TORI_ANOTHER
                            TOTAL_FURIKAE(2) = TOTAL_FURIKAE(2) + CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode))
                        Case Else
                            TOTAL_URIKAKE(2) = TOTAL_URIKAKE(2) + CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode))
                    End Select
                    
                Case P_HN_ETC               '���̑�
                    
                    Select Case StrConv(P_SHURIAGE_REC.TORI_KBN, vbUnicode)
                        Case P_TORI_SYANAI, P_TORI_ANOTHER
                            TOTAL_FURIKAE(3) = TOTAL_FURIKAE(3) + CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode))
                        Case Else
                            TOTAL_URIKAKE(3) = TOTAL_URIKAKE(3) + CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode))
                    End Select
                
                Case P_HN_HAKEN             '�h��
                    
                    Select Case StrConv(P_SHURIAGE_REC.TORI_KBN, vbUnicode)
                        Case P_TORI_SYANAI, P_TORI_ANOTHER
                            TOTAL_FURIKAE(4) = TOTAL_FURIKAE(4) + CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode))
                        Case Else
                            TOTAL_URIKAKE(4) = TOTAL_URIKAKE(4) + CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode))
                    End Select
                Case Else
                    
                    Select Case StrConv(P_SHURIAGE_REC.TORI_KBN, vbUnicode)
                        Case P_TORI_SYANAI, P_TORI_ANOTHER
                            TOTAL_FURIKAE(3) = TOTAL_FURIKAE(3) + CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode))
                        Case Else
                            TOTAL_URIKAKE(3) = TOTAL_URIKAKE(3) + CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode))
                    End Select
            End Select
            
            
            
            '����ŕ�
            
'2007.05.24            Select Case StrConv(P_SHURIAGE_REC.TORI_KBN, vbUnicode)
'2007.05.24                Case P_TORI_SYANAI, P_TORI_ANOTHER
'2007.05.24                Case Else
'2007.05.24                    YMD = StrConv(P_SHURIAGE_REC.URIAGE_DT, vbUnicode)
'2007.05.24
'2007.05.24
'2007.05.24                    If CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)) >= 0 Then
'2007.05.24                        If YMD < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
'2007.05.24                            ZEI = Int(CDbl(CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)) * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
'2007.05.24                                    CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
'2007.05.24                        Else
'2007.05.24                            ZEI = Int(CDbl(CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
'2007.05.24                                    CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
'2007.05.24                        End If
'2007.05.24                    Else
'2007.05.24
'2007.05.24                        wkKINGAKU = CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)) * -1
'2007.05.24
'2007.05.24                        If YMD < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
'2007.05.24                            ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
'2007.05.24                                    CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
'2007.05.24                        Else
'2007.05.24                            ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
'2007.05.24                                    CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
'2007.05.24                        End If
'2007.05.24                        ZEI = ZEI * -1
'2007.05.24                    End If
'2007.05.24
'2007.05.24                    TOTAL_URIKAKE(5) = TOTAL_URIKAKE(5) + ZEI
'2007.05.24
'2007.05.24
'2007.05.24            End Select
            
            
            '2007.05.24
            If IsNumeric(StrConv(P_SHURIAGE_REC.ZEI_KIN, vbUnicode)) Then
                TOTAL_URIKAKE(5) = TOTAL_URIKAKE(5) + CLng(StrConv(P_SHURIAGE_REC.ZEI_KIN, vbUnicode))
            End If
            '2007.05.24
            
            
            
            
            
            
            
            
            
            '���ޔ���W�v�ް�(1)�ǂݍ���
            
            Select Case StrConv(P_SHURIAGE_REC.TORI_KBN, vbUnicode)
                Case P_TORI_SYANAI, P_TORI_ANOTHER
                    Call UniCode_Conv(K0_P_SHURI_SUM.TORI_KBN, P_TORI_SYANAI)
                Case Else
                    Call UniCode_Conv(K0_P_SHURI_SUM.TORI_KBN, P_TORI_GENERAL)
            End Select
            Call UniCode_Conv(K0_P_SHURI_SUM.TOKUI_CODE, StrConv(P_SHURIAGE_REC.TOKUI_CODE, vbUnicode))
        
            sts = BTRV(BtOpGetEqual, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K0_P_SHURI_SUM, Len(K0_P_SHURI_SUM), 0)
            Select Case sts
                Case BtNoErr
                    upd_com = BtOpUpdate
                Case BtErrKeyNotFound
                    upd_com = BtOpInsert
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "���ޔ���W�v�ް�")
                    Exit Function
            End Select
        
        
            If upd_com = BtOpInsert Then
            
                Call UniCode_Conv(P_SHURI_SUM_REC.G_SYUSHI, "")
                
                
                Select Case StrConv(P_SHURIAGE_REC.TORI_KBN, vbUnicode)
                    Case P_TORI_SYANAI, P_TORI_ANOTHER
                        Call UniCode_Conv(P_SHURI_SUM_REC.TORI_KBN, P_TORI_SYANAI)
                    Case Else
                        Call UniCode_Conv(P_SHURI_SUM_REC.TORI_KBN, P_TORI_GENERAL)
                End Select
                Call UniCode_Conv(P_SHURI_SUM_REC.TOKUI_CODE, StrConv(P_SHURIAGE_REC.TOKUI_CODE, vbUnicode))
            
                For i = 0 To 5
                    Call UniCode_Conv(P_SHURI_SUM_REC.URIAGE_TBL(i).URIAGE, "00000000")
                Next i
            
            End If
        
        
        
            Select Case Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
                Case P_HN_HANBAI            '�̔�
                    
                    Call UniCode_Conv(P_SHURI_SUM_REC.URIAGE_TBL(0).URIAGE, _
                                    Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(0).URIAGE, vbUnicode)) + _
                                    CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)), "00000000"))
                Case P_HN_SEIZOU            '����
                    Call UniCode_Conv(P_SHURI_SUM_REC.URIAGE_TBL(1).URIAGE, _
                                    Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(1).URIAGE, vbUnicode)) + _
                                    CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)), "00000000"))
                Case P_HN_YATIN             '�ƒ�
                    Call UniCode_Conv(P_SHURI_SUM_REC.URIAGE_TBL(2).URIAGE, _
                                    Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(2).URIAGE, vbUnicode)) + _
                                    CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)), "00000000"))
                Case P_HN_ETC               '���̑�
                    Call UniCode_Conv(P_SHURI_SUM_REC.URIAGE_TBL(3).URIAGE, _
                                    Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(3).URIAGE, vbUnicode)) + _
                                    CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)), "00000000"))
                Case P_HN_HAKEN             '�h��
                    Call UniCode_Conv(P_SHURI_SUM_REC.URIAGE_TBL(4).URIAGE, _
                                    Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(4).URIAGE, vbUnicode)) + _
                                    CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)), "00000000"))
                Case Else
                    Call UniCode_Conv(P_SHURI_SUM_REC.URIAGE_TBL(3).URIAGE, _
                                    Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(3).URIAGE, vbUnicode)) + _
                                    CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)), "00000000"))
            
            End Select
        
'2007.05.24            Select Case StrConv(P_SHURIAGE_REC.TORI_KBN, vbUnicode)
'2007.05.24                Case P_TORI_SYANAI, P_TORI_ANOTHER
'2007.05.24                Case Else
'2007.05.24                    YMD = StrConv(P_SHURIAGE_REC.URIAGE_DT, vbUnicode)
'2007.05.24
'2007.05.24                    If CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)) >= 0 Then
'2007.05.24                        If YMD < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
'2007.05.24                            ZEI = Int(CDbl(CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)) * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
'2007.05.24                                    CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
'2007.05.24                        Else
'2007.05.24                            ZEI = Int(CDbl(CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)) * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
'2007.05.24                                    CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
'2007.05.24                        End If
'2007.05.24                    Else
'2007.05.24
'2007.05.24                        wkKINGAKU = CLng(StrConv(P_SHURIAGE_REC.KINGAKU, vbUnicode)) * -1
'2007.05.24
'2007.05.24                        If YMD < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
'2007.05.24                            ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
'2007.05.24                                    CDbl(CDbl(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)) / 10))
'2007.05.24                        Else
'2007.05.24                            ZEI = Int(CDbl(wkKINGAKU * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
'2007.05.24                                    CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
'2007.05.24                        End If
'2007.05.24                        ZEI = ZEI * -1
'2007.05.24                    End If
'2007.05.24
'2007.05.24
'2007.05.24
'2007.05.24
'2007.05.24                    Call UniCode_Conv(P_SHURI_SUM_REC.URIAGE_TBL(5).URIAGE, _
'2007.05.24                                    Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(5).URIAGE, vbUnicode)) + _
'2007.05.24                                    ZEI, "00000000"))
'2007.05.24
'2007.05.24
'2007.05.24            End Select
        
        
            '2007.05.24
            If IsNumeric(StrConv(P_SHURIAGE_REC.ZEI_KIN, vbUnicode)) Then
                Call UniCode_Conv(P_SHURI_SUM_REC.URIAGE_TBL(5).URIAGE, _
                                Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(5).URIAGE, vbUnicode)) + _
                                CLng(StrConv(P_SHURIAGE_REC.ZEI_KIN, vbUnicode)), "00000000"))
            End If
            '2007.05.24
        
        
            sts = BTRV(upd_com, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K0_P_SHURI_SUM, Len(K0_P_SHURI_SUM), 0)
            Select Case sts
                Case BtNoErr
                Case Else
                    Call File_Error(sts, upd_com, "���ޔ���W�v�ް�")
                    Exit Function
            End Select
        
        
        
        End If
        
        com = BtOpGetNext
    
    Loop

    If Data_Flg Then
        '���vں���
        Call UniCode_Conv(P_SHURI_SUM_REC.TORI_KBN, P_TORI_GENERAL)
        Call UniCode_Conv(P_SHURI_SUM_REC.TOKUI_CODE, "")
    
        For i = 0 To 5
            Call UniCode_Conv(P_SHURI_SUM_REC.URIAGE_TBL(i).URIAGE, Format(TOTAL_URIKAKE(i)))
        Next i
    
        sts = BTRV(BtOpInsert, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K0_P_SHURI_SUM, Len(K0_P_SHURI_SUM), 0)
        Select Case sts
            Case BtNoErr
            Case Else
                Call File_Error(sts, BtOpInsert, "���ޔ���W�v�ް�")
                Exit Function
        End Select
        '���vں���
        Call UniCode_Conv(P_SHURI_SUM_REC.TORI_KBN, P_TORI_SYANAI)
        Call UniCode_Conv(P_SHURI_SUM_REC.TOKUI_CODE, "")
    
        For i = 0 To 5
            Call UniCode_Conv(P_SHURI_SUM_REC.URIAGE_TBL(i).URIAGE, Format(TOTAL_FURIKAE(i)))
        Next i
    
        sts = BTRV(BtOpInsert, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K0_P_SHURI_SUM, Len(K0_P_SHURI_SUM), 0)
        Select Case sts
            Case BtNoErr
            Case Else
                Call File_Error(sts, BtOpInsert, "���ޔ���W�v�ް�")
                Exit Function
        End Select
    
    
    End If

    PR000151.MousePointer = vbDefault

   SHURI_SUM_Make1_Proc = False

End Function



Private Function Kingaku_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���z�W�v
'----------------------------------------------------------------------------
Dim sts As Integer
Dim com As Integer
Dim Kin As Long

    
    com = BtOpGetFirst
    Do
    
        DoEvents
    
        sts = BTRV(com, P_SHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), K0_P_SHURIAGE, Len(K0_P_SHURIAGE), 0)
            
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���ޔ���W�v�ް�")
                Exit Function
        End Select

        Kin = Round(CDbl(StrConv(P_SHURIAGE_REC.TANKA, vbUnicode)) * CLng(StrConv(P_SHURIAGE_REC.URIAGE_QTY, vbUnicode)), 1)


        If Kin < 0 Then
            Call UniCode_Conv(P_SHURIAGE_REC.KINGAKU, Format(Kin, "00000000"))
        Else
            Call UniCode_Conv(P_SHURIAGE_REC.KINGAKU, Format(Kin, "000000000"))
        End If

        sts = BTRV(BtOpUpdate, P_SHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), K0_P_SHURIAGE, Len(K0_P_SHURIAGE), 0)
            
        Select Case sts
            Case BtNoErr
            Case Else
                Call File_Error(sts, com, "���ޔ���W�v�ް�")
                Exit Function
        End Select
    


        com = BtOpGetNext
    Loop
End Function
