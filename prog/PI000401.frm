VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PI000401 
   Caption         =   "���ގd������"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15915
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
   ScaleWidth      =   15915
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CheckBox Check1 
      Caption         =   "POS�݌Ɍv��"
      Height          =   375
      Index           =   0
      Left            =   10920
      TabIndex        =   21
      Top             =   3840
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   8505
      Sorted          =   -1  'True
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   5
      Top             =   1080
      Width           =   2115
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   8085
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1080
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   11
      Left            =   4830
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   19
      Left            =   8190
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1590
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      Index           =   18
      Left            =   11550
      MaxLength       =   8
      TabIndex        =   20
      Top             =   3360
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      Index           =   17
      Left            =   8190
      MaxLength       =   8
      TabIndex        =   19
      Top             =   3360
      Width           =   1590
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   16
      Left            =   11550
      MaxLength       =   7
      TabIndex        =   18
      Top             =   2760
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   15
      Left            =   8190
      MaxLength       =   10
      TabIndex        =   17
      Top             =   2760
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   13
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   14
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   12
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   10
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   9
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   8
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2400
      Width           =   750
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   1
      Left            =   2310
      Locked          =   -1  'True
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2850
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2040
      Width           =   750
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   2
      Left            =   2310
      Locked          =   -1  'True
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2850
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2115
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1560
      Width           =   750
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   3
      Left            =   4095
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2745
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   1575
      MaxLength       =   5
      TabIndex        =   0
      Top             =   240
      Width           =   750
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   3975
      Left            =   525
      TabIndex        =   23
      Top             =   5160
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   7011
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "��������"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "������"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "�����於"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "���ޕi��"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "�i��"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "������"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "�����c"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "�݌Ɏc"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "�[���\���"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2699"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2593"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2011"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1905"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3810"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3704"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=2699"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2593"
      Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=0"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=4180"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=4075"
      Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=0"
      Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(25)=   "Column(5).Width=2064"
      Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=1958"
      Splits(0)._ColumnProps(28)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(30)=   "Column(6).Width=2064"
      Splits(0)._ColumnProps(31)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(6)._WidthInPix=1958"
      Splits(0)._ColumnProps(33)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(34)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(35)=   "Column(7).Width=2064"
      Splits(0)._ColumnProps(36)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(7)._WidthInPix=1958"
      Splits(0)._ColumnProps(38)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(39)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(40)=   "Column(8).Width=2328"
      Splits(0)._ColumnProps(41)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(8)._WidthInPix=2223"
      Splits(0)._ColumnProps(43)=   "Column(8)._ColStyle=0"
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
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=62,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(45)  =   ":id=62,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(46)  =   ":id=62,.fontname=�l�r �S�V�b�N"
      _StyleDefs(47)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=44"
      _StyleDefs(48)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=45"
      _StyleDefs(49)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=47"
      _StyleDefs(50)  =   "Splits(0).Columns(2).Style:id=16,.parent=43"
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
      _StyleDefs(66)  =   "Splits(0).Columns(5).Style:id=32,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(67)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(68)  =   ":id=32,.fontname=�l�r �S�V�b�N"
      _StyleDefs(69)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=44"
      _StyleDefs(70)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=45"
      _StyleDefs(71)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=47"
      _StyleDefs(72)  =   "Splits(0).Columns(6).Style:id=24,.parent=43,.alignment=1"
      _StyleDefs(73)  =   "Splits(0).Columns(6).HeadingStyle:id=21,.parent=44"
      _StyleDefs(74)  =   "Splits(0).Columns(6).FooterStyle:id=22,.parent=45"
      _StyleDefs(75)  =   "Splits(0).Columns(6).EditorStyle:id=23,.parent=47"
      _StyleDefs(76)  =   "Splits(0).Columns(7).Style:id=20,.parent=43,.alignment=1"
      _StyleDefs(77)  =   "Splits(0).Columns(7).HeadingStyle:id=17,.parent=44"
      _StyleDefs(78)  =   "Splits(0).Columns(7).FooterStyle:id=18,.parent=45"
      _StyleDefs(79)  =   "Splits(0).Columns(7).EditorStyle:id=19,.parent=47"
      _StyleDefs(80)  =   "Splits(0).Columns(8).Style:id=70,.parent=43,.alignment=0,.bold=0,.fontsize=975"
      _StyleDefs(81)  =   ":id=70,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(82)  =   ":id=70,.fontname=�l�r �S�V�b�N"
      _StyleDefs(83)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=44"
      _StyleDefs(84)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=45"
      _StyleDefs(85)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=47"
      _StyleDefs(86)  =   "Named:id=33:Normal"
      _StyleDefs(87)  =   ":id=33,.parent=0"
      _StyleDefs(88)  =   "Named:id=34:Heading"
      _StyleDefs(89)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(90)  =   ":id=34,.wraptext=-1"
      _StyleDefs(91)  =   "Named:id=35:Footing"
      _StyleDefs(92)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(93)  =   "Named:id=36:Selected"
      _StyleDefs(94)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(95)  =   "Named:id=37:Caption"
      _StyleDefs(96)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(97)  =   "Named:id=38:HighlightRow"
      _StyleDefs(98)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(99)  =   "Named:id=39:EvenRow"
      _StyleDefs(100) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(101) =   "Named:id=40:OddRow"
      _StyleDefs(102) =   ":id=40,.parent=33"
      _StyleDefs(103) =   "Named:id=41:RecordSelector"
      _StyleDefs(104) =   ":id=41,.parent=34"
      _StyleDefs(105) =   "Named:id=42:FilterBar"
      _StyleDefs(106) =   ":id=42,.parent=33"
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
      TabIndex        =   35
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
      TabIndex        =   34
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
      TabIndex        =   33
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
      Index           =   8
      Left            =   7920
      TabIndex        =   32
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
      Index           =   7
      Left            =   6600
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�[���ύX"
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
      Index           =   6
      Left            =   5760
      TabIndex        =   30
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
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�� �V"
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
      TabIndex        =   28
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ݾ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   9720
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
      Left            =   240
      TabIndex        =   24
      Top             =   9720
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�d���敪"
      Height          =   255
      Index           =   18
      Left            =   6930
      TabIndex        =   55
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�j"
      Height          =   255
      Index           =   17
      Left            =   5880
      TabIndex        =   54
      Top             =   3480
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�i�����[�i�ϐ�"
      Height          =   255
      Index           =   16
      Left            =   2730
      TabIndex        =   52
      Top             =   3480
      Width           =   1905
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���z"
      Height          =   255
      Index           =   15
      Left            =   7455
      TabIndex        =   51
      Top             =   3840
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�����c"
      Height          =   255
      Index           =   14
      Left            =   10710
      TabIndex        =   50
      Top             =   3480
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "����������"
      Height          =   255
      Index           =   13
      Left            =   6615
      TabIndex        =   49
      Top             =   3480
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�����N��"
      Height          =   255
      Index           =   12
      Left            =   10290
      TabIndex        =   48
      Top             =   2880
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�����"
      Height          =   255
      Index           =   9
      Left            =   7245
      TabIndex        =   47
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�݌Ɏc"
      Height          =   255
      Index           =   10
      Left            =   735
      TabIndex        =   46
      Top             =   4320
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "����ۯ�"
      Height          =   255
      Index           =   11
      Left            =   525
      TabIndex        =   45
      Top             =   4680
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�P��"
      Height          =   255
      Index           =   8
      Left            =   945
      TabIndex        =   44
      Top             =   3840
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "������"
      Height          =   255
      Index           =   7
      Left            =   630
      TabIndex        =   43
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�[���\���"
      Height          =   255
      Index           =   6
      Left            =   210
      TabIndex        =   42
      Top             =   3000
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�[����"
      Height          =   255
      Index           =   5
      Left            =   525
      TabIndex        =   41
      Top             =   2520
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "������"
      Height          =   255
      Index           =   4
      Left            =   525
      TabIndex        =   40
      Top             =   2160
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�S����"
      Height          =   255
      Index           =   2
      Left            =   525
      TabIndex        =   39
      Top             =   1680
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���ޕi��"
      Height          =   255
      Index           =   1
      Left            =   525
      TabIndex        =   38
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "������"
      Height          =   255
      Index           =   0
      Left            =   630
      TabIndex        =   37
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "������"
      Height          =   255
      Index           =   3
      Left            =   630
      TabIndex        =   36
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "PI000401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private NOUKI_MODE  As Boolean
Private Input_Mode  As Boolean

Private WS_NO       As String * 10
    
Private KASO_NYUKA  As String * 2           '���בq��
Private POS_UMU     As Boolean              'POS���т̗L��
    
Private MEMO_TEXT   As String               '��������
   
'�e�L�X�g�p�Y��

Private Const ptxORDER_NO% = 0              '������
Private Const ptxORDER_DT% = 1              '������
Private Const ptxHIN_GAI% = 2               '�i��
Private Const ptxHIN_NAME% = 3              '�i��
Private Const ptxG_SHIIRE_KBN% = 4          '�d���敪
Private Const ptxTANTO_CODE% = 5            '�S���Һ���
Private Const ptxTANTO_NAME% = 6            '�S���Җ���
Private Const ptxORDER_CODE% = 7            '������
Private Const ptxDELI_CODE% = 8             '�[����
Private Const ptxY_NOUKI_DT% = 9            '�[���\���
Private Const ptxORDER_QTY% = 10            '������
Private Const ptxUKEIRE_QTY% = 11           '����ϐ�
Private Const ptxTANKA% = 12                '�P��
Private Const ptxZAIKO_QTY% = 13            '�݌Ɏc
Private Const ptxLOT% = 14                  '����ۯ�
Private Const ptxUKEIRE_DT% = 15            '�����
Private Const ptxKEIJYO_YM% = 16            '�v��N��
Private Const ptxKONKAI_UKEIRE_QTY% = 17    '����[�i����
Private Const ptxZAN_QTY% = 18              '�����c
Private Const ptxKINGAKU% = 19              '���z
'�R���{�p�Y��
Private Const pcmbG_SHIIRE_KBN% = 0         '�d���敪
Private Const pcmbORDER% = 1                '������
Private Const pcmbDELI% = 2                 '�[����


'�R�}���h����@�\
Private Const cmdNOUKI% = 6                 '������

'�����ޯ���p�Y��
Private Const chkZAIKO_F% = 0

'Glid�p��

Private SHORDER  As New XArrayDB

Private Const Min_Row% = 1                  '�ŏ��s��
Private Const Min_Col% = 0                  '�ŏ���
Private Const Max_Col% = 8                  '�ő��


Private Const colORDER_DT% = 0              '������
Private Const colORDER_NO% = 1              '������
Private Const colORDER_NAME% = 2            '�����於
Private Const colHIN_GAI% = 3               '�i��
Private Const colHIN_NAME% = 4              '�i��
Private Const colORDER_QTY% = 5             '������
Private Const colZAN_QTY% = 6               '�����c
Private Const colZAIKO_QTY% = 7             '�݌Ɏc
Private Const colY_NOUKI_DT% = 8            '�[���\���



Private Sort_Tbl(colORDER_DT To colY_NOUKI_DT) _
                As Integer                  '��Ă̐��� 0:���� 1:�~��
Private Tbl_Set_F   As Boolean

Private Save_UKEIRE_QTY     As Long             '������̃Z�[�u
                                            



Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PI000401.MousePointer = vbHourglass

    Call Ctrl_Lock(PI000401)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PI000401)


    PI000401.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
    
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer

Dim SUMI_QTY    As Long
Dim MI_QTY      As Long
    
Dim wkDate      As String
    
    
    Error_Check_Proc = True
    
    Select Case Mode
        
        Case ptxORDER_NO    '������
        
            If Not NOUKI_MODE Then
            
                If Trim(Text1(ptxORDER_NO).Text) = "" Then
                    '�����Ȃ�����
                
                    Call Input_Area_Proc(1)
                
                Else
            
                    Call Input_Area_Proc(0)
            
            
                    If IsNumeric(Text1(ptxORDER_NO).Text) Then
                        Text1(ptxORDER_NO).Text = Format(CLng(Text1(ptxORDER_NO).Text), "00000")
                    End If
                    '���ޒ����ް�������
                    
                    If Text1(ptxORDER_NO).Text = StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode) Then
                        sts = BtNoErr
                    Else
                        sts = P_SHORDER_Read_Proc()
                    End If
                    Select Case sts
                        Case False, BtNoErr
                                    
                            If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = P_KAN_ON Then
                                MsgBox "�d�������ς݂ł��B"
                                Text1(Mode).SetFocus
                                Exit Function
                            End If
                        
                            If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
                                MsgBox "�L�����Z�������ς݂ł��B"
                                Text1(Mode).SetFocus
                                Exit Function
                            End If
                        
                        
                        Case BtErrKeyNotFound
                            MsgBox "���͂������ڂ̓G���[�ł��B"
                            Text1(Mode).SetFocus
                            Exit Function
                        Case Else
                            Exit Function
                    End Select
        
                End If
            End If
        
        Case ptxHIN_GAI     '�i�ԊO
            If Not NOUKI_MODE Then
            
                sts = Hin_Item_Disp_Proc()
                Select Case sts
                    Case False
                    Case BtErrKeyNotFound
                        MsgBox "���͂������ڂ̓G���[�ł��B"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Exit Function
                End Select
            
                If Not POS_UMU Then      '2006.04.26 �o�n�r�Ȃ��Ȃ�
                
                
                    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" And _
                        Trim(StrConv(ITEMREC.ST_RETU, vbUnicode)) = "" And _
                        Trim(StrConv(ITEMREC.ST_REN, vbUnicode)) = "" And _
                        Trim(StrConv(ITEMREC.ST_DAN, vbUnicode)) = "" Then
                
                        MsgBox "�W���I�Ԃ��ݒ肳��Ă��܂���B"
                        Text1(Mode).SetFocus
                        Exit Function
                
                    End If
                
                End If
            
            
            
            
            End If
        
        Case ptxG_SHIIRE_KBN    '�d���敪
            If Not NOUKI_MODE Then
        
                Combo1(pcmbG_SHIIRE_KBN).ListIndex = -1
                For i = 0 To Combo1(pcmbG_SHIIRE_KBN).ListCount - 1
                
                    If Text1(ptxG_SHIIRE_KBN).Text = Trim(Left(Right(Combo1(pcmbG_SHIIRE_KBN).List(i), 3), 2)) Then
                        Combo1(pcmbG_SHIIRE_KBN).ListIndex = i
                        Exit For
                    End If
                
                Next i
        
                If i = -1 Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            
            End If
        
        Case ptxTANTO_CODE  '�S����
            If Not NOUKI_MODE Then
                Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTANTO_CODE).Text)
                
                sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
                Select Case sts
                    Case BtNoErr
                    
                        Text1(ptxTANTO_NAME).Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                    Case BtErrKeyNotFound
                        Text1(ptxTANTO_NAME).Text = ""
                        MsgBox "���͂������ڂ̓G���[�ł��B"
                        Text1(Mode).SetFocus
                        Exit Function
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
                        Exit Function
                
                End Select
            
            
            
            End If
        
        
        Case ptxORDER_CODE      '������
            If Not NOUKI_MODE Then
        
                Combo1(pcmbORDER).ListIndex = -1
                For i = 0 To Combo1(pcmbORDER).ListCount - 1
                
                    If Text1(ptxORDER_CODE).Text = Trim(Right(Combo1(pcmbORDER).List(i), 5)) Then
                        Combo1(pcmbORDER).ListIndex = i
                        Exit For
                    End If
                
                Next i
        
                If i = -1 Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            
            End If
        
        
        Case ptxY_NOUKI_DT  '�[���\���
        
            If NOUKI_MODE Then
            
                If Not IsDate(Text1(ptxY_NOUKI_DT).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxY_NOUKI_DT).Text = Format(CDate(Text1(ptxY_NOUKI_DT).Text), "YYYY/MM/DD")
                End If
            
            End If
        
        Case ptxUKEIRE_DT   '�����
            
            If Not NOUKI_MODE Then
            
                If Not IsDate(Text1(ptxUKEIRE_DT).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxUKEIRE_DT).Text = Format(CDate(Text1(ptxUKEIRE_DT).Text), "YYYY/MM/DD")
                    
                    If Input_Mode Then
                    
                        '������
                        Text1(ptxORDER_DT).Text = Format(CDate(Text1(ptxUKEIRE_DT).Text), "YYYY/MM/DD")
                        '�[���\���
                        Text1(ptxY_NOUKI_DT).Text = Format(CDate(Text1(ptxUKEIRE_DT).Text), "YYYY/MM/DD")
                    
                    End If
                
                
                End If
        
            End If
        
        Case ptxKEIJYO_YM       '�����N��
            
            If Not NOUKI_MODE Then
            
                wkDate = Text1(ptxKEIJYO_YM).Text & "/01"
                
                If Not IsDate(wkDate) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    
                    wkDate = Format(CDate(Text1(ptxKEIJYO_YM).Text), "YYYY/MM/DD")
                    
                    Text1(ptxKEIJYO_YM).Text = Mid(wkDate, 1, 7)
                End If
            End If
        
        Case ptxKONKAI_UKEIRE_QTY   '�����
    
            If Not NOUKI_MODE Then
            
                If Not IsNumeric(Text1(ptxKONKAI_UKEIRE_QTY).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxKONKAI_UKEIRE_QTY).Text = Format(CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text), "#0")
                    
                                        
                    If Input_Mode Then
                        '������
                        Text1(ptxORDER_QTY).Text = Format(CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text), "#0")
                        '����ϐ�
                        Text1(ptxUKEIRE_QTY).Text = "0"
                    
                    End If
                    
                    
''                    If CLng(Text1(ptxORDER_QTY).Text) - CLng(Text1(ptxUKEIRE_QTY).Text) < CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text) Then
''                        MsgBox "���͂������ڂ̓G���[�ł��B"
''                       Text1(Mode).SetFocus
''                        Exit Function
''                    End If
                    
                    
                    
                    If CLng(CLng(Text1(ptxORDER_QTY).Text) - CLng(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)) - CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)) < 0 Then
                        Text1(ptxZAN_QTY).Text = "0"
                    Else
                        
                        If Save_UKEIRE_QTY = CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text) Then
                        Else
                        
                            Text1(ptxZAN_QTY).Text = Format(CLng(Text1(ptxORDER_QTY).Text) - CLng(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)) - CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text), "#0")
                            Save_UKEIRE_QTY = CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)
                        End If
                    End If
                    Text1(ptxKINGAKU).Text = Format(CDbl(Text1(ptxTANKA).Text) * CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text), "#,##0")
                        
                
                    If CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text) <= 0 Then
                                        
                        Check1(chkZAIKO_F).Value = vbChecked
                    End If
                End If
    
    
    
            End If
    
        Case ptxZAN_QTY         '�����c
    
            If Not NOUKI_MODE Then
    
    
                If Not IsNumeric(Text1(ptxZAN_QTY).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    Text1(ptxZAN_QTY).Text = Format(CLng(Text1(ptxZAN_QTY).Text), "#0")
                    '�����c
                    If (CLng(Text1(ptxORDER_QTY).Text) - CLng(Text1(ptxUKEIRE_QTY).Text) - CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text)) = CLng(Text1(ptxZAN_QTY).Text) Or _
                        CLng(Text1(ptxZAN_QTY).Text) = 0 Then
                    Else
                        MsgBox "���͂������ڂ̓G���[�ł��B"
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
                
                        
                End If
            
            End If
    
    
        Case ptxTANKA       '�P��
    
            If Not NOUKI_MODE Then
    
    
                If Not IsNumeric(Text1(ptxTANKA).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    '�P��
                    Text1(ptxTANKA).Text = Format(CDbl(Text1(ptxTANKA).Text), "#0.00")
                    '���z�v�Z
                    If IsNumeric(Text1(ptxORDER_QTY).Text) Then
                        Text1(ptxKINGAKU).Text = Format(CLng(CDbl(Text1(ptxTANKA).Text) * CDbl(Text1(ptxKONKAI_UKEIRE_QTY).Text)), "#0")
                    End If
                        
                        
                End If
            
            End If
    
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function
Private Function Item_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   ��ʕ\��
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim i           As Integer

Dim SUMI_QTY    As Long
Dim MI_QTY      As Long

    Item_Disp_Proc = True
    
    Call Input_Area_Proc(0)
    
    
    
        
    Text1(ptxORDER_NO).Text = StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode)        '������
                                                                                '������
    Text1(ptxORDER_DT).Text = Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 7, 2)
        
    Text1(ptxHIN_GAI).Text = StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)          '�i��
        
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    
    End Select
    Text1(ptxHIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    If Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                            StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
        Exit Function
    
    End If
    Text1(ptxZAIKO_QTY).Text = Format(SUMI_QTY + MI_QTY, "#0")
        
        
    Text1(ptxG_SHIIRE_KBN).Text = StrConv(P_SHORDER_REC.G_SHIIRE_KBN, vbUnicode)   '�d���敪
    Combo1(pcmbG_SHIIRE_KBN).ListIndex = -1
    For i = 0 To Combo1(pcmbG_SHIIRE_KBN).ListCount - 1
    
        If Text1(ptxG_SHIIRE_KBN).Text = Trim(Left(Right(Combo1(pcmbG_SHIIRE_KBN).List(i), 3), 2)) Then
            Combo1(pcmbG_SHIIRE_KBN).ListIndex = i
            Exit For
        End If
    
    Next i
        
        
        
    Text1(ptxTANTO_CODE).Text = StrConv(P_SHORDER_REC.TANTO_CODE, vbUnicode)       '�S���Һ��ށ^����
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTANTO_CODE).Text)

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            Text1(ptxTANTO_NAME).Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Text1(ptxTANTO_NAME).Text = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
            Exit Function
    
    End Select
                                                                                    '������
    Text1(ptxORDER_CODE).Text = Trim(StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode))
    Combo1(pcmbORDER).ListIndex = -1
    For i = 0 To Combo1(pcmbORDER).ListCount - 1
    
        If Text1(ptxORDER_CODE).Text = Trim(Right(Combo1(pcmbORDER).List(i), 5)) Then
            Combo1(pcmbORDER).ListIndex = i
            Exit For
        End If
    
    Next i
                                                                                    '�[����
    Text1(ptxDELI_CODE).Text = Trim(StrConv(P_SHORDER_REC.DELI_CODE, vbUnicode))
    Combo1(pcmbDELI).ListIndex = -1
    For i = 0 To Combo1(pcmbDELI).ListCount - 1
    
        If Text1(ptxDELI_CODE).Text = Trim(Right(Combo1(pcmbDELI).List(i), 5)) Then
            Combo1(pcmbDELI).ListIndex = i
            Exit For
        End If
    
    Next i
                                                                                    
                                                                                    '�[���\���
    Text1(ptxY_NOUKI_DT).Text = Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 7, 2)
                                                                                    
                                                                                    
                                                                                    '������
    Text1(ptxORDER_QTY).Text = Format(CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)), "#0")
                                                                                    '����ϐ�
    Text1(ptxUKEIRE_QTY).Text = Format(CLng(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)), "#0")
                                                                                    '�P��
    Text1(ptxTANKA).Text = Format(CDbl(StrConv(P_SHORDER_REC.TANKA, vbUnicode)), "#0.00")
                                                                                    '����ۯ�
    Text1(ptxLOT).Text = Format(CLng(StrConv(P_SHORDER_REC.LOT, vbUnicode)), "#0")
                                                                                    '��������
    Text1(ptxKONKAI_UKEIRE_QTY).Text = Format(CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) - _
                                        CLng(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)), "#0")
                                                                                    '�����c
    Text1(ptxZAN_QTY).Text = "0"
                                                                                    '���z
    Text1(ptxKINGAKU).Text = Format(CDbl(StrConv(P_SHORDER_REC.TANKA, vbUnicode)) * _
                                    CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text), "#,##0")
    
                                                                                    '�݌Ɍv��L��
    If StrConv(ITEMREC.ZAIKO_F, vbUnicode) <> P_ZAIKO_F_OFF Then
        Check1(chkZAIKO_F).Value = vbChecked
    Else
        Check1(chkZAIKO_F).Value = vbUnchecked
    End If
    Item_Disp_Proc = False

End Function

Private Function Cancel_Proc() As Integer
'----------------------------------------------------------------------------
'                  ���ޒ����ް���ݾٍX�V
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim ans         As Integer
Dim com         As Integer

Dim SEQNO       As Integer



Dim i           As Integer


    Cancel_Proc = True
                                        
                                        
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If

    
    
    
    '---------------------------------------------------    '���ޒ����ް���ݾ�
    
    Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, Text1(ptxORDER_NO).Text)
    
    Do
    
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
            
        Select Case sts
            Case BtNoErr
            
                
                Exit Do
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B< P_SHORDER.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            
            
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���ޒ����ް�")
                GoTo Abort_Tran
        End Select

    Loop
    
    
    Call UniCode_Conv(P_SHORDER_REC.CANCEL_F, P_CANCEL_ON)  '��ݾ��׸�
                                                            '��ݾٓ���
    Call UniCode_Conv(P_SHORDER_REC.CANCEL_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                                                                                
                                                            '�X�V����
    Call UniCode_Conv(P_SHORDER_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpUpdate, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
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
                Call File_Error(sts, BtOpUpdate, "���ޒ����ް�")
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
    
    Cancel_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    
    Call Input_UnLock

End Function
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                  ���ޒ����ް��X�V
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer

Dim SEQNO           As Integer


    Update_Proc = True
                                        
    Call Input_Lock
                                        
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
    '---------------------------------------------------    '���ޒ����f�[�^�X�V
    If Input_Mode Then
    
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
                        Update_Proc = True
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�Ǘ��}�X�^")
                    GoTo Abort_Tran
            
            End Select
        
        
        Loop
    
        '���������{�P
    
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
                                                                                '������
        Call UniCode_Conv(P_SHORDER_REC.ORDER_NO, StrConv(P_KANRIREC.ORDER_NO, vbUnicode))
                                                                                '������
        Call UniCode_Conv(P_SHORDER_REC.ORDER_DT, Format(Text1(ptxORDER_DT).Text, "YYYYMMDD"))
        
        Call UniCode_Conv(P_SHORDER_REC.PRINT_DATETIME, "")                     '���s����
        Call UniCode_Conv(P_SHORDER_REC.TANTO_CODE, Text1(ptxTANTO_CODE).Text)  '�S����
    
    
        Call UniCode_Conv(P_SHORDER_REC.JGYOBU, SHIZAI)                         '���ƕ��i�����ށj
        Call UniCode_Conv(P_SHORDER_REC.NAIGAI, NAIGAI_NAI)                     '�����O
        Call UniCode_Conv(P_SHORDER_REC.HIN_GAI, Text1(ptxHIN_GAI).Text)        '�i��
        Call UniCode_Conv(P_SHORDER_REC.ORDER_CODE, Text1(ptxORDER_CODE).Text)  '�����溰��
        Call UniCode_Conv(P_SHORDER_REC.DELI_CODE, "")                          '�[���溰��
        Call UniCode_Conv(P_SHORDER_REC.ORDER_QTY, Format(CDbl(Text1(ptxORDER_QTY).Text), _
                                                                "00000000.00")) '������
        Call UniCode_Conv(P_SHORDER_REC.Y_NOUKI_DT, Format(CDate(Text1(ptxY_NOUKI_DT).Text), _
                                                                "YYYYMMDD"))    '�\��[��
        Call UniCode_Conv(P_SHORDER_REC.TANKA, Format(CDbl(Text1(ptxTANKA).Text), _
                                                                "00000000.00")) '�P��
        Call UniCode_Conv(P_SHORDER_REC.LOT, "00000001")
    
        Call UniCode_Conv(P_SHORDER_REC.KAN_F, P_KAN_ON)                        '�����׸ށi�������j
        Call UniCode_Conv(P_SHORDER_REC.KAN_DT, Format(Now, "YYYYMMDD"))        '������
        Call UniCode_Conv(P_SHORDER_REC.BUNNOU_CNT, "01")                       '�����
        Call UniCode_Conv(P_SHORDER_REC.UKEIRE_QTY, Format(CDbl(Text1(ptxKONKAI_UKEIRE_QTY).Text), "00000000.00"))
    
        Call UniCode_Conv(P_SHORDER_REC.CANCEL_F, P_CANCEL_OFF)                 '��ݾ��׸�
        Call UniCode_Conv(P_SHORDER_REC.CANCEL_DATETIME, "")                    '��ݾٓ���
    
        Call UniCode_Conv(P_SHORDER_REC.PRINT_F, P_PRINT_ON)                    '����׸�(����ςƂ���)
    
        Call UniCode_Conv(P_SHORDER_REC.WS_NO, WS_NO)                           '���͒[��
                                                                                '�d���敪
        Call UniCode_Conv(P_SHORDER_REC.G_SHIIRE_KBN, Text1(ptxG_SHIIRE_KBN).Text)
        
        '�i��Ͻ��Ǎ���
        Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                MsgBox "�i�ڃ}�X�^�����[���ŕύX����܂����B�X�V�����𒆎~���܂��B"
                GoTo Abort_Tran
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���ޒ����ް�")
                GoTo Abort_Tran
        End Select
        '���x�P��
        Call UniCode_Conv(P_SHORDER_REC.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))
        
        
        '�󕥐�Ͻ��Ǎ���
        Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxORDER_CODE).Text)
        sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
            
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                MsgBox "�󕥐�}�X�^�����[���ŕύX����܂����B�X�V�����𒆎~���܂��B"
                GoTo Abort_Tran
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�󕥐�Ͻ�")
                GoTo Abort_Tran
        End Select
    
                                                                                    '�����敪
        Call UniCode_Conv(P_SHORDER_REC.TORI_KBN, StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode))
        
        
        Call UniCode_Conv(P_SHORDER_REC.FILLER, "")
    
    
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
    
    
    
    Else
    
        '���ޒ����f�[�^����
        Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, Text1(ptxORDER_NO).Text)
        
        Do
        
            sts = BTRV(BtOpGetEqual + BtSNoWait, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
                
            Select Case sts
                Case BtNoErr
                
                    
                    Exit Do
                
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SHORDER.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        GoTo Abort_Tran
                    End If
                
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���ޒ����f�[�^")
                    GoTo Abort_Tran
            End Select
    
        Loop
        
        If CDbl(Text1(ptxZAN_QTY).Text) = 0 Then
            Call UniCode_Conv(P_SHORDER_REC.KAN_F, P_KAN_ON)                   '�����׸�
            Call UniCode_Conv(P_SHORDER_REC.KAN_DT, Format(Now, "YYYYMMDD"))   '������
            If CInt(StrConv(P_SHORDER_REC.BUNNOU_CNT, vbUnicode)) = 0 Then     '���[��
            Else
                Call UniCode_Conv(P_SHORDER_REC.BUNNOU_CNT, Format(CInt(CInt(StrConv(P_SHORDER_REC.BUNNOU_CNT, vbUnicode)) + 1), "000"))
            End If
        End If
        
        Call UniCode_Conv(P_SHORDER_REC.UKEIRE_QTY, Format(CDbl(CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)) + CDbl(Text1(ptxKONKAI_UKEIRE_QTY).Text)), "00000000.00"))
                                                            '�X�V����
        Call UniCode_Conv(P_SHORDER_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
        Do
            
            DoEvents
            
            sts = BTRV(BtOpUpdate, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
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
                    Call File_Error(sts, BtOpUpdate, "���ޒ����ް�")
                    GoTo Abort_Tran
            End Select
        
        Loop
    End If
    
    SEQNO = 0
    
    
    
    '���ގ�������ް�����
    Call UniCode_Conv(K0_P_SHUKEIRE.ORDER_NO, StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode))
    Call UniCode_Conv(K0_P_SHUKEIRE.SEQNO, "")
    
    com = BtOpGetGreater
    
    Do
    
        sts = BTRV(com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K0_P_SHUKEIRE, Len(K0_P_SHUKEIRE), 0)
            
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode) <> StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "���ގ������")
                GoTo Abort_Tran
        End Select
        
        
        
        SEQNO = SEQNO + 1
        
        
        com = BtOpGetNext
        
    Loop
        
                                                                                '������
    Call UniCode_Conv(P_SHUKEIRE_REC.ORDER_NO, StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode))                                                                                         '�����
                                                                                '������
    Call UniCode_Conv(P_SHUKEIRE_REC.ORDER_CODE, StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode))
                                                                                '�����
    Call UniCode_Conv(P_SHUKEIRE_REC.UKEIRE_DT, Format(Text1(ptxUKEIRE_DT).Text, "YYYYMMDD"))
                                                                                '�������
    Call UniCode_Conv(P_SHUKEIRE_REC.UKEIRE_QTY, Format(CDbl(Text1(ptxKONKAI_UKEIRE_QTY).Text), "00000000.00"))
                                                                                '����P��
    Call UniCode_Conv(P_SHUKEIRE_REC.UKEIRE_TANKA, StrConv(P_SHORDER_REC.TANKA, vbUnicode))
                                                                                '������z
    Call UniCode_Conv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, Format(CLng(CDbl(Text1(ptxKONKAI_UKEIRE_QTY).Text) * _
                                                        CDbl(StrConv(P_SHORDER_REC.TANKA, vbUnicode))), "00000000"))
        
        
    If CDbl(Text1(ptxZAN_QTY).Text) = 0 Then
        Call UniCode_Conv(P_SHUKEIRE_REC.LAST_F, P_UKEIRE_END)
    Else
        Call UniCode_Conv(P_SHUKEIRE_REC.LAST_F, P_UKEIRE_CON)
    End If
                                                                                '�v��N��
    Call UniCode_Conv(P_SHUKEIRE_REC.KEIJYO_YM, Mid(Text1(ptxKEIJYO_YM), 1, 4) & Mid(Text1(ptxKEIJYO_YM), 6, 2))
        
    Call UniCode_Conv(P_SHUKEIRE_REC.FILLER, "")
                                                        '�X�V����
    Call UniCode_Conv(P_SHUKEIRE_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    Do
            
        SEQNO = SEQNO + 1
                                                '�ǔ�
        Call UniCode_Conv(P_SHUKEIRE_REC.SEQNO, Format(SEQNO, "000"))
            
        DoEvents
            
        sts = BTRV(BtOpInsert, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K0_P_SHUKEIRE, Len(K0_P_SHUKEIRE), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrDuplicates
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SHUKEIRE.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "���ގ������")
                GoTo Abort_Tran
        End Select
        
    Loop
    '------------------------------------------------ POS���׏���
    If POS_UMU Then
        If Check1(chkZAIKO_F).Value = vbChecked Then
    
            If POS_NYUKA_Update_Proc("  ", "  ", "  ", "  ") Then
                GoTo Abort_Tran
            End If
        End If
    
    Else
        
        If Check1(chkZAIKO_F).Value = vbChecked Then
    
            'POS�Ȃ��͕W���I�Ԃɍ݌Ɍv��    2006.04.24
    
    
            If POS_NYUKA_Update_Proc(StrConv(ITEMREC.ST_SOKO, vbUnicode), _
                                        StrConv(ITEMREC.ST_RETU, vbUnicode), _
                                        StrConv(ITEMREC.ST_REN, vbUnicode), _
                                        StrConv(ITEMREC.ST_DAN, vbUnicode)) Then
                GoTo Abort_Tran
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
Private Function NOUKI_Update_Proc() As Integer
'----------------------------------------------------------------------------
'                  ���ޒ����ް��X�V
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer




    NOUKI_Update_Proc = True
                                        
    Call Input_Lock
                                        
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
    '---------------------------------------------------    '���ޒ����f�[�^�X�V
    '���ޒ����f�[�^����
    Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, Text1(ptxORDER_NO).Text)
    
    Do
    
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
            
        Select Case sts
            Case BtNoErr
            
                
                Exit Do
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SHORDER.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            
            
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���ޒ����f�[�^")
                GoTo Abort_Tran
        End Select

    Loop
    
                                                        '�\��̕���
    Call UniCode_Conv(P_SHORDER_REC.Y_NOUKI_DT, Format(Text1(ptxY_NOUKI_DT).Text, "YYYYMMDD"))
                                                        '�X�V����
    Call UniCode_Conv(P_SHORDER_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpUpdate, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SHORDER.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
                    If sts Then
                        Call File_Error(sts, BtOpUnlock, "���ޒ����ް�")
                    End If
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "���ޒ����ް�")
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
    
    NOUKI_Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    
    Call Input_UnLock

End Function


Private Sub Check2_Click()

End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    Select Case Index
        Case pcmbG_SHIIRE_KBN   '�d���敪
            Text1(ptxG_SHIIRE_KBN).Text = Trim(Left(Right(Combo1(pcmbG_SHIIRE_KBN).Text, 3), 2))
        Case pcmbORDER          '������
            Text1(ptxORDER_CODE).Text = Trim(Right(Combo1(pcmbORDER).Text, 5))
        Case pcmbDELI           '�[����
            Text1(ptxDELI_CODE).Text = Trim(Right(Combo1(pcmbDELI).Text, 5))
    End Select
    
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    
    Select Case Index
        Case pcmbG_SHIIRE_KBN   '�d���敪
            Text1(ptxG_SHIIRE_KBN).Text = Trim(Left(Right(Combo1(pcmbG_SHIIRE_KBN).Text, 3), 2))
        Case pcmbORDER          '������
            Text1(ptxORDER_CODE).Text = Trim(Right(Combo1(pcmbORDER).Text, 5))
        Case pcmbDELI           '�[����
            Text1(ptxDELI_CODE).Text = Trim(Right(Combo1(pcmbDELI).Text, 5))
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans         As Integer
Dim i           As Integer


Dim sts         As Integer

    Select Case Index
        Case P_CMD_Upd        '�X�V
            
            
            For i = ptxORDER_NO To ptxKINGAKU
            
                If Error_Check_Proc(i) Then     '�G���[�`�F�b�N
                    Exit Sub
                End If
            
            Next i
            
            
            If NOUKI_MODE Then
            
                ans = MsgBox("�[���ύX���܂����H", vbYesNo + vbQuestion, "�m�F����")
                If ans = vbYes Then
                    If NOUKI_Update_Proc() Then
                        Unload Me
                    End If
                    
                    If List_Disp_Proc() Then
                        Unload Me
                    End If
                    
                    If Init_Proc() Then
                        Unload Me
                    End If
                
                    Text1(ptxORDER_NO).SetFocus
                
                Else
                
                    Text1(ptxY_NOUKI_DT).SetFocus
                
                End If
            
            
            Else
                ans = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
                If ans = vbYes Then
                    If Update_Proc() Then
                        Unload Me
                    End If
                    
                    If List_Disp_Proc() Then
                        Unload Me
                    End If
                    
                    If Init_Proc() Then
                        Unload Me
                    End If
                
                    Text1(ptxORDER_NO).SetFocus
                
                Else
                    Text1(ptxUKEIRE_DT).SetFocus
                End If
            
            
            End If
            
        Case P_CMD_DEL                      '�폜
        
        
            '���ޒ����ް�
            Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, Text1(ptxORDER_NO))
            sts = BTRV(BtOpGetEqual, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
                
            Select Case sts
                Case BtNoErr
                                   
                    If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = P_KAN_ON Then
                        MsgBox "�����o�^�ςł��B"
                        Text1(ptxORDER_NO).SetFocus
                        Exit Sub
                    End If
                
                    If CLng(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)) <> 0 Then
                        MsgBox "�d�����т��L��܂��B"
                        Text1(ptxORDER_NO).SetFocus
                        Exit Sub
                    End If
                
                    If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
                        MsgBox "�L�����Z���ςł��B"
                        Text1(ptxORDER_NO).SetFocus
                        Exit Sub
                    End If
                
                Case BtErrKeyNotFound
                    MsgBox "���ޒ����ް����o�^�ł��B"
                    Text1(ptxORDER_NO).SetFocus
                    Exit Sub
                Case Else
                    Unload Me
            End Select
        
        
        
            ans = MsgBox("��ݾق��܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Cancel_Proc() Then
                    Unload Me
                End If
                
                If Init_Proc() Then
                    Unload Me
                End If
            
                If List_Disp_Proc() Then
                    Unload Me
                End If
            
            
            End If
            
            Text1(ptxTANTO_CODE).SetFocus
    
        Case P_CMD_DSP                      '����/�\��
            
            If List_Disp_Proc() Then
                Unload Me
            End If
        
        
        Case cmdNOUKI
        
            If NOUKI_MODE Then
                Call Input_Area_Set(0)
                NOUKI_MODE = False
                Text1(ptxUKEIRE_DT).SetFocus
            Else
                Call Input_Area_Set(1)
                NOUKI_MODE = True
                Text1(ptxY_NOUKI_DT).SetFocus
            End If
        
        Case P_CMD_OUT                      '�ް��o��
        
        Case P_CMD_PRT                      '���
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
Dim sBuffer As String

    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If

    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        WS_NO = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        WS_NO = "???"
    End If

                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    End If
    LOG_F = RTrim(c)
                                'POS���їL���̎�荞��
    If GetIni(StrConv(App.EXEName, vbUpperCase), "POS_UMU", "P_SYS", c) Then
        POS_UMU = False
    Else
        If RTrim(c) = "0" Then
            POS_UMU = False
        Else
            POS_UMU = True
        End If
    End If
'''     POS�Ȃ��ł��݌Ɍv�シ��2006.04.24
'''    If POS_UMU Then
                                '���׉��z�q�ɂ̎�荞��
        If GetIni(StrConv(App.EXEName, vbUpperCase), "NYUKA_SOKO", "P_SYS", c) Then
            Beep
            MsgBox "���׉��z�q�ɔԍ��̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
            End
        End If
        KASO_NYUKA = RTrim(c)
    
    
                                '�u���ޒʏ���ׁv�̗v��
        If GetIni(StrConv(App.EXEName, vbUpperCase), "YOIN_TU_NYUKA", "P_SYS", c) Then
            Call Log_Out(LOG_F, "[P_SYS.INI][" & StrConv(App.EXEName, vbUpperCase) & "[YOIN_TU_NYUKA] READ ERROR")
            MsgBox "���ޒʏ���חp�v���̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
            End
        End If
        P_YOIN_TU_NYUKA = Trim(c)
                                '�u���ޑO�ؑ��E�v�̗v��
        If GetIni(StrConv(App.EXEName, vbUpperCase), "YOIN_MAE_SOUSAI", "P_SYS", c) Then
            Call Log_Out(LOG_F, "[P_SYS.INI][" & StrConv(App.EXEName, vbUpperCase) & "[YOIN_MAE_SOUSAI] READ ERROR")
            MsgBox "���ޑO�ؑ��E�p�v���̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
            End
        End If
        P_YOIN_MAE_SOUSAI = Trim(c)
    
    
                                    '����������荞��
        If GetIni(App.EXEName, "MEMO", "P_SYS", c) Then
            MEMO_TEXT = ""
        Else
            MEMO_TEXT = RTrim(c)
        End If
    
    
'''    End If
                                '�I�}�X�^�n�o�d�m
    If TANA_Open(BtOpenNomal) Then
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
                                '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�󕥐�}�X�^�n�o�d�m
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ޒ����ް��n�o�d�m
    If P_SHORDER_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ގ�������ް��n�o�d�m
    If P_SHUKEIRE_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '�݌��ް��n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '����Ͻ��n�o�d�m
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ޑO���ް�
    If P_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    '---------------------------    POS�ݸ�p̧��
                                '�i�ڃ}�X�^�n�o�d�m�i�f�[�^�X�V�p�j
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�v���}�X�^�n�o�d�m
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ԃ}�X�^�n�o�d�m
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ח\��f�[�^�t�@�C���n�o�d�m
    If Y_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ɉړ����n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
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
    
    '�d���敪�̃Z�b�g
    If Code_Set_Proc(pcmbG_SHIIRE_KBN, P_KBN01_CD, 0) Then
        Unload Me
    End If
    
    
    '������
    If Ukeharai_Set_Proc(pcmbORDER) Then
        Unload Me
    End If
    '�[����
    If Ukeharai_Set_Proc(pcmbDELI) Then
        Unload Me
    End If
    
    
    '��ʏ����ݒ�
    If Init_Proc() Then
        Unload Me
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
                                            
                                            
                                            
                                            '�I�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�I�}�X�^")
        End If
    End If
                                            
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
                                            '�S���҃}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�S���҃}�X�^")
        End If
    End If
                                            '�󕥐�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�󕥐�}�X�^")
        End If
    End If
                                            '���ޒ����ް��b�k�n�r�d
    sts = BTRV(BtOpClose, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ޒ����ް�")
        End If
    End If
                                            '���ގ�������ް��b�k�n�r�d
    sts = BTRV(BtOpClose, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K0_P_SHUKEIRE, Len(K0_P_SHUKEIRE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ގ�������ް�")
        End If
    End If
                                            '�݌��ް��b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌��ް�")
        End If
    End If
                                            '����Ͻ��b�k�n�r�d
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "����Ͻ�")
        End If
    End If
    '-------------------------------------- POS�ݸ���
                                            '�i�ڃ}�X�^�i�f�[�^�X�V�p�j�b�k�n�r�d
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�v���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�v���}�X�^")
        End If
    End If
                                            '���ԃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ԃ}�X�^")
        End If
    End If
                                            '���ח\��f�[�^�t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ח\��f�[�^�t�@�C��")
        End If
    End If
                                            '�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, Y_NYUREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PI000401 = Nothing

    End
End Sub





Private Sub TDBGrid1_DblClick()
Dim sts As Integer
    
    Text1(ptxORDER_NO).Text = SHORDER(TDBGrid1.Bookmark, colORDER_NO)
    '���ޒ����f�[�^�̃`�F�b�N
    sts = P_SHORDER_Read_Proc()
    Select Case sts
        Case False, BtNoErr
                    
            If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = P_PRINT_ON Then
                MsgBox "���[���ŏ����������Ă��܂��B"
                TDBGrid1.SetFocus
                Exit Sub
            End If
            Save_UKEIRE_QTY = 0
        
        Case BtErrKeyNotFound
            MsgBox "���[���ŏ����������Ă��܂��B"
            TDBGrid1.SetFocus
            Exit Sub
        Case Else
            Exit Sub
    End Select
    
    Text1(ptxUKEIRE_DT).SetFocus
    

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
                    
        SHORDER.QuickSort Min_Row, SHORDER.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
        
        Set TDBGrid1.Array = SHORDER
        
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

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    If Error_Check_Proc(Index) Then     '�G���[�`�F�b�N
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
    
    
    
    For i = ptxORDER_NO To ptxKINGAKU
        Text1(i).Text = ""
    Next i
    '�����������
    Text1(ptxUKEIRE_DT).Text = Format(Now, "YYYY/MM/DD")
    '�v�㌎������
    Text1(ptxKEIJYO_YM).Text = Left(Format(Now, "YYYY/MM/DD"), 7)


    Combo1(pcmbG_SHIIRE_KBN).ListIndex = 0


    For i = pcmbORDER To pcmbDELI
        
        Combo1(i).ListIndex = -1
    
    Next i


    Check1(chkZAIKO_F).Value = vbUnchecked

    If List_Disp_Proc() Then
        Exit Function
    End If

    '��ď��̏�����
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0             '��̫�ď���
    Next i

    Sort_Tbl(colHIN_NAME) = 9       '��ď��O


    NOUKI_MODE = False
    Call Input_Area_Set(0)

    Call UniCode_Conv(ITEMREC.JGYOBU, "")
    Call UniCode_Conv(ITEMREC.NAIGAI, "")
    Call UniCode_Conv(ITEMREC.HIN_GAI, "")
    Save_UKEIRE_QTY = 0
    

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
Private Function P_SHORDER_Read_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���ޒ����f�[�^�̓ǂݍ���
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
    
    
    P_SHORDER_Read_Proc = True
    
    
    '���ޒ����ް�
    Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, Text1(ptxORDER_NO))
    sts = BTRV(BtOpGetEqual, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
        
    Select Case sts
        Case BtNoErr
        
        
        
        
        Case Else
            P_SHORDER_Read_Proc = sts
            Exit Function
    
    End Select
    
    
    If Item_Disp_Proc() Then
        Exit Function
    End If
    
    P_SHORDER_Read_Proc = False
        
    

End Function
Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'           ���ޒ����ް��̕\��
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim com     As Integer

Dim Row     As Long

    List_Disp_Proc = True
    PI000401.MousePointer = vbHourglass
    
    
    
    Set SHORDER = Nothing
    Tbl_Set_F = False
    
    
    
    com = BtOpGetFirst
    
    Row = Min_Row - 1
       
    Do
    
        DoEvents
    
        sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K2_P_SHORDER, Len(K2_P_SHORDER), 2)
            
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���ޒ����ް�")
                Exit Function
        End Select
    
    
        If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Or _
            StrConv(P_SHORDER_REC.KAN_F, vbUnicode) = P_KAN_ON Then
        Else
    
            Row = Row + 1
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
            Tbl_Set_F = True
        End If
        
        com = BtOpGetNext
    
    Loop
    
    Set TDBGrid1.Array = SHORDER
            
    If Row <> (Min_Row - 1) Then
        SHORDER.QuickSort Min_Row, SHORDER.UpperBound(1), colORDER_NO, XORDER_ASCEND, XTYPE_STRING
    End If
            
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    
    
    PI000401.MousePointer = vbDefault
    List_Disp_Proc = False
    


End Function

Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           ���ޒ����ް��̓��e���د�ނɾ�Ă���
'----------------------------------------------------------------------------
Dim sts As Integer

Dim SUMI_QTY    As Long
Dim MI_QTY      As Long





    Grid_Set_Proc = True
    
    SHORDER.ReDim Min_Row, Row, Min_Col, Max_Col


    '������
    SHORDER(Row, colORDER_DT) = Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.ORDER_DT, vbUnicode), 7, 2)
    '������
    SHORDER(Row, colORDER_NO) = StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode)
    '������
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�󕥐�}�X�^")
            Exit Function
    End Select
    '������
    SHORDER(Row, colORDER_NAME) = StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode) & " " & _
                                StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode)
    '�i��
    SHORDER(Row, colHIN_GAI) = StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)
    '�i��
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    End Select
    '�i��
    SHORDER(Row, colHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    '������
    SHORDER(Row, colORDER_QTY) = Format(CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)), "#,##0")
    '�����c
    SHORDER(Row, colZAN_QTY) = Format(CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) - _
                                        CLng(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)), "#,##0")
    '�݌Ɏc
    If Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                            StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
        Exit Function
    
    End If
    SHORDER(Row, colZAIKO_QTY) = Format(SUMI_QTY + MI_QTY, "#,##0")
    '�[���\���
    SHORDER(Row, colY_NOUKI_DT) = Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 7, 2)

    Grid_Set_Proc = False

End Function

Private Sub Input_Area_Set(Mode As Integer)
'----------------------------------------------------------------------------
'           ���̓G���A�̐؂�ւ�
'----------------------------------------------------------------------------
                
                
    Select Case Mode
        Case 0  '�[��--���ʏ�
                
            Text1(ptxY_NOUKI_DT).BackColor = G_INPUT_NG
            Text1(ptxY_NOUKI_DT).Locked = True
            Text1(ptxY_NOUKI_DT).TabStop = False

            Text1(ptxUKEIRE_DT).BackColor = G_INPUT_OK
            Text1(ptxUKEIRE_DT).Locked = False
            Text1(ptxUKEIRE_DT).TabStop = True

            Text1(ptxKEIJYO_YM).BackColor = G_INPUT_OK
            Text1(ptxKEIJYO_YM).Locked = False
            Text1(ptxKEIJYO_YM).TabStop = True

            Text1(ptxKONKAI_UKEIRE_QTY).BackColor = G_INPUT_OK
            Text1(ptxKONKAI_UKEIRE_QTY).Locked = False
            Text1(ptxKONKAI_UKEIRE_QTY).TabStop = True

            Text1(ptxZAN_QTY).BackColor = G_INPUT_OK
            Text1(ptxZAN_QTY).Locked = False
            Text1(ptxZAN_QTY).TabStop = True

        Case 1  '�ʏ�--���[��
                
            Text1(ptxY_NOUKI_DT).BackColor = G_INPUT_OK
            Text1(ptxY_NOUKI_DT).Locked = False
            Text1(ptxY_NOUKI_DT).TabStop = True

            Text1(ptxUKEIRE_DT).BackColor = G_INPUT_NG
            Text1(ptxUKEIRE_DT).Locked = True
            Text1(ptxUKEIRE_DT).TabStop = False

            Text1(ptxKEIJYO_YM).BackColor = G_INPUT_NG
            Text1(ptxKEIJYO_YM).Locked = True
            Text1(ptxKEIJYO_YM).TabStop = False

            Text1(ptxKONKAI_UKEIRE_QTY).BackColor = G_INPUT_NG
            Text1(ptxKONKAI_UKEIRE_QTY).Locked = True
            Text1(ptxKONKAI_UKEIRE_QTY).TabStop = False

            Text1(ptxZAN_QTY).BackColor = G_INPUT_NG
            Text1(ptxZAN_QTY).Locked = True
            Text1(ptxZAN_QTY).TabStop = False

    End Select


End Sub
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
Private Sub Input_Area_Proc(Mode As Integer)
'----------------------------------------------------------------------------
'                   ���͉\�̈�̐؂�ւ�
'----------------------------------------------------------------------------
    
    
    Select Case Mode
        Case 0      '�m�[�}��
    
            Input_Mode = False
    
            '�i��
            Text1(ptxHIN_GAI).BackColor = G_INPUT_NG
            Text1(ptxHIN_GAI).Locked = True
            Text1(ptxHIN_GAI).TabStop = False
    
            '�S����
            Text1(ptxTANTO_CODE).BackColor = G_INPUT_NG
            Text1(ptxTANTO_CODE).Locked = True
            Text1(ptxTANTO_CODE).TabStop = False
            '������
            Text1(ptxORDER_CODE).BackColor = G_INPUT_NG
            Text1(ptxORDER_CODE).Locked = True
            Text1(ptxORDER_CODE).TabStop = False
            
            Combo1(pcmbORDER).BackColor = G_INPUT_NG
            Combo1(pcmbORDER).Locked = True
            Combo1(pcmbORDER).TabStop = False
            '�P��
            Text1(ptxTANKA).BackColor = G_INPUT_NG
            Text1(ptxTANKA).Locked = True
            Text1(ptxTANKA).TabStop = False
            '�����c
            Text1(ptxZAN_QTY).BackColor = G_INPUT_OK
            Text1(ptxZAN_QTY).Locked = False
            Text1(ptxZAN_QTY).TabStop = True
    
    
    

    
        Case 1      '�����Ȃ���
    
            Input_Mode = True
    
    
            '�i��
            Text1(ptxHIN_GAI).BackColor = G_INPUT_OK
            Text1(ptxHIN_GAI).Locked = False
            Text1(ptxHIN_GAI).TabStop = True
                            
            '�S����
            Text1(ptxTANTO_CODE).BackColor = G_INPUT_OK
            Text1(ptxTANTO_CODE).Locked = False
            Text1(ptxTANTO_CODE).TabStop = True
            '������
            Text1(ptxORDER_CODE).BackColor = G_INPUT_OK
            Text1(ptxORDER_CODE).Locked = False
            Text1(ptxORDER_CODE).TabStop = True
            
            Combo1(pcmbORDER).BackColor = G_INPUT_OK
            Combo1(pcmbORDER).Locked = False
            Combo1(pcmbORDER).TabStop = True
    
            '�P��
            Text1(ptxTANKA).BackColor = G_INPUT_OK
            Text1(ptxTANKA).Locked = False
            Text1(ptxTANKA).TabStop = True
            '�����c
            Text1(ptxZAN_QTY).BackColor = G_INPUT_NG
            Text1(ptxZAN_QTY).Locked = True
            Text1(ptxZAN_QTY).TabStop = False
    
    End Select

End Sub

Private Function Hin_Item_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   �i�ڃ}�X�^�����������e�\��
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long
Dim i           As Integer

    Hin_Item_Disp_Proc = True
    
    
    If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI And _
        StrConv(ITEMREC.NAIGAI, vbUnicode) = NAIGAI_NAI And _
        Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) = Trim(Text1(ptxHIN_GAI).Text) Then
    
        Hin_Item_Disp_Proc = False
        Exit Function
    End If
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
        
            Text1(ptxHIN_NAME).Text = ""
            Text1(ptxZAIKO_QTY).Text = ""
        
            Hin_Item_Disp_Proc = BtErrKeyNotFound
            Exit Function
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    
    End Select
    Text1(ptxHIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    If Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                            StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
        Exit Function
    
    End If
    Text1(ptxZAIKO_QTY).Text = Format(SUMI_QTY + MI_QTY, "#0")
        
        
    Text1(ptxG_SHIIRE_KBN).Text = StrConv(ITEMREC.G_SHIIRE_KBN, vbUnicode)   '�d���敪
    Combo1(pcmbG_SHIIRE_KBN).ListIndex = -1
    For i = 0 To Combo1(pcmbG_SHIIRE_KBN).ListCount - 1
    
        If Text1(ptxG_SHIIRE_KBN).Text = Trim(Left(Right(Combo1(pcmbG_SHIIRE_KBN).List(i), 3), 2)) Then
            Combo1(pcmbG_SHIIRE_KBN).ListIndex = i
            Exit For
        End If
    
    Next i
    
    
    If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_ON Then
        Check1(chkZAIKO_F).Value = vbChecked
    Else
        Check1(chkZAIKO_F).Value = vbUnchecked
    End If
        
    Hin_Item_Disp_Proc = False
End Function
Private Function POS_NYUKA_Update_Proc(SOKO As String, Retu As String, Ren As String, Dan As String) As Integer
'----------------------------------------------------------------------------
'                   POS�p�݌Ɂ����ח\��X�V
'           POS���і����́A�W���I�Ԃɍ݌Ɍv�シ��2006.04.24
'----------------------------------------------------------------------------
                                            
Dim sts         As Integer
Dim com         As Integer


Dim DEN_NO      As String * 6
Dim ID_NO       As String * 9
Dim ans         As Integer
                                            
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long

Dim WK_Qty      As Long     '�O�؎c���[�N
Dim WK_E_QTY    As Long     '��s�o�א����[�N
                                            
Dim MAEGARI_QTY As Long
                                            
Dim SOUSAI_QTY  As Long
                                            
Dim Upd_QTY     As Long     '2007.05.03
                                            
Dim TO_SOKO     As String * 2
Dim TO_RETU     As String * 2
Dim TO_REN      As String * 2
Dim TO_DAN      As String * 2
                                            
    POS_NYUKA_Update_Proc = True
                                        
'    Call Input_Lock

''    If CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text) <= 0 Then
''        POS_NYUKA_Update_Proc = False
''        Exit Function
''    End If
    

    If Trim(SOKO) = "" Then
        TO_SOKO = KASO_NYUKA
        TO_RETU = "01"
        TO_REN = "01"
        TO_DAN = "01"
    Else
        '�o�n�r���і����͕W���I�Ԃ�
        TO_SOKO = SOKO
        TO_RETU = Retu
        TO_REN = Ren
        TO_DAN = Dan
    
    
        Call UniCode_Conv(K0_TANA.Soko_No, TO_SOKO)
        Call UniCode_Conv(K0_TANA.Retu, TO_RETU)
        Call UniCode_Conv(K0_TANA.Ren, TO_REN)
        Call UniCode_Conv(K0_TANA.Dan, TO_DAN)

    
        sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                '���o�^�͓��׉��z��
                TO_SOKO = KASO_NYUKA
                TO_RETU = "01"
                TO_REN = "01"
                TO_DAN = "01"
                    
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�I�}�X�^")
                Exit Function
        
        End Select
    
    
    End If








    WK_E_QTY = 0
                                            
    SUMI_QTY = 0
                            '���ޕi�͑S�Ė����i�Ƃ��Ĉ���
    MI_QTY = CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode))
                
                
    '���ޓ��������ް�(�O���ް�)�X�V
    Call UniCode_Conv(K0_P_NYU.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_P_NYU.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_P_NYU.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
    Call UniCode_Conv(K0_P_NYU.NYUKA_DT, "")
    
    com = BtOpGetGreater
    
    Do
        DoEvents
                
        Do
            sts = BTRV(com + BtSNoWait, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
            Select Case sts
                Case BtNoErr
                    
                    If StrConv(P_NYUREC.JGYOBU, vbUnicode) <> SHIZAI Or _
                        StrConv(P_NYUREC.NAIGAI, vbUnicode) <> NAIGAI_NAI Or _
                        StrConv(P_NYUREC.HIN_GAI, vbUnicode) <> StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode) Then
                        
                        sts = BTRV(BtOpUnlock, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "���ޑO���ް�")
                            Exit Function
                        End If
                        sts = BtErrEOF
                        Exit Do
                    End If
                    If IsNumeric(StrConv(P_NYUREC.SOUSAI_QTY, vbUnicode)) Then
                        SOUSAI_QTY = CLng(StrConv(P_NYUREC.SOUSAI_QTY, vbUnicode))
                    Else
                        SOUSAI_QTY = 0
                    End If
                    MAEGARI_QTY = CLng(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode)) - SOUSAI_QTY
                    If MAEGARI_QTY > MI_QTY Then
                        SOUSAI_QTY = SOUSAI_QTY + MI_QTY        '2007.05.03
                        MI_QTY = MAEGARI_QTY - MI_QTY
                        Call UniCode_Conv(P_NYUREC.SOUSAI_DT, Format(Now, "YYYYMMDD"))
                '        Call UniCode_Conv(P_NYUREC.SOUSAI_QTY, Format(MI_QTY, "00000000"))
                        Call UniCode_Conv(P_NYUREC.SOUSAI_QTY, Format(SOUSAI_QTY, "00000000"))
                
                        Do
                        
                            sts = BTRV(BtOpUpdate, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                    If ans = vbCancel Then
                                        Exit Function
                                    End If
                                Case Else
                                    Call File_Error(sts, BtOpUpdate, "���ޑO���ް�")
                                    Exit Function
                            End Select
                        
                        Loop
                        WK_E_QTY = CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode))  '��s������
                        Exit Do
                    Else
                        Do
                            sts = BTRV(BtOpDelete, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                    If ans = vbCancel Then
                                        Exit Function
                                    End If
                                Case Else
                                    Call File_Error(sts, BtOpDelete, "���ޑO���ް�")
                                    Exit Function
                            End Select
                        Loop
                        
                        
                        MI_QTY = MI_QTY - MAEGARI_QTY
                        WK_E_QTY = WK_E_QTY + MAEGARI_QTY
                    
                        If MI_QTY = 0 Then
                            sts = BtErrEOF
                            Exit Do
                        End If
                    
                    End If
            
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                   End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "���ޑO���ް�")
                    Exit Function
            End Select
        
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
        
        com = BtOpGetNext
    
    Loop
                                            '���ח\��ҏW
    Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)            '�����敪
    Call UniCode_Conv(Y_NYUREC.DT_SYU, "R")                     '�f�[�^���
    Call UniCode_Conv(Y_NYUREC.JGYOBU, SHIZAI)                  '���ƕ�
    Call UniCode_Conv(Y_NYUREC.NAIGAI, NAIGAI_NAI)              '�����O
    Call UniCode_Conv(Y_NYUREC.JGYOBA, "")                      '���Ə�
    Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")                    '�f�[�^�敪
    Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")                    '����敪
                                                                '�h�c��
    sts = Den_No_Set_Proc(11, SHIZAI, ID_NO)
    If sts Then
        Exit Function
    End If
    
    Call UniCode_Conv(Y_NYUREC.ID_NO, ID_NO)
    Call UniCode_Conv(Y_NYUREC.TEXT_NO, ID_NO)
                                                                '�i�ڔԍ�
    Call UniCode_Conv(Y_NYUREC.HIN_NO, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
                                                                
                                                                '�`�[��
    sts = Den_No_Set_Proc(10, SHIZAI, DEN_NO)
    If sts Then
        Exit Function
    End If
    Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
                                                                '�\�萔��
    Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode)), "0000000"))
    Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")                   '�o�ɐ�
    Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")                 '�o�Ɏ��x
                                                                '�o�ɓ��t
    Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, Format(Now, "YYYYMMDD"))
    Call UniCode_Conv(Y_NYUREC.TANKA, "")                       '�P��
    Call UniCode_Conv(Y_NYUREC.ODER_NO, "")                     '�I�[�_�[�ԍ�
    Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")                     '�A�C�e���ԍ�
    Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")                   '�I�[�_�[����
    Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")                 '���`��
                                                                '�o�ד�
    Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode))
                                                                '�I�ԂP
    Call UniCode_Conv(Y_NYUREC.TANABAN1, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                            StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                            StrConv(ITEMREC.ST_DAN, vbUnicode))
        
    Call UniCode_Conv(Y_NYUREC.TANABAN2, "")                    '�I�ԂQ
    Call UniCode_Conv(Y_NYUREC.TANABAN3, "")                    '�I�ԂR
    Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")                   '�o�ɐ於��
    Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")                     '�����敪
    Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")                '�����敪����
    Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")                     '���Y���P
    Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")                     '���Y���Q
    Call UniCode_Conv(Y_NYUREC.BIKOU2, "")                      '���l�Q
    Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")                     '�̔��敪
    Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")                   '�����敪
    Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")                  '�ƯďC��ID-NO
    Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")               '�݌Ɉ�������
    Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")              '�����Ǘ��ԍ�
    Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")                  '�󒍎c����
    Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")                  '�����敪
    Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")                '���i���[������x
    Call UniCode_Conv(Y_NYUREC.BIKOU1, "")                      '���l�P
    Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")                   '���[�敪
    Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")                  '�󒍕i�ڔԍ�
                                                                '�i��
    Call UniCode_Conv(Y_NYUREC.HIN_NAME, StrConv(ITEMREC.HIN_NAME, vbUnicode))
    Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")              '�i�ԕύX�敪
    Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")             '���W���[�������敪
    Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")                 '�c�݌ɂ܂Ƃߍ݌Ɏ��x�R�[�h
    Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")                   '�w��[��
    Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")            '�T�[�r�X��ЊǗ��ԍ�
    Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")                   '�@��i�ڃR�[�h
    Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")             '���K�i���i�敪
    Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD")) '�������t
                                                                '��s���א�
    Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, Format(WK_E_QTY, "00000000"))
    Call UniCode_Conv(Y_NYUREC.FILLER, "")
    
    Do
        sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
        
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case BtErrDuplicates
                                        '�������ԃf�[�^�d���͍Ĕ��s
                sts = Den_No_Set_Proc(11, SHIZAI, ID_NO)
                If sts Then
                    Exit Function
                End If

                Call UniCode_Conv(Y_NYUREC.ID_NO, ID_NO)
                Call UniCode_Conv(Y_NYUREC.TEXT_NO, ID_NO)
                
            Case Else
                Call File_Error(sts, BtOpInsert, "���ח\��f�[�^")
                Exit Function
        End Select
    Loop
                            
    sts = Nyuko_Update_Proc(SHIZAI, _
                            NAIGAI_NAI, _
                            StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode), _
                            StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode), _
                            (TO_SOKO & TO_RETU & TO_REN & TO_DAN), _
                            P_YOIN_TU_NYUKA, _
                            SUMI_QTY, _
                            CLng(Text1(ptxKONKAI_UKEIRE_QTY).Text), _
                            WS_NO, _
                            StrConv(P_SHORDER_REC.TANTO_CODE, vbUnicode), , _
                            MEMO_TEXT, _
                            StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode), _
                            StrConv(P_SHORDER_REC.TANKA, vbUnicode), _
                            StrConv(P_SHUKEIRE_REC.KEIJYO_YM, vbUnicode))
                            
                            
    If sts Then
        Exit Function
    End If


    '�O�؂萔�ō݌Ƀf�[�^�X�V�i�|�j
    If WK_E_QTY <> 0 Then
    '�݌Ƀf�[�^LOCK
        If Zaiko_Lock_Proc((TO_SOKO & TO_RETU & TO_REN & TO_DAN), _
                            SHIZAI, _
                            NAIGAI_NAI, _
                            StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode), _
                            WS_NO) Then
            Exit Function

        End If

        MI_QTY = WK_E_QTY
        SUMI_QTY = 0

        If Syuko_Update_Proc(SHIZAI, _
                            NAIGAI_NAI, _
                            StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode), _
                            StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode), _
                            (TO_SOKO & TO_RETU & TO_REN & TO_DAN), _
                            P_YOIN_MAE_SOUSAI, _
                            0, WK_E_QTY, 0, _
                            WS_NO, WS_NO) Then
            Exit Function

        End If






    End If



    POS_NYUKA_Update_Proc = False
End Function


