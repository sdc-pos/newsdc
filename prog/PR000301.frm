VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PR000301 
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
   LinkMode        =   1  '���
   LinkTopic       =   "Form1"
   ScaleHeight     =   10305
   ScaleWidth      =   14985
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   3840
      MaxLength       =   12
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   2040
      MaxLength       =   12
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   1
      Left            =   4800
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   3
      Left            =   10680
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   2
      Left            =   7920
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   2040
      MaxLength       =   12
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   5415
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   9551
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "�i��"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "�i��"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "���x�P��"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "�O���c"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "���ɐ�"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "�o�ɐ�"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "�����݌�"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "�d���P��"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "�O���݌ɋ��z"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "�����݌ɋ��z"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "�O���|����"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "�d����CODE"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "�ŏI�o�הN����"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "�ŏI�o�ɐ�"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "����t���ɐ���"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "�O�؎c"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "�����݌ɐ�"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "�����O���c"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "�o�^���t"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "�݌Ƀf�[�^�F�݌�"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   20
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=20"
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
      Splits(0)._ColumnProps(15)=   "Column(3).Width=1879"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=1773"
      Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=1879"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=1773"
      Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(25)=   "Column(5).Width=1879"
      Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=1773"
      Splits(0)._ColumnProps(28)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(30)=   "Column(6).Width=1879"
      Splits(0)._ColumnProps(31)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(6)._WidthInPix=1773"
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
      Splits(0)._ColumnProps(45)=   "Column(9).Width=2699"
      Splits(0)._ColumnProps(46)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(9)._WidthInPix=2593"
      Splits(0)._ColumnProps(48)=   "Column(9)._ColStyle=2"
      Splits(0)._ColumnProps(49)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(50)=   "Column(10).Width=2699"
      Splits(0)._ColumnProps(51)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(10)._WidthInPix=2593"
      Splits(0)._ColumnProps(53)=   "Column(10)._ColStyle=2"
      Splits(0)._ColumnProps(54)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(55)=   "Column(11).Width=1640"
      Splits(0)._ColumnProps(56)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(11)._WidthInPix=1535"
      Splits(0)._ColumnProps(58)=   "Column(11)._ColStyle=2"
      Splits(0)._ColumnProps(59)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(60)=   "Column(12).Width=3043"
      Splits(0)._ColumnProps(61)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(12)._WidthInPix=2937"
      Splits(0)._ColumnProps(63)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(64)=   "Column(13).Width=2408"
      Splits(0)._ColumnProps(65)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(66)=   "Column(13)._WidthInPix=2302"
      Splits(0)._ColumnProps(67)=   "Column(13)._ColStyle=2"
      Splits(0)._ColumnProps(68)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(69)=   "Column(14).Width=3069"
      Splits(0)._ColumnProps(70)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(14)._WidthInPix=2963"
      Splits(0)._ColumnProps(72)=   "Column(14)._ColStyle=2"
      Splits(0)._ColumnProps(73)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(74)=   "Column(15).Width=1879"
      Splits(0)._ColumnProps(75)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(76)=   "Column(15)._WidthInPix=1773"
      Splits(0)._ColumnProps(77)=   "Column(15)._ColStyle=2"
      Splits(0)._ColumnProps(78)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(79)=   "Column(16).Width=3810"
      Splits(0)._ColumnProps(80)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(81)=   "Column(16)._WidthInPix=3704"
      Splits(0)._ColumnProps(82)=   "Column(16)._ColStyle=2"
      Splits(0)._ColumnProps(83)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(84)=   "Column(17).Width=3493"
      Splits(0)._ColumnProps(85)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(86)=   "Column(17)._WidthInPix=3387"
      Splits(0)._ColumnProps(87)=   "Column(17)._ColStyle=2"
      Splits(0)._ColumnProps(88)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(89)=   "Column(18).Width=3810"
      Splits(0)._ColumnProps(90)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(18)._WidthInPix=3704"
      Splits(0)._ColumnProps(92)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(93)=   "Column(19).Width=3810"
      Splits(0)._ColumnProps(94)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(95)=   "Column(19)._WidthInPix=3704"
      Splits(0)._ColumnProps(96)=   "Column(19).Order=20"
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
      _StyleDefs(70)  =   "Splits(0).Columns(6).Style:id=82,.parent=43,.alignment=1"
      _StyleDefs(71)  =   "Splits(0).Columns(6).HeadingStyle:id=79,.parent=44"
      _StyleDefs(72)  =   "Splits(0).Columns(6).FooterStyle:id=80,.parent=45"
      _StyleDefs(73)  =   "Splits(0).Columns(6).EditorStyle:id=81,.parent=47"
      _StyleDefs(74)  =   "Splits(0).Columns(7).Style:id=20,.parent=43,.alignment=1"
      _StyleDefs(75)  =   "Splits(0).Columns(7).HeadingStyle:id=17,.parent=44"
      _StyleDefs(76)  =   "Splits(0).Columns(7).FooterStyle:id=18,.parent=45"
      _StyleDefs(77)  =   "Splits(0).Columns(7).EditorStyle:id=19,.parent=47"
      _StyleDefs(78)  =   "Splits(0).Columns(8).Style:id=98,.parent=43,.alignment=1"
      _StyleDefs(79)  =   "Splits(0).Columns(8).HeadingStyle:id=95,.parent=44"
      _StyleDefs(80)  =   "Splits(0).Columns(8).FooterStyle:id=96,.parent=45"
      _StyleDefs(81)  =   "Splits(0).Columns(8).EditorStyle:id=97,.parent=47"
      _StyleDefs(82)  =   "Splits(0).Columns(9).Style:id=24,.parent=43,.alignment=1"
      _StyleDefs(83)  =   "Splits(0).Columns(9).HeadingStyle:id=21,.parent=44"
      _StyleDefs(84)  =   "Splits(0).Columns(9).FooterStyle:id=22,.parent=45"
      _StyleDefs(85)  =   "Splits(0).Columns(9).EditorStyle:id=23,.parent=47"
      _StyleDefs(86)  =   "Splits(0).Columns(10).Style:id=102,.parent=43,.alignment=1"
      _StyleDefs(87)  =   "Splits(0).Columns(10).HeadingStyle:id=99,.parent=44"
      _StyleDefs(88)  =   "Splits(0).Columns(10).FooterStyle:id=100,.parent=45"
      _StyleDefs(89)  =   "Splits(0).Columns(10).EditorStyle:id=101,.parent=47"
      _StyleDefs(90)  =   "Splits(0).Columns(11).Style:id=78,.parent=43,.alignment=1"
      _StyleDefs(91)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=44"
      _StyleDefs(92)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=45"
      _StyleDefs(93)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=47"
      _StyleDefs(94)  =   "Splits(0).Columns(12).Style:id=62,.parent=43"
      _StyleDefs(95)  =   "Splits(0).Columns(12).HeadingStyle:id=59,.parent=44"
      _StyleDefs(96)  =   "Splits(0).Columns(12).FooterStyle:id=60,.parent=45"
      _StyleDefs(97)  =   "Splits(0).Columns(12).EditorStyle:id=61,.parent=47"
      _StyleDefs(98)  =   "Splits(0).Columns(13).Style:id=70,.parent=43,.alignment=1"
      _StyleDefs(99)  =   "Splits(0).Columns(13).HeadingStyle:id=67,.parent=44"
      _StyleDefs(100) =   "Splits(0).Columns(13).FooterStyle:id=68,.parent=45"
      _StyleDefs(101) =   "Splits(0).Columns(13).EditorStyle:id=69,.parent=47"
      _StyleDefs(102) =   "Splits(0).Columns(14).Style:id=115,.parent=43,.alignment=1"
      _StyleDefs(103) =   "Splits(0).Columns(14).HeadingStyle:id=112,.parent=44"
      _StyleDefs(104) =   "Splits(0).Columns(14).FooterStyle:id=113,.parent=45"
      _StyleDefs(105) =   "Splits(0).Columns(14).EditorStyle:id=114,.parent=47"
      _StyleDefs(106) =   "Splits(0).Columns(15).Style:id=86,.parent=43,.alignment=1"
      _StyleDefs(107) =   "Splits(0).Columns(15).HeadingStyle:id=83,.parent=44"
      _StyleDefs(108) =   "Splits(0).Columns(15).FooterStyle:id=84,.parent=45"
      _StyleDefs(109) =   "Splits(0).Columns(15).EditorStyle:id=85,.parent=47"
      _StyleDefs(110) =   "Splits(0).Columns(16).Style:id=90,.parent=43,.alignment=1"
      _StyleDefs(111) =   "Splits(0).Columns(16).HeadingStyle:id=87,.parent=44"
      _StyleDefs(112) =   "Splits(0).Columns(16).FooterStyle:id=88,.parent=45"
      _StyleDefs(113) =   "Splits(0).Columns(16).EditorStyle:id=89,.parent=47"
      _StyleDefs(114) =   "Splits(0).Columns(17).Style:id=106,.parent=43,.alignment=1"
      _StyleDefs(115) =   "Splits(0).Columns(17).HeadingStyle:id=103,.parent=44"
      _StyleDefs(116) =   "Splits(0).Columns(17).FooterStyle:id=104,.parent=45"
      _StyleDefs(117) =   "Splits(0).Columns(17).EditorStyle:id=105,.parent=47"
      _StyleDefs(118) =   "Splits(0).Columns(18).Style:id=94,.parent=43"
      _StyleDefs(119) =   "Splits(0).Columns(18).HeadingStyle:id=91,.parent=44"
      _StyleDefs(120) =   "Splits(0).Columns(18).FooterStyle:id=92,.parent=45"
      _StyleDefs(121) =   "Splits(0).Columns(18).EditorStyle:id=93,.parent=47"
      _StyleDefs(122) =   "Splits(0).Columns(19).Style:id=111,.parent=43"
      _StyleDefs(123) =   "Splits(0).Columns(19).HeadingStyle:id=108,.parent=44"
      _StyleDefs(124) =   "Splits(0).Columns(19).FooterStyle:id=109,.parent=45"
      _StyleDefs(125) =   "Splits(0).Columns(19).EditorStyle:id=110,.parent=47"
      _StyleDefs(126) =   "Named:id=33:Normal"
      _StyleDefs(127) =   ":id=33,.parent=0"
      _StyleDefs(128) =   "Named:id=34:Heading"
      _StyleDefs(129) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(130) =   ":id=34,.wraptext=-1"
      _StyleDefs(131) =   "Named:id=35:Footing"
      _StyleDefs(132) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(133) =   "Named:id=36:Selected"
      _StyleDefs(134) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(135) =   "Named:id=37:Caption"
      _StyleDefs(136) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(137) =   "Named:id=38:HighlightRow"
      _StyleDefs(138) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(139) =   "Named:id=39:EvenRow"
      _StyleDefs(140) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(141) =   "Named:id=40:OddRow"
      _StyleDefs(142) =   ":id=40,.parent=33"
      _StyleDefs(143) =   "Named:id=41:RecordSelector"
      _StyleDefs(144) =   ":id=41,.parent=34"
      _StyleDefs(145) =   "Named:id=42:FilterBar"
      _StyleDefs(146) =   ":id=42,.parent=33"
      _StyleDefs(147) =   "Named:id=107:Rstyle_Red"
      _StyleDefs(148) =   ":id=107,.parent=42,.bgcolor=&HFFFF&,.fgcolor=&HFF&"
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
      Left            =   12495
      TabIndex        =   18
      Top             =   9720
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�J �z"
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
      Left            =   11445
      TabIndex        =   17
      Top             =   9720
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���ו\"
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
      Index           =   9
      Left            =   10395
      TabIndex        =   16
      Top             =   9720
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�W�v�\"
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
      Index           =   8
      Left            =   9345
      TabIndex        =   15
      Top             =   9720
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�f�[�^"
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
      Left            =   8085
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   9720
      Width           =   1065
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
      Left            =   7035
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9720
      Width           =   1065
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
      Left            =   5985
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9720
      Width           =   1065
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
      Left            =   4935
      TabIndex        =   11
      Top             =   9720
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ďW�v"
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
      Left            =   3675
      TabIndex        =   10
      Top             =   9720
      Width           =   1065
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
      Left            =   2625
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9720
      Width           =   1065
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
      Left            =   1575
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9720
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�I���J�n"
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
      Index           =   0
      Left            =   525
      TabIndex        =   7
      Top             =   9720
      Width           =   1065
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   1695
      Index           =   0
      Left            =   210
      TabIndex        =   26
      Top             =   1560
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   2990
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
      Columns(1).Caption=   "�O���݌ɋ��z"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "�������ɋ��z"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "�����o�ɋ��z"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "�����݌ɋ��z"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "���z����-�O��"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1984"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2937"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2831"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=2"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2937"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2831"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2937"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2831"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2937"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2831"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2937"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2831"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
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
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=66,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(45)  =   ":id=66,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(46)  =   ":id=66,.fontname=�l�r �S�V�b�N"
      _StyleDefs(47)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=44"
      _StyleDefs(48)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=45"
      _StyleDefs(49)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=47"
      _StyleDefs(50)  =   "Splits(0).Columns(2).Style:id=32,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(51)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(52)  =   ":id=32,.fontname=�l�r �S�V�b�N"
      _StyleDefs(53)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=44"
      _StyleDefs(54)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=45"
      _StyleDefs(55)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=47"
      _StyleDefs(56)  =   "Splits(0).Columns(3).Style:id=74,.parent=43,.alignment=1"
      _StyleDefs(57)  =   "Splits(0).Columns(3).HeadingStyle:id=71,.parent=44"
      _StyleDefs(58)  =   "Splits(0).Columns(3).FooterStyle:id=72,.parent=45"
      _StyleDefs(59)  =   "Splits(0).Columns(3).EditorStyle:id=73,.parent=47"
      _StyleDefs(60)  =   "Splits(0).Columns(4).Style:id=82,.parent=43,.alignment=1"
      _StyleDefs(61)  =   "Splits(0).Columns(4).HeadingStyle:id=79,.parent=44"
      _StyleDefs(62)  =   "Splits(0).Columns(4).FooterStyle:id=80,.parent=45"
      _StyleDefs(63)  =   "Splits(0).Columns(4).EditorStyle:id=81,.parent=47"
      _StyleDefs(64)  =   "Splits(0).Columns(5).Style:id=20,.parent=43,.alignment=1"
      _StyleDefs(65)  =   "Splits(0).Columns(5).HeadingStyle:id=17,.parent=44"
      _StyleDefs(66)  =   "Splits(0).Columns(5).FooterStyle:id=18,.parent=45"
      _StyleDefs(67)  =   "Splits(0).Columns(5).EditorStyle:id=19,.parent=47"
      _StyleDefs(68)  =   "Named:id=33:Normal"
      _StyleDefs(69)  =   ":id=33,.parent=0"
      _StyleDefs(70)  =   "Named:id=34:Heading"
      _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(72)  =   ":id=34,.wraptext=-1"
      _StyleDefs(73)  =   "Named:id=35:Footing"
      _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   "Named:id=36:Selected"
      _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=37:Caption"
      _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(79)  =   "Named:id=38:HighlightRow"
      _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(81)  =   "Named:id=39:EvenRow"
      _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(83)  =   "Named:id=40:OddRow"
      _StyleDefs(84)  =   ":id=40,.parent=33"
      _StyleDefs(85)  =   "Named:id=41:RecordSelector"
      _StyleDefs(86)  =   ":id=41,.parent=34"
      _StyleDefs(87)  =   "Named:id=42:FilterBar"
      _StyleDefs(88)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label3 
      Height          =   375
      Index           =   1
      Left            =   8190
      TabIndex        =   30
      Top             =   9360
      Width           =   4635
   End
   Begin VB.Label Label3 
      Height          =   375
      Index           =   0
      Left            =   3570
      TabIndex        =   29
      Top             =   9360
      Width           =   4635
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�O����s����"
      Height          =   255
      Index           =   7
      Left            =   5565
      TabIndex        =   28
      Top             =   960
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   360
      Left            =   7350
      TabIndex        =   27
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�`"
      Height          =   255
      Index           =   6
      Left            =   3360
      TabIndex        =   25
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�������͓�"
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   24
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�v��N��"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   23
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�O��J�z��"
      Height          =   255
      Index           =   2
      Left            =   9360
      TabIndex        =   22
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   ")"
      Height          =   255
      Index           =   1
      Left            =   12000
      TabIndex        =   21
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "(�O��I���J�n��"
      Height          =   255
      Index           =   0
      Left            =   5850
      TabIndex        =   20
      Top             =   360
      Width           =   1965
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�I���J�n��"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   19
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "PR000301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'2010.10.28 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'���ޒI���G���[���O�p
Private STANA_LOG_Path      As String       '۸�̧���߽�iINI�捞�j
Private STANA_LOG_F         As String       '۸�̧�ٖ���
Private STANA_LOG_Out_Msg   As String       '���ޒI���װ۸ޏo�͗pܰ�

Private ZAIKO_MINUS_Msg     As String       '�݌�ϲŽү����

Private BEF_Hin_GAI         As String       '�����O���݌ɕ\������
'2010.10.28 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Private IN_YOIN         As Variant          ' ���O���ɗv��
Private OUT_YOIN        As Variant          ' ���O�o�ɗv��
Private START_DT        As String           ' ����I�����J�n��
Private LAST_START_DT   As String           ' �O��I�����J�n��
Private LAST_SHIME_DT   As String           ' �O����ߏ�����


Private PR00030_LOG_F   As String           ' �I���������p���O  2007.10.16

Private STAT_F   As Integer                 ' �I������ԃt���O  2007.10.16


Private ZEI_SHIIRE_KBN As String * 2



Private G_SYUSHI_TBL    As Variant          ' �Ώێ��x      2007.11.13


Private ZAIKO_FILE      As String           ' �݌Ƀ`�F�b�NF  2015.03.05



'�e�L�X�g�p�Y��
Private Const ptxSTART_DT% = 0              '�I�����J�n��
Private Const ptxKEIJYO_YM% = 1             '�v��N��
Private Const ptxLAST_START_DT% = 2         '�O��I�����J�n��
Private Const ptxLAST_SHIME_DT% = 3         '�O����ߓ�

Private Const ptxS_INPUT_DT% = 4            '�������͊J�n��
Private Const ptxE_INPUT_DT% = 5            '�������͏I����


'�R���{�p�Y��

'�`�F�b�N�{�b�N�X�p�Y��


Private Const pcmdStart% = 0                '�J�n
Private Const pcmdRE_Start% = 3             '�ďW�v
Private Const pcmdNext% = 10                '�J�z


'Glid�p��---------------------------------

Private Const pSum_GridSTOCK% = 0
Private Const pGridSTOCK% = 1


Private Sum_STOCK       As New XArrayDB

Private Const Sum_Min_Row% = 1              '�ŏ��s��
Private Const Sum_Min_Col% = 0              '�ŏ���
Private Const Sum_Max_Col% = 5              '�ő��

Private Const colSum_G_SYUSHI% = 0          '���x�P��
Private Const colSum_ZEN_ZAIKO_KIN% = 1     '�O���݌ɋ��z
Private Const colSum_NYUKO_KIN% = 2         '�������ɋ��z
Private Const colSum_SYUKO_KIN% = 3         '�����o�ɋ��z
Private Const colSum_ZAIKO_KIN% = 4         '�����݌ɋ��z
Private Const colSum_SA_KIN% = 5            '���z



Private STOCK       As New XArrayDB


Private Const Min_Row% = 1                  '�ŏ��s��
Private Const Min_Col% = 0                  '�ŏ���
Private Const Max_Col% = 18                 '�ő��       2010.10.28 Upd

Private Const colHIN_GAI% = 0               '���ޕi��
Private Const colHIN_NAME% = 1              '�i��
Private Const colG_SYUSHI% = 2              '�݌Ɍ��i���x�j
Private Const colZEN_ZAIKO_QTY% = 3         '�O���݌�
Private Const colNYUKO_QTY% = 4             '��������
Private Const colSYUKO_QTY% = 5             '�����o��
Private Const colZAIKO_QTY% = 6             '�����݌�
Private Const colSHI_TANKA% = 7             '�d���P��


Private Const colZEN_ZAIKO_KIN% = 8         '�O���݌ɋ��z


Private Const colZAIKO_KIN% = 9             '�����݌ɋ��z

Private Const colSA_ZAIKO_KIN% = 10         '�O���|����


Private Const colSHI_CODE% = 11             '�����d���溰��

Private Const colLAST_SYUKA_DT% = 12        '�ŏI�o�ɓ�
Private Const colLAST_SYUKA_QTY% = 13       '�ŏI�o�ɐ�


Private Const colSAKI_SHIIRE% = 14          '����t�d����   '2017.04.22


Private Const colMAEGARI_QTY% = 15          '�O�ؐ�


Private Const colMOTO_ZAIKO_QTY% = 16       '�����݌ɐ�

Private Const colSAV_ZEN_ZAIKO% = 17        '�����O���c     2010.10.28 Add

Private Const colINPUT_DATE% = 18           '���͓��t


Private Sort_Tbl(Min_Col To Max_Col) _
                As Integer                  '��Ă̐��� 0:���� 1:�~��
Private Tbl_Set_F   As Boolean

'   �I�����ް����߼޼��ݸ�
Private wP_STOCK_POS    As POSBLK
Private wP_STOCK_REC    As tmpP_STOCK_REC_Tag
Private K0_wP_STOCK     As KEY0_tmpP_STOCK
Private K1_wP_STOCK     As KEY1_tmpP_STOCK
Private K2_wP_STOCK     As KEY2_tmpP_STOCK





'2010.01.14
Private wkP_STOCK_POS    As POSBLK
Private wkP_STOCK_REC    As P_STOCK_REC_Tag
Private K0_wkP_STOCK     As KEY0_P_STOCK
'2010.01.14

Private K1_wkP_STOCK     As KEY1_P_STOCK    '2011.02.22






Private RE_UPDATE_F     As Integer          '2012.12.31



'EXCEL�V�[�g    2009.01.17
Private exSheet          As String
Private Const LStart% = 4

Private Const exHin_Gai% = 1
Private Const exHin_Name% = 2
Private Const exG_SYUSHI% = 3
Private Const exZEN_ZAIKO_QTY% = 4
Private Const exNYUKO_QTY% = 5
Private Const exSYUKO_QTY% = 6
Private Const exZAIKO_QTY% = 7
Private Const exSHI_TANKA% = 8
Private Const exZAIKO_KIN% = 9
Private Const exSHI_CODE% = 10
Private Const exLAST_SYUKA_DT% = 11
Private Const exLAST_SYUKA_QTY% = 12
Private Const exMAEGARI_QTY% = 13
Private Const exLOCATION% = 14

'Private Const Last_Update_Day$ = "���ލ݌ɒI�����\���s PR00030 (2018.07.25 14:00)"
Private Const Last_Update_Day$ = "���ލ݌ɒI�����\���s PR00030 (2018.08.21 14:00)"

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PR000301.MousePointer = vbHourglass

    Call Ctrl_Lock(PR000301)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PR000301)


    PR000301.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer, Sel As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer

Dim wkdate      As String * 10

Dim i           As Integer

    Error_Check_Proc = True

    Select Case Mode




        Case ptxSTART_DT        '�I�����J�n��


          If Not IsDate(Text1(Mode).Text) Then
'              MsgBox "���͂������ڂ̓G���[�ł��B"              '2016.01.07
              MsgBox "�I���J�n���𐳂������͂��ĉ������B"       '2016.01.07
              Text1(Mode).SetFocus
              Exit Function
          Else

              Text1(Mode).Text = Format(CDate(Text1(Mode).Text), "YYYY/MM/DD")

              Text1(ptxKEIJYO_YM).Text = Left(Format(CDate(Text1(Mode).Text), "YYYY/MM/DD"), 7)


              If Text1(Mode).Text < Text1(ptxLAST_START_DT).Text Then
'                  MsgBox "���͂������ڂ̓G���[�ł��B"                  '2016.01.07
                    MsgBox "�I���J�n�����O��J�n�����ߋ����t�ł�"     '2016.01.07
                    Text1(Mode).SetFocus
                    Exit Function
              End If

              If Text1(Mode).Text < Text1(ptxLAST_SHIME_DT).Text Then
'                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    MsgBox "�I���J�n�����O��J�z�����ߋ����t�ł�"     '2016.01.07
                    Text1(Mode).SetFocus
                    Exit Function
              End If

          End If
        Case ptxS_INPUT_DT
            If Sel = 0 Then
            Else
                If Trim(Text1(Mode).Text) = "" Then
                Else
                    If Not IsDate(Text1(Mode).Text) Then
'                        MsgBox "���͂������ڂ̓G���[�ł��B"                '2016.01.07
                        MsgBox "�����J�n���𐳂������͂��ĉ������B"         '2016.01.07
                        Text1(Mode).SetFocus
                        Exit Function
                    Else

                        Text1(Mode).Text = Format(CDate(Text1(Mode).Text), "YYYY/MM/DD")
                    End If
                End If
            End If

        Case ptxE_INPUT_DT
            If Sel = 0 Then
            Else
                If Trim(Text1(Mode).Text) = "" Then
                Else

                    If Not IsDate(Text1(Mode).Text) Then
'                        MsgBox "���͂������ڂ̓G���[�ł��B"                '2016.01.07
                        MsgBox "�����I�����𐳂������͂��ĉ������B"         '2016.01.07
                        Text1(Mode).SetFocus
                        Exit Function
                    Else

                        Text1(Mode).Text = Format(CDate(Text1(Mode).Text), "YYYY/MM/DD")
                    End If
                End If
            End If

    End Select


    Error_Check_Proc = False


End Function


Private Sub Command1_Click(Index As Integer)
Dim ans         As Integer
Dim i           As Integer

Dim c           As String * 128

Dim sts         As Integer
Dim cmd         As Integer


Dim ZAIKO_CHK_F   As Boolean  '2015.03.05

Dim yn          As Integer  '2016.01.07

    Select Case Index

        Case pcmdStart          '�J�n


            If Trim(ZAIKO_FILE) <> "" Then                                  '2015.03.05


                If ZAIKO_CHK_PROC(ZAIKO_CHK_F) Then                         '2015.03.05
                    Exit Sub                                                '2015.03.05
                End If                                                      '2015.03.05
    
                If ZAIKO_CHK_F Then                                         '2015.03.05
                    sts = Shell("Notepad.exe " & ZAIKO_FILE, vbNormalFocus) '2015.03.05
                End If                                                      '2015.03.05
            
            End If                                                          '2015.03.05



            For i = ptxSTART_DT To ptxE_INPUT_DT

                If Error_Check_Proc(i, 0) Then    '�G���[�`�F�b�N
                    Exit Sub
                End If

            Next i


            If MULTI_TANKA_CHECK_PROC(yn) Then          '2016.01.07
                Exit Sub                                '2016.01.07
            End If                                      '2016.01.07

            If yn = vbNo Then                           '2016.01.07
                Exit Sub                                '2016.01.07
            End If                                      '2016.01.07


            START_DT = Text1(ptxSTART_DT).Text

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'            '2007.10.16��
'            Select Case STAT_F
'
'                Case 1
'
'
'
'                    If Left(START_DT, 7) = Left(LAST_START_DT, 7) Then
'
'                        ans = MsgBox("�I���J�n�����͊��Ɏ��s����Ă��܂��B�I���J�n���������s���܂����H", vbYesNo + vbDefaultButton2 + vbCritical, "�m�F����")
'                        If ans = vbNo Then
'                            Exit Sub
'                        End If
'                    Else
'
'                        ans = MsgBox("�J�z�������I�����Ă��܂���B", vbOK + vbCritical, "�m�F����")
'
'                        Exit Sub
'
'
'
'                    End If
'
'
'
'                Case 9
'
'                    If Left(START_DT, 7) = Left(LAST_START_DT, 7) Then
'                        MsgBox "�J�z�����͏I�����Ă��܂��B"
'                        Exit Sub
'
'                    Else
'                    End If
'            End Select
'
'
'
'
'
'            '2007.10.16��
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






            ans = MsgBox("�����̒I�����������J�n���܂����H", vbYesNo, "�m�F����")
            If ans = vbYes Then

                '2007.10.16


                If Trim(GLB_SYUSHI_F) = "" Then '2007.11.13
                    Call LOG_OUT(PR00030_LOG_F, "���ޒI���������u�J�n�v")
                Else
                    Call LOG_OUT(PR00030_LOG_F, "[" & Trim(GLB_SYUSHI_F) & "] " & "���ޒI���������u�J�n�v")
                End If


'2010.10.28 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                                            '���ޒI���װ۸�̧�ٖ��� �ҏW
                STANA_LOG_F = STANA_LOG_Path & "\���ޒI���G���[" & _
                              Format(Now, "yyyymmddhhnn") & ".txt"
                STANA_LOG_Out_Msg = ""      '���ޒI���װ۸ޏo�͗L��ү���� �ر

                '���ޒI���i�ڃ}�X�^�ۑ� �t�@�C���N���A
                Do
                    DoEvents

                    sts = BTRV(BtOpGetFirst, T_ITEMSV_POS, T_ITEMSVREC, Len(T_ITEMSVREC), K0_T_ITEMSV, Len(K0_T_ITEMSV), 0)
                    Select Case sts
                        Case BtNoErr
                            sts = BTRV(BtOpDelete, T_ITEMSV_POS, T_ITEMSVREC, Len(T_ITEMSVREC), K0_T_ITEMSV, Len(K0_T_ITEMSV), 0)
                            If sts <> BtNoErr Then
                                Call File_Error(sts, BtOpDelete, "���ޒI���i�ڃ}�X�^�ۑ�")
                                Exit Sub
                            End If

                        Case BtErrEOF
                            Exit Do

                        Case Else
                            Call File_Error(sts, BtOpGetFirst, "���ޒI���i�ڃ}�X�^�ۑ�")
                            Exit Sub
                    End Select
                Loop
'2010.10.28 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


                If Update_Proc() Then
                    Unload Me
                End If

                If List_Disp_Proc() Then
                    Unload Me
                End If

                STAT_F = 1

                                    '2016.01.07 P_SYS.INI -- > PR00030.INI
                If WriteIni(App.EXEName, "LAST_FUNCTION" & Trim(GLB_SYUSHI_F), App.EXEName, Format(Now, "YYYY/MM/DD HH:MM:SS") & " �I���J�n") Then
'                If WriteIni(App.EXEName, "LAST_FUNCTION" & Trim(GLB_SYUSHI_F), "p_sys", Format(Now, "YYYY/MM/DD HH:MM:SS") & " �I���J�n") Then
                    Beep
                    MsgBox ("INI�t�@�C���̏������݂Ɏ��s���܂����B" & App.EXEName & "LAST_FUNCTION")
                    Unload Me
                End If


                DoEvents

                Label2.Caption = Format(Now, "YYYY/MM/DD HH:MM:SS") & " �I���J�n"


'2010.10.28 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
''                MsgBox "�I���J�n�����@�I��"     '2010.01.14

                If STANA_LOG_Out_Msg = "" Then
                    MsgBox "�I���J�n�����@�I��"
                Else
                    MsgBox "�I���J�n�����@�I��" & vbCrLf & vbCrLf & _
                           "���ޒI���G���[���O���쐬����܂����B" & vbCrLf & _
                           "�i" & STANA_LOG_F & "�j", vbExclamation, "�x��"
                End If
'2010.10.28 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

            End If



        Case pcmdRE_Start       '�ďW�v

            For i = ptxSTART_DT To ptxE_INPUT_DT

                If Error_Check_Proc(i, 1) Then    '�G���[�`�F�b�N
                    Exit Sub
                End If

            Next i


            ans = MsgBox("�����̒I�����Čv�Z���J�n���܂����H", vbYesNo, "�m�F����")
            If ans = vbYes Then

                '2007.10.16

                If Trim(GLB_SYUSHI_F) = "" Then '2007.11.13
                    Call LOG_OUT(PR00030_LOG_F, "���ޒI���������u�ďW�v�v")
                Else
                    Call LOG_OUT(PR00030_LOG_F, "[" & Trim(GLB_SYUSHI_F) & "] " & "���ޒI���������u�ďW�v�v")
                End If


'2010.10.28 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                                            '���ޒI���װ۸�̧�ٖ��� �ҏW
                STANA_LOG_F = STANA_LOG_Path & "\���ޒI���G���[" & _
                              Format(Now, "yyyymmddhhnn") & ".txt"
                STANA_LOG_Out_Msg = ""      '���ޒI���װ۸ޏo�͗L��ү���� �ر
'2010.10.28 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


                If RE_Update_Proc() Then
                    Unload Me
                End If

                If List_Disp_Proc() Then
                    Unload Me
                End If
                                                        '2016.01.07 P_SYS.INI -- > PR00030.INI
                If WriteIni(App.EXEName, "LAST_FUNCTION" & Trim(GLB_SYUSHI_F), App.EXEName, Format(Now, "YYYY/MM/DD HH:MM:SS") & " �ďW�v") Then
'                If WriteIni(App.EXEName, "LAST_FUNCTION" & Trim(GLB_SYUSHI_F), "p_sys", Format(Now, "YYYY/MM/DD HH:MM:SS") & " �ďW�v") Then
                    Beep
                    MsgBox ("INI�t�@�C���̏������݂Ɏ��s���܂����B" & App.EXEName & "LAST_FUNCTION")
                    Unload Me
                End If

                DoEvents

                Label2.Caption = Format(Now, "YYYY/MM/DD HH:MM:SS") & " �ďW�v"


'2010.10.28 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
''                MsgBox "�I���ďW�v�����@�I��"     '2010.01.14

                If STANA_LOG_Out_Msg = "" Then
                    MsgBox "�I���ďW�v�����@�I��"
                Else
                    MsgBox "�I���ďW�v�����@�I��" & vbCrLf & vbCrLf & _
                           "���ޒI���G���[���O���쐬����܂����B" & vbCrLf & _
                           "�i" & STANA_LOG_F & "�j", vbExclamation, "�x��"
                End If
'2010.10.28 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

            End If



        Case P_CMD_DSP                      '����/�\��

            For i = ptxSTART_DT To ptxE_INPUT_DT

                If Error_Check_Proc(i, 0) Then    '�G���[�`�F�b�N
                    Exit Sub
                End If

            Next i


            If List_Disp_Proc() Then
                Unload Me
            End If



        Case P_CMD_OUT                      '�ް��o��


            '2009.01.16
            ans = MsgBox("�f�[�^�o�͂��J�n���܂����H", vbYesNo, "�m�F����")
            If ans = vbYes Then


                If Data_Out_Proc() Then
                    Unload Me
                End If



            End If







        Case P_CMD_PRT                      '���   �W�v�\�̈���̂�    2007.07.12

            For i = ptxSTART_DT To ptxE_INPUT_DT

                If Error_Check_Proc(i, 0) Then    '�G���[�`�F�b�N
                    Exit Sub
                End If

            Next i

            ans = MsgBox("�u�W�v�\�v������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then

                If Print_Proc(1) Then
                    Unload Me
                End If
                                                    '2016.01.07 P_SYS.INI -- > PR00030.INN
                If WriteIni(App.EXEName, "LAST_FUNCTION" & Trim(GLB_SYUSHI_F), App.EXEName, Format(Now, "YYYY/MM/DD HH:MM:SS") & " �W�v�\���") Then
'                If WriteIni(App.EXEName, "LAST_FUNCTION" & Trim(GLB_SYUSHI_F), "p_sys", Format(Now, "YYYY/MM/DD HH:MM:SS") & " �W�v�\���") Then
                    Beep
                    MsgBox ("INI�t�@�C���̏������݂Ɏ��s���܂����B" & App.EXEName & "LAST_FUNCTION")
                    Unload Me
                End If

                DoEvents

                Label2.Caption = Format(Now, "YYYY/MM/DD HH:MM:SS") & " �W�v�\���"

                MsgBox "�I���W�v�\��������@�I��"     '2010.01.14



            End If

        Case 9                              '���   ���ו\�̈���̂�    2007.07.12

            For i = ptxSTART_DT To ptxE_INPUT_DT

                If Error_Check_Proc(i, 0) Then    '�G���[�`�F�b�N
                    Exit Sub
                End If

            Next i

            ans = MsgBox("�u���ו\�v������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then

                If Print_Proc(2) Then
                    Unload Me
                End If
                                                    '2016.01.07 P_SYS.INI PR00030.INI
                If WriteIni(App.EXEName, "LAST_FUNCTION" & Trim(GLB_SYUSHI_F), App.EXEName, Format(Now, "YYYY/MM/DD HH:MM:SS") & " ���ו\���") Then
'                If WriteIni(App.EXEName, "LAST_FUNCTION" & Trim(GLB_SYUSHI_F), "p_sys", Format(Now, "YYYY/MM/DD HH:MM:SS") & " ���ו\���") Then
                    Beep
                    MsgBox ("INI�t�@�C���̏������݂Ɏ��s���܂����B" & App.EXEName & "LAST_FUNCTION")
                    Unload Me
                End If

                DoEvents

                Label2.Caption = Format(Now, "YYYY/MM/DD HH:MM:SS") & " ���ו\���"

                MsgBox "�I�����ו\��������@�I��"     '2010.01.14


            End If


        Case pcmdNext                       '�J�z

            
'            START_DT = Text1(ptxSTART_DT).Text          '2016.09.30
            
            
            '2007.10.16��
            Select Case STAT_F
                Case 0


                    If Left(START_DT, 7) = Left(LAST_START_DT, 7) Then
                        MsgBox "�����̒I�������������s����Ă��܂���B"
                        Exit Sub

                    Else

                        MsgBox "�����̒I�������������s����Ă��܂���B"
                        Exit Sub





                    End If


                Case 1



                    If Left(START_DT, 7) = Left(LAST_START_DT, 7) Then
                    Else

                        MsgBox "�����̒I�������������s����Ă��܂���B"
                        Exit Sub


                    End If



                Case 9

                    
                    
                    
                    If Left(START_DT, 7) = Left(LAST_START_DT, 7) Then

                        MsgBox "�J�z�����͏I�����Ă��܂��B"
                        Exit Sub
                    Else
                    
                    End If






            End Select




            '2007.10.16��







            ans = MsgBox("�J�z���������{���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                ans = MsgBox("�����́u���ޒI���\�v�̈���͏I�����Ă��܂����H", vbYesNo + vbQuestion, "�m�F����")

                If ans = vbYes Then

                    '2007.10.16
                    If Trim(GLB_SYUSHI_F) = "" Then '2007.11.13
                        Call LOG_OUT(PR00030_LOG_F, "���ޒI���������u�J�z�v")
                    Else
                        Call LOG_OUT(PR00030_LOG_F, "[" & Trim(GLB_SYUSHI_F) & "] " & "���ޒI���������u�J�z�v�v")
                    End If


                    If Next_Proc() Then
                        Unload Me
                    End If

                    STAT_F = 9
                                                                            '2016.01.07 P_SYS.INI -- > PR00030.INI
                    If WriteIni(App.EXEName, "LAST_FUNCTION" & Trim(GLB_SYUSHI_F), App.EXEName, Format(Now, "YYYY/MM/DD HH:MM:SS") & " �J�z") Then
'                    If WriteIni(App.EXEName, "LAST_FUNCTION" & Trim(GLB_SYUSHI_F), "p_sys", Format(Now, "YYYY/MM/DD HH:MM:SS") & " �J�z") Then
                        Beep
                        MsgBox ("INI�t�@�C���̏������݂Ɏ��s���܂����B" & App.EXEName & "LAST_FUNCTION")
                        Unload Me
                    End If


                    DoEvents
                    Label2.Caption = Format(Now, "YYYY/MM/DD HH:MM:SS") & " �J�z"


                    Text1(3).Text = Format(Now, "YYYY/MM/DD")   '2010.01.14


                End If

            End If


        Case P_CMD_End                      '�I��



                                    '�h�m�h����F
                                                            '2016.01.07 P_SYS.INI -- > PR00030.INI
            If WriteIni(App.EXEName, "STAT" & Trim(GLB_SYUSHI_F), App.EXEName, Format(STAT_F, "0")) Then
'            If WriteIni(App.EXEName, "STAT" & Trim(GLB_SYUSHI_F), "p_sys", Format(STAT_F, "0")) Then
                Beep
                MsgBox ("INI�t�@�C���̏������݂Ɏ��s���܂����B" & App.EXEName & "STAT")
                Unload Me
            End If


            Unload Me

        Case 5


            ans = MsgBox(TDBGrid1(1).Bookmark & "�s�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then



                If Gyo_Update_Proc(TDBGrid1(1).Bookmark) Then
                    Unload Me
                End If




            End If


    End Select

End Sub

Private Sub Form_DblClick()
'    PrintForm              2017.04.22
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �j���� �c������ �O����
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command1(KeyCode - vbKeyF1).Value = True


        Case vbKeyZ
            If Shift = vbShiftMask Then

                Command1(5).Enabled = True
                Command1(5).Caption = "�X �V"

                TDBGrid1(1).AllowUpdate = True


            End If


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

'    If App.PrevInstance Then
'        Beep
'        MsgBox "����v���O�������s���ł��B"
'        End
'    End If

                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)

                                '���ޒI���G���[���O�p�X��荞��     2010.10.28
    If GetIni("FILE", "STANA_LOG", "SYS", c) Then
        Beep
        MsgBox "���ޒI���G���[���O�p�X�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    STANA_LOG_Path = RTrim(c)

    '-----------------------------------    *
    ' P_SYS.INI -- > PR00030.INI 2015.03.03
    '-----------------------------------    *
    



                                '�I���������p���O�t�@�C������荞�� 2007.10.16
    If GetIni(StrConv(App.EXEName, vbUpperCase), "LOGF", StrConv(App.EXEName, vbUpperCase), c) Then     '2016.01.07
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "LOGF", "p_sys", c) Then     '2016.01.07
        Beep
        MsgBox "�I���������p���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    PR00030_LOG_F = RTrim(c)



                                '�I���������p�`�F�b�N�t�@�C������荞�� 2015.03.05
    If GetIni(StrConv(App.EXEName, vbUpperCase), "ZAIKO_FILE", StrConv(App.EXEName, vbUpperCase), c) Then
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "ZAIKO_FILE", "p_sys", c) Then
        ZAIKO_FILE = ""
    Else
       ZAIKO_FILE = RTrim(c)
    End If





                                '�Ώێ��x�̊l�� 2007.11.13
    If Trim(GLB_SYUSHI_F) = "" Then

    Else

        If GetIni(StrConv(App.EXEName, vbUpperCase), GLB_SYUSHI_F, StrConv(App.EXEName, vbUpperCase), c) Then   '2016.01.07
'        If GetIni(StrConv(App.EXEName, vbUpperCase), GLB_SYUSHI_F, "p_sys", c) Then   '2016.01.07
            Beep
            MsgBox "�Ώێ��x�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
        End If

        G_SYUSHI_TBL = Split(Trim(c), ",", -1)
    End If


                                '�I�������F��荞�� 2007.10.16-->���x�敪�ǉ� 2007.11.13
    If GetIni(StrConv(App.EXEName, vbUpperCase), "STAT" & Trim(GLB_SYUSHI_F), StrConv(App.EXEName, vbUpperCase), c) Then    '2016.01.07
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "STAT" & Trim(GLB_SYUSHI_F), "p_sys", c) Then    '2016.01.07
        STAT_F = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            STAT_F = 0
        Else
            STAT_F = CInt(Trim(c))
        End If
    End If




                                '����ŕ�
    If GetIni(StrConv(App.EXEName, vbUpperCase), "ZEI_SHIIRE_KBN", StrConv(App.EXEName, vbUpperCase), c) Then   '2016.01.07
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "ZEI_SHIIRE_KBN", "p_sys", c) Then   '2016.01.07
        ZEI_SHIIRE_KBN = ""
    Else
        ZEI_SHIIRE_KBN = Trim(c)
    End If


                                '���O���ɗv����荞��
    If GetIni(StrConv(App.EXEName, vbUpperCase), "IN_YOIN", StrConv(App.EXEName, vbUpperCase), c) Then          '2016.01.07
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "IN_YOIN", "p_sys", c) Then          '2016.01.07
        IN_YOIN = ""
    Else
        IN_YOIN = Split(Trim(c), ",", -1)
    End If

                                '���O�o�ɗv����荞��
    If GetIni(StrConv(App.EXEName, vbUpperCase), "OUT_YOIN", StrConv(App.EXEName, vbUpperCase), c) Then         '2016.01.07
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "OUT_YOIN", "p_sys", c) Then         '2016.01.07
        OUT_YOIN = ""
    Else
        OUT_YOIN = Split(Trim(c), ",", -1)
    End If


                                '�J�n����荞��-->���x�敪�ǉ� 2007.11.13
    If GetIni(StrConv(App.EXEName, vbUpperCase), "LAST_START_DT" & Trim(GLB_SYUSHI_F), StrConv(App.EXEName, vbUpperCase), c) Then   '2016.01.07
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "LAST_START_DT" & Trim(GLB_SYUSHI_F), "p_sys", c) Then   '2016.01.07
        START_DT = ""
    Else
        START_DT = Trim(c)
    End If


                                '�O��I�����J�n����荞��-->���x�敪�ǉ� 2007.11.13
    If GetIni(StrConv(App.EXEName, vbUpperCase), "LAST_START_DT" & Trim(GLB_SYUSHI_F), StrConv(App.EXEName, vbUpperCase), c) Then   '2016.01.07
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "LAST_START_DT" & Trim(GLB_SYUSHI_F), "p_sys", c) Then   '2016.01.07
        LAST_START_DT = ""
    Else
        LAST_START_DT = Trim(c)
    End If

                                '�O��I�������ߓ���荞��-->���x�敪�ǉ� 2007.11.13
    If GetIni(StrConv(App.EXEName, vbUpperCase), "LAST_SHIME_DT" & Trim(GLB_SYUSHI_F), StrConv(App.EXEName, vbUpperCase), c) Then   '2016.01.07
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "LAST_SHIME_DT" & Trim(GLB_SYUSHI_F), "p_sys", c) Then   '2016.01.07
        LAST_SHIME_DT = ""
    Else
        LAST_SHIME_DT = Trim(c)
    End If


                                '�O�񏈗����e��荞�� 2007.10.31-->���x�敪�ǉ� 2007.11.13
    If GetIni(StrConv(App.EXEName, vbUpperCase), "LAST_FUNCTION" & Trim(GLB_SYUSHI_F), StrConv(App.EXEName, vbUpperCase), c) Then   '2016.01.07
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "LAST_FUNCTION" & Trim(GLB_SYUSHI_F), "P_sys", c) Then   '2016.01.07
        Label2.Caption = ""
    Else
        Label2.Caption = Trim(c)
    End If


                                ' 2012.12.13 RE_UPDATE_F
    If GetIni(StrConv(App.EXEName, vbUpperCase), "RE_UPDATE_F", StrConv(App.EXEName, vbUpperCase), c) Then  '2016.01.07
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "RE_UPDATE_F", "p_sys", c) Then  '2016.01.07
        RE_UPDATE_F = 0
    Else
        If Trim(c) = "1" Then
            RE_UPDATE_F = 1
        Else
            RE_UPDATE_F = 0
        End If
    End If







                                'EXCEL�o�͗p��̫�ļ�� 2009.01.17
    If GetIni(StrConv(App.EXEName, vbUpperCase), "EXCEL_SHEET", StrConv(App.EXEName, vbUpperCase), c) Then  '2016.01.07
'    If GetIni(StrConv(App.EXEName, vbUpperCase), "EXCEL_SHEET", "p_sys", c) Then  '2016.01.07
        exSheet = ""


        Command1(P_CMD_OUT).Enabled = False


    Else
        exSheet = Trim(c)

        Command1(P_CMD_OUT).Enabled = True


    End If






                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '�݌ɂn�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '�݌Ɉړ����n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '����Ͻ��n�o�d�m
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If



                                '���ޒ����n�o�d�m
    If P_SHORDER_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ގ���n�o�d�m
    If P_SHUKEIRE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ޒI�����n�o�d�m
    If P_STOCK_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ޑO���ް��n�o�d�m
    If P_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If


                                '���ޒI�����W�v�n�o�d�m
    If P_STOCKSUM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�Ǘ��}�X�^�n�o�d�m
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If






                                '���ޒI�����n�o�d�m 2010.01.14
    If wkP_STOCK_Open(BtOpenNomal) Then
        Unload Me
    End If


'2010.10.28 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                                '���ޒI���i�ڃ}�X�^�ۑ��n�o�d�m
    If T_ITEMSV_Open(BtOpenNomal) Then
        Unload Me
    End If
'2010.10.28 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    

    Load PR000302



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

    PR000301.Caption = Last_Update_Day

    '��ʏ����ݒ�
    If Init_Proc() Then
        Unload Me
    End If

'2010.10.28 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    STANA_LOG_F = ""            '���ޒI���װ۸�̧�ٖ��� �ر�i=۸ޏo�͖����j
    STANA_LOG_Out_Msg = ""      '���ޒI���װ۸ޏo�͗L��ү���� �ر
'2010.10.28 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Show
    If List_Disp_Proc() Then
        Unload Me
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer

    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<   2012.12.22  �ǉ��i�I��F�������R�s�[�j
                                    '�h�m�h����F
                                                    '2016.01.07 P_SYS.INI PR00030.INI
            If WriteIni(App.EXEName, "STAT" & Trim(GLB_SYUSHI_F), App.EXEName, Format(STAT_F, "0")) Then
'            If WriteIni(App.EXEName, "STAT" & Trim(GLB_SYUSHI_F), "p_sys", Format(STAT_F, "0")) Then
                Beep
                MsgBox ("INI�t�@�C���̏������݂Ɏ��s���܂����B" & App.EXEName & "STAT")
                Unload Me
            End If
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   �����܂ŁB
            
            
            
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If

                                            '�݌��ް��b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌��ް�")
        End If
    End If

                                            '�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If



                                            '�R�[�h�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�R�[�h�}�X�^")
        End If
    End If



                                            '���ޒ����b�k�n�r�d
    sts = BTRV(BtOpClose, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ޒ���")
        End If
    End If

                                            '���ޑO�؂b�k�n�r�d
    sts = BTRV(BtOpClose, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ޑO��")
        End If
    End If

                                            '���ގ���b�k�n�r�d
    sts = BTRV(BtOpClose, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K0_P_SHUKEIRE, Len(K0_P_SHUKEIRE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ގ��")
        End If
    End If

                                            '���ޒI�����b�k�n�r�d
    sts = BTRV(BtOpClose, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ޒI��")
        End If
    End If
                                            '���ޒI�����b�k�n�r�d
    sts = BTRV(BtOpClose, wP_STOCK_POS, wP_STOCK_REC, Len(wP_STOCK_REC), K0_wP_STOCK, Len(K0_wP_STOCK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "w���ޒI��")
        End If
    End If
                                            '���ޒI���W�v�b�k�n�r�d
    sts = BTRV(BtOpClose, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ޒI���W�v")
        End If
    End If

'2010.10.28 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                                            '���ޒI���i�ڃ}�X�^�ۑ��b�k�n�r�d
    sts = BTRV(BtOpClose, T_ITEMSV_POS, T_ITEMSVREC, Len(T_ITEMSVREC), K0_T_ITEMSV, Len(K0_T_ITEMSV), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ޒI���i�ڃ}�X�^�ۑ�")
        End If
    End If
'2010.10.28 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<



                                            '�Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�Ǘ��}�X�^")
        End If
    End If

    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PR000301 = Nothing
    Set PR000302 = Nothing
    Set PR00030F1 = Nothing
    Set PR00030F2 = Nothing


    End
End Sub

Private Sub TDBGrid1_FetchRowStyle(Index As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid80.StyleDisp)

    If STOCK(Bookmark, colZEN_ZAIKO_QTY) <> STOCK(Bookmark, colSAV_ZEN_ZAIKO) Then
        RowStyle = "Rstyle_Red"
    Else
        RowStyle = "Normal"
    End If

End Sub

Private Sub TDBGrid1_HeadClick(Index As Integer, ByVal ColIndex As Integer)

    Select Case Index

        Case pGridSTOCK

            If Sort_Tbl(ColIndex) = 0 Then
                Sort_Tbl(ColIndex) = 1
            Else
                If Sort_Tbl(ColIndex) = 1 Then
                    Sort_Tbl(ColIndex) = 0
                End If

            End If

            If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then

                STOCK.QuickSort Min_Row, STOCK.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING

                Set TDBGrid1(Index).Array = STOCK

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


    If Error_Check_Proc(Index, 1) Then   '�G���[�`�F�b�N
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



    For i = ptxSTART_DT To ptxLAST_SHIME_DT
        Text1(i).Text = ""
    Next i
    '����N��������



    If Trim(START_DT) = "" Then
        START_DT = Format(Now, "YYYY/MM/DD")
    End If


    Text1(ptxSTART_DT).Text = START_DT
    Text1(ptxKEIJYO_YM).Text = Left(START_DT, 7)


    Text1(ptxLAST_START_DT).Text = LAST_START_DT
    Text1(ptxLAST_SHIME_DT).Text = LAST_SHIME_DT



    '��ď��̏�����

    '��ď��̏�����
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0               '��̫�ď���
    Next i

    Sort_Tbl(colHIN_NAME) = 9       '��ď��O

    Init_Proc = False

End Function

Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'           ���ޒI�����ް��̕\��
'----------------------------------------------------------------------------
Dim sts                 As Integer
Dim com                 As Integer

Dim ROW                 As Long

Dim SKIP_FLG            As Boolean

Dim i                   As Integer
Dim Mode                As Integer

Dim ZEN_ZAIKO           As Long

Dim svG_SYUSHI          As String * 3   '2006.11.22
Dim svJGYOBU            As String * 1   '2006.11.22
Dim svNAIGAI            As String * 1   '2006.11.22
Dim svHIN_GAI           As String * 20  '2006.11.22

Dim c               As String * 128     '2006.11.22
Dim FileName        As String           '2006.11.22

Dim yn                  As Integer      '2010.12.20



    List_Disp_Proc = True
    PR000301.MousePointer = vbHourglass


                                'tmp�I�����t�@�C����荞��
    If GetIni("FILE", tmpP_STOCK_ID, "SYS", c) Then
        Beep
        MsgBox "tmp���ނ̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        Exit Function
    End If


'2010.12.20 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    FileName = Trim(c)
''    On Error Resume Next
''    Kill (fileName)
''    On Error GoTo 0"tmp���ޒI�����ް�"

    On Error GoTo List_Disp_Proc_Error
    Kill (FileName)
'2010.12.20 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


    If tmpP_STOCK_Open(0) Then
        Exit Function
    End If







'2011.02.14 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    com = BtOpGetFirst
    
    
    Do
        DoEvents
    
        sts = BTRV(com, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), K0_tmpP_STOCK, Len(K0_tmpP_STOCK), 0)
        
        Select Case sts
            Case BtNoErr
            
                sts = BTRV(BtOpDelete, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), K0_tmpP_STOCK, Len(K0_tmpP_STOCK), 0)
                If sts Then
                    Call File_Error(sts, BtOpDelete, "tmp���ޒI���W�v�ް�")
                    Exit Function
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "tmp���ޒI���W�v�ް�")
                Exit Function
        End Select
    
    Loop
    




'2011.02.14 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<










    If tmpP_STOCK_MAKE_Proc() Then
        Exit Function
    End If


                                'w���ޒI�����n�o�d�m
'    If wP_STOCK_Open(BtOpenNomal) Then
'        Unload Me
'    End If


    com = BtOpGetFirst

    Do
        sts = BTRV(com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)
        Select Case sts
            Case BtNoErr
                If CDbl(StrConv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, vbUnicode)) = 0 And _
                   CDbl(StrConv(P_STOCKSUM_REC.NYUKO_KIN, vbUnicode)) = 0 And _
                   CDbl(StrConv(P_STOCKSUM_REC.SYUKO_KIN, vbUnicode)) = 0 And _
                   CDbl(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode)) = 0 Then
                    sts = BTRV(BtOpDelete, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)
                    Select Case sts
                        Case BtNoErr
                        Case Else
                            Call File_Error(sts, BtOpDelete, "���ޒI���W�v�ް�")
                            Exit Function
                    End Select
                End If

            Case BtErrEOF
                Exit Do

            Case Else
                Call File_Error(sts, com, "���ޒI���W�v�ް�")
                Exit Function

        End Select

        com = BtOpGetNext

    Loop


    '���x�P��
    Set Sum_STOCK = Nothing

    ROW = Sum_Min_Row - 1

    com = BtOpGetFirst
    Do
        DoEvents

        sts = BTRV(com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)


        Select Case sts
            Case BtNoErr

            Case BtErrEOF
                Exit Do

            Case Else
                Call File_Error(sts, com, "���ޒI�����W�v�ް�")
                Exit Function

        End Select


        ROW = ROW + 1

        Sum_STOCK.ReDim Sum_Min_Row, ROW, Sum_Min_Col, Sum_Max_Col


        '�i��
        Sum_STOCK(ROW, colSum_G_SYUSHI) = StrConv(P_STOCKSUM_REC.G_SYUSHI, vbUnicode)
        '�O���݌ɋ��z   CLng--> Val 2016.01.08
        Sum_STOCK(ROW, colSum_ZEN_ZAIKO_KIN) = Format(Val(StrConv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, vbUnicode)), "#,##0")
        '�������ɋ��z   CLng--> Val 2016.01.08
        Sum_STOCK(ROW, colSum_NYUKO_KIN) = Format(Val(StrConv(P_STOCKSUM_REC.NYUKO_KIN, vbUnicode)), "#,##0")
        '�����o�ɋ��z   CLng--> Val 2016.01.08
        Sum_STOCK(ROW, colSum_SYUKO_KIN) = Format(Val(StrConv(P_STOCKSUM_REC.SYUKO_KIN, vbUnicode)), "#,##0")
        '�����݌ɋ��z   CLng--> Val 2016.01.08
        Sum_STOCK(ROW, colSum_ZAIKO_KIN) = Format(Val(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode)), "#,##0")
        '�������z
        Sum_STOCK(ROW, colSum_SA_KIN) = Format(Sum_STOCK(ROW, colSum_ZAIKO_KIN) - Sum_STOCK(ROW, colSum_ZEN_ZAIKO_KIN), "#,##0")


        com = BtOpGetGreater
    Loop


    Set TDBGrid1(pSum_GridSTOCK).Array = Sum_STOCK
    TDBGrid1(pSum_GridSTOCK).ReBind
    TDBGrid1(pSum_GridSTOCK).Update
    TDBGrid1(pSum_GridSTOCK).MoveFirst



'2010.10.28 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    ZAIKO_MINUS_Msg = ""        '�݌�ϲŽү���� �ر
    BEF_Hin_GAI = ""
'2010.10.28 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


    Set STOCK = Nothing

    ROW = Min_Row - 1

    com = BtOpGetFirst

    Do
        DoEvents

        sts = BTRV(com, tmpP_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K2_tmpP_STOCK, Len(K2_tmpP_STOCK), 2)


        Select Case sts
            Case BtNoErr


            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���ޒI�����ް�")
                Exit Function
        End Select




        ROW = ROW + 1
        If Grid_Set_Proc(ROW, ZEN_ZAIKO) Then
            Exit Function
        End If




        com = BtOpGetGreater
    Loop


    Set TDBGrid1(pGridSTOCK).Array = STOCK
    TDBGrid1(pGridSTOCK).ReBind
    TDBGrid1(pGridSTOCK).Update
    TDBGrid1(pGridSTOCK).MoveFirst

    DoEvents


    '2006.11.22 ��
                                            'tmp���ޒI�����b�k�n�r�d
    sts = BTRV(BtOpClose, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), K0_tmpP_STOCK, Len(K0_tmpP_STOCK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "tmp���ޒI��")
        End If
    End If

                                            '���ޒI�����b�k�n�r�d
    sts = BTRV(BtOpClose, wP_STOCK_POS, wP_STOCK_REC, Len(wP_STOCK_REC), K0_wP_STOCK, Len(K0_wP_STOCK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "w���ޒI��")
        End If
    End If


'2010.10.28 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    '�݌�ϲŽү���ފm�F�@�����ˈ�����ݎg�p��
    '�@�@�@�@�@�@�@�@�@�@�L��ˈ�����ݎg�p�s��
    If ZAIKO_MINUS_Msg = "" Then
        Command1(P_CMD_PRT).Enabled = True
    Else
        Command1(P_CMD_PRT).Enabled = False
        MsgBox ZAIKO_MINUS_Msg, vbExclamation, "�x��"
    End If
'2010.10.28 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


    PR000301.MousePointer = vbDefault
    List_Disp_Proc = False

    Exit Function



'2010.12.20 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
List_Disp_Proc_Error:
    If Err.Number = 70 Then
        yn = MsgBox("���[���Ŏ��ޒI�����W�v���ׁ̈A���s�ł��܂���" & vbCr & vbLf & _
                    "�Ď��s���܂����H", vbOKCancel + vbExclamation, Err.Source)

        If yn = vbOK Then
            Resume
        End If
    Else
        If Err.Number = 53 Then
            Resume Next
        Else
            MsgBox "[" & Err.Number & "] " & Err.Description, vbOKCancel + vbExclamation, Err.Source
        End If
    End If
'2010.12.20 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Function

Private Function Grid_Set_Proc(ROW As Long, ZEN_ZAIKO As Long) As Integer
'----------------------------------------------------------------------------
'           ���ޒI�����ް��̓��e���د�ނɾ�Ă���
'----------------------------------------------------------------------------
Dim sts                 As Integer

Dim com                 As Integer
Dim Save_Jgyobu         As String
Dim Save_Naigai         As String
Dim Save_Hin_Gai        As String
Dim Save_G_Syushi       As String


Dim wkSAKI_SHIIRE       As Long             '2017.04.22



On Error Resume Next

Label3(0).Caption = "��ʕ\��"


    Grid_Set_Proc = True

    STOCK.ReDim Min_Row, ROW, Min_Col, Max_Col


    '�i��
    STOCK(ROW, colHIN_GAI) = StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)




    '�i��
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_STOCK_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_STOCK_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_GAI, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    End Select
    STOCK(ROW, colHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)

Label3(1).Caption = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))



'2010.10.28 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    '���ޒI���W�v�Ώەi�ڃ}�X�^�ޔ�
    Call UniCode_Conv(K0_T_ITEMSV.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_T_ITEMSV.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_T_ITEMSV.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))

    sts = BTRV(BtOpInsert, T_ITEMSV_POS, ITEMREC, Len(T_ITEMSVREC), K0_T_ITEMSV, Len(K0_T_ITEMSV), 0)
    Select Case sts
        Case BtNoErr, BtErrDuplicates

        Case Else
            Call File_Error(sts, BtOpInsert, "���ޒI���i�ڃ}�X�^�ۑ�")
    End Select


    '�d���P�� �ݒ�L���`�F�b�N
    If Not IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)) Then
'        STANA_LOG_Out_Msg = "���ƕ��F" & StrConv(ITEMREC.JGYOBU, vbUnicode) & ", " & _
'                            "�����O�F" & StrConv(ITEMREC.NAIGAI, vbUnicode) & ", " & _
'                            "�O���i�ԁF" & StrConv(ITEMREC.HIN_GAI, vbUnicode) & ", " & _
'                            "�i���F" & StrConv(ITEMREC.HIN_NAME, vbUnicode) & ", " & _
'                            "�u�d���P�����ݒ肳��Ă��܂���v"
        STANA_LOG_Out_Msg = "�O���i�ԁF" & StrConv(ITEMREC.HIN_GAI, vbUnicode) & ", " & _
                            "�i���F" & StrConv(ITEMREC.HIN_NAME, vbUnicode) & ", " & _
                            "�u�d���P�����ݒ肳��Ă��܂���v"
        Call STANA_ErrLogPut(STANA_LOG_Out_Msg)
    End If


    '���x�`�F�b�N�i�i�ځF���ޒ����j
    Call UniCode_Conv(K1_P_SHORDER.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K1_P_SHORDER.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K1_P_SHORDER.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
    Call UniCode_Conv(K1_P_SHORDER.ORDER_DT, Format(Text1(1).Text & "/01", "yyyymmdd"))
    Call UniCode_Conv(K1_P_SHORDER.ORDER_NO, "")

    com = BtOpGetGreaterEqual
    Do
        sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K1_P_SHORDER, Len(K1_P_SHORDER), 1)

        Select Case sts
            Case BtNoErr
                If StrConv(P_SHORDER_REC.JGYOBU, vbUnicode) <> StrConv(ITEMREC.JGYOBU, vbUnicode) Or _
                   StrConv(P_SHORDER_REC.NAIGAI, vbUnicode) <> StrConv(ITEMREC.NAIGAI, vbUnicode) Or _
                   StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode) <> StrConv(ITEMREC.HIN_GAI, vbUnicode) Then
                    Exit Do
                End If

                If StrConv(P_SHORDER_REC.G_SYUSHI, vbUnicode) <> StrConv(ITEMREC.G_SYUSHI, vbUnicode) Then
'                    STANA_LOG_Out_Msg = "���ƕ��F" & StrConv(ITEMREC.JGYOBU, vbUnicode) & ", " & _
'                                        "�����O�F" & StrConv(ITEMREC.NAIGAI, vbUnicode) & ", " & _
'                                        "�O���i�ԁF" & StrConv(ITEMREC.HIN_GAI, vbUnicode) & ", " & _
'                                        "�i���F" & StrConv(ITEMREC.HIN_NAME, vbUnicode) & ", " & _
'                                        "�u���x�P�ʂ���v���܂���i�i�ځ����ޒ����j�v"
'                    STANA_LOG_Out_Msg = "�O���i�ԁF" & StrConv(ITEMREC.HIN_GAI, vbUnicode) & ", " & _
'                                        "�i���F" & StrConv(ITEMREC.HIN_NAME, vbUnicode) & ", " & _
'                                        "�u���x�P�ʂ���v���܂���i�i�ځ����ޒ����j�v"
                    
                    
                    STANA_LOG_Out_Msg = "�����ԍ�=" & StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode) & "-001" & "," & "�O���i�ԁF" & StrConv(ITEMREC.HIN_GAI, vbUnicode) & ", " & _
                                        "�i���F" & StrConv(ITEMREC.HIN_NAME, vbUnicode) & ", " & _
                                        "�u���x�P�ʂ���v���܂���i�i�ځ����ޒ����j�v"
                    
                    
                    Call STANA_ErrLogPut(STANA_LOG_Out_Msg)
                    Exit Do
                End If

            Case BtErrEOF
                Exit Do

            Case Else
                Call File_Error(sts, BtOpGetEqual, "���ޒ����ް�")
                Exit Function

        End Select

        com = BtOpGetNext

    Loop
'2010.10.28 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<




    '�݌Ɍ��i���x�j
    STOCK(ROW, colG_SYUSHI) = StrConv(P_STOCK_REC.G_SYUSHI, vbUnicode)


    '�������ɐ�
    STOCK(ROW, colNYUKO_QTY) = Format(Val(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)), "#,##0")
    '�����o�ɐ�
    STOCK(ROW, colSYUKO_QTY) = Format(Val(StrConv(P_STOCK_REC.SYUKO_QTY, vbUnicode)), "#,##0")
    '�����݌ɐ�
    STOCK(ROW, colZAIKO_QTY) = Format(Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)), "#,##0")

'2010.10.28 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    If Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) < 0 Then
        ZAIKO_MINUS_Msg = "�����݌Ƀ}�C�i�X�̕i�ڂ��L��܂��i�W�v�\����s�j"
    End If
'2010.10.28 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    '�d���P��
    If IsNumeric(StrConv(P_STOCK_REC.TANKA, vbUnicode)) Then
        STOCK(ROW, colSHI_TANKA) = Format(CDbl(StrConv(P_STOCK_REC.TANKA, vbUnicode)), "#,##0.00")
    Else
        STOCK(ROW, colSHI_TANKA) = ""
    End If
    '�d����
    STOCK(ROW, colSHI_CODE) = StrConv(P_STOCK_REC.CODE, vbUnicode)




    If IsNumeric(StrConv(P_STOCK_REC.ZEN_ZAIKO_KIN, vbUnicode)) Then
        STOCK(ROW, colZEN_ZAIKO_KIN) = Format(CCur(StrConv(P_STOCK_REC.ZEN_ZAIKO_KIN, vbUnicode)), "#,##0")
    Else
        STOCK(ROW, colZEN_ZAIKO_KIN) = ""
    End If



    '�݌ɋ��z

    If IsNumeric(StrConv(P_STOCK_REC.TANKA, vbUnicode)) Then
'        STOCK(Row, colZAIKO_KIN) = Format(CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) * _
'                                    CDbl(StrConv(P_STOCK_REC.TANKA, vbUnicode)), "#,##0")
        
        If Not IsNumeric(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) Then
            Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, "00000000")
        End If
        If Not IsNumeric(StrConv(P_STOCK_REC.TANKA, vbUnicode)) Then
            Call UniCode_Conv(P_STOCK_REC.TANKA, "00000000")
        End If
            
        STOCK(ROW, colZAIKO_KIN) = Format(ToRoundUp(CCur(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) * _
                                    CCur(StrConv(P_STOCK_REC.TANKA, vbUnicode)), 0), "#,##0")


    Else
        STOCK(ROW, colZAIKO_KIN) = ""
    End If



    STOCK(ROW, colSA_ZAIKO_KIN) = Format(CLng(STOCK(ROW, colZEN_ZAIKO_KIN)) - CLng(STOCK(ROW, colZAIKO_KIN)), "#,##0")



    '�O���݌� 2006.11.22
    STOCK(ROW, colZEN_ZAIKO_QTY) = Format(Val(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)), "#,##0")
'    STOCK(Row, colZEN_ZAIKO_QTY) = Format(ZEN_ZAIKO, "#,##0")

    '�ŏI�o�ד�
    STOCK(ROW, colLAST_SYUKA_DT) = Mid(StrConv(P_STOCK_REC.LAST_SYUKA_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_STOCK_REC.LAST_SYUKA_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_STOCK_REC.LAST_SYUKA_DT, vbUnicode), 7, 2)
    '�ŏI�o�ɐ�
    If IsNumeric(StrConv(P_STOCK_REC.LAST_SYUKA_QTY, vbUnicode)) Then
        STOCK(ROW, colLAST_SYUKA_QTY) = Format(CLng(StrConv(P_STOCK_REC.LAST_SYUKA_QTY, vbUnicode)), "#,##0")
    Else
        STOCK(ROW, colLAST_SYUKA_QTY) = ""
    End If



'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    ����t���ɕ�    2017.04.22

    If SAKI_SHIIRE_Proc(ROW, wkSAKI_SHIIRE) Then
        Exit Function
    End If
    STOCK(ROW, colSAKI_SHIIRE) = Format(wkSAKI_SHIIRE, "#,##0")
'>>>>>  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    ����t���ɕ�    2017.04.22
    
    
    
    '�O�ؐ�
    If IsNumeric(StrConv(P_STOCK_REC.MAEGARI_QTY, vbUnicode)) Then
        STOCK(ROW, colMAEGARI_QTY) = Format(CLng(StrConv(P_STOCK_REC.MAEGARI_QTY, vbUnicode)), "#,##0")
    Else
        STOCK(ROW, colMAEGARI_QTY) = ""
    End If

    '�����݌ɐ�
    If IsNumeric(StrConv(P_STOCK_REC.MOTO_ZAIKO_QTY, vbUnicode)) Then
        STOCK(ROW, colMOTO_ZAIKO_QTY) = Format(CLng(StrConv(P_STOCK_REC.MOTO_ZAIKO_QTY, vbUnicode)), "#,##0")
    Else
        STOCK(ROW, colMOTO_ZAIKO_QTY) = ""
    End If

'2010.10.28 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    '�����O���c
    If BEF_Hin_GAI <> StrConv(P_STOCK_REC.HIN_GAI, vbUnicode) Then
        Call UniCode_Conv(K0_T_ITEMSV.JGYOBU, StrConv(P_STOCK_REC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_T_ITEMSV.NAIGAI, StrConv(P_STOCK_REC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_T_ITEMSV.HIN_GAI, StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))

        sts = BTRV(BtOpGetEqual, T_ITEMSV_POS, T_ITEMSVREC, Len(T_ITEMSVREC), K0_T_ITEMSV, Len(K0_T_ITEMSV), 0)
        Select Case sts
            Case BtNoErr, BtErrDuplicates

            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "���ޒI���i�ڃ}�X�^�ۑ�", 0)
        End Select

        If IsNumeric(StrConv(T_ITEMSVREC.G_ZEN_ZAIKO_QTY, vbUnicode)) Then
            STOCK(ROW, colSAV_ZEN_ZAIKO) = Format(Val(StrConv(T_ITEMSVREC.G_ZEN_ZAIKO_QTY, vbUnicode)), "#,0")
        Else
''            STOCK(Row, colSAV_ZEN_ZAIKO) = ""
            STOCK(ROW, colSAV_ZEN_ZAIKO) = "0"          '2010.12.17 Upd
        End If

    Else
        '����O���i�Ԃ��A������ꍇ�A�Q�s�ڈȍ~�͑O���c�Ɠ����l���Z�b�g
        STOCK(ROW, colSAV_ZEN_ZAIKO) = STOCK(ROW, colZEN_ZAIKO_QTY)

    End If

    BEF_Hin_GAI = StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)
'2010.10.28 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    '���͓�
    STOCK(ROW, colINPUT_DATE) = StrConv(P_STOCK_REC.INPUT_DATE, vbUnicode)

    If IsNumeric(StrConv(P_STOCK_REC.FILLER, vbUnicode)) Then
        STOCK(ROW, 18) = Format(Val(StrConv(P_STOCK_REC.FILLER, vbUnicode)), "#")
    End If
On Error GoTo 0


    Grid_Set_Proc = False

    Exit Function



    Call UniCode_Conv(K2_wP_STOCK.G_SYUSHI, StrConv(P_STOCK_REC.G_SYUSHI, vbUnicode))

    Call UniCode_Conv(K2_wP_STOCK.JGYOBU, StrConv(P_STOCK_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K2_wP_STOCK.NAIGAI, StrConv(P_STOCK_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K2_wP_STOCK.HIN_GAI, StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))
    Call UniCode_Conv(K2_wP_STOCK.INPUT_DATE, StrConv(P_STOCK_REC.INPUT_DATE, vbUnicode))
    Call UniCode_Conv(K2_wP_STOCK.CODE, StrConv(P_STOCK_REC.CODE, vbUnicode))
    Call UniCode_Conv(K2_wP_STOCK.TANKA, StrConv(P_STOCK_REC.TANKA, vbUnicode))


    Save_G_Syushi = StrConv(P_STOCK_REC.G_SYUSHI, vbUnicode)

    Save_Jgyobu = StrConv(P_STOCK_REC.JGYOBU, vbUnicode)
    Save_Naigai = StrConv(P_STOCK_REC.NAIGAI, vbUnicode)
    Save_Hin_Gai = StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)

    com = BtOpGetGreater


    Do
        DoEvents
        sts = BTRV(com, wP_STOCK_POS, wP_STOCK_REC, Len(wP_STOCK_REC), K2_wP_STOCK, Len(K2_wP_STOCK), 2)

        Select Case sts
            Case BtNoErr
                If Save_Jgyobu <> StrConv(wP_STOCK_REC.JGYOBU, vbUnicode) Or _
                    Save_Naigai <> StrConv(wP_STOCK_REC.NAIGAI, vbUnicode) Or _
                    Save_Hin_Gai <> StrConv(wP_STOCK_REC.HIN_GAI, vbUnicode) Or _
                    Save_G_Syushi <> StrConv(wP_STOCK_REC.G_SYUSHI, vbUnicode) Then
                    Exit Do
                End If

            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "w���ޒI���ް�")
                Exit Function
        End Select



        ROW = ROW + 1
'        If i > 0 Then
            STOCK.ReDim Min_Row, ROW, Min_Col, Max_Col

            STOCK(ROW, colHIN_GAI) = ""
            STOCK(ROW, colHIN_NAME) = ""
            STOCK(ROW, colG_SYUSHI) = ""
            STOCK(ROW, colZEN_ZAIKO_QTY) = ""           '--->2006.11.22 ����


            STOCK(ROW, colLAST_SYUKA_DT) = ""
            STOCK(ROW, colLAST_SYUKA_QTY) = ""


'        End If
        '�O���݌�   2006.11.22
'''        STOCK(Row, colZEN_ZAIKO_QTY) = Format(CLng(StrConv(wP_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)), "#,##0")
        '�������ɐ� Clng --> Val 2016.01.08
        STOCK(ROW, colNYUKO_QTY) = Format(Val(StrConv(wP_STOCK_REC.NYUKO_QTY, vbUnicode)), "#,##0")

        '�����o�ɐ� Clng --> Val 2016.01.08
        STOCK(ROW, colSYUKO_QTY) = Format(Val(StrConv(wP_STOCK_REC.SYUKO_QTY, vbUnicode)), "#,##0")

        '�����݌ɐ� Clng --> Val 2016.01.08
        STOCK(ROW, colZAIKO_QTY) = Format(Val(StrConv(wP_STOCK_REC.ZAIKO_QTY, vbUnicode)), "#,##0")

        '�d���P��
        If IsNumeric(StrConv(wP_STOCK_REC.TANKA, vbUnicode)) Then
            STOCK(ROW, colSHI_TANKA) = Format(CDbl(StrConv(wP_STOCK_REC.TANKA, vbUnicode)), "#,##0.00")
        Else
            STOCK(ROW, colSHI_TANKA) = ""
        End If

        '�d����
        STOCK(ROW, colSHI_CODE) = StrConv(wP_STOCK_REC.CODE, vbUnicode)

        '�݌ɋ��z
        If IsNumeric(StrConv(wP_STOCK_REC.TANKA, vbUnicode)) Then
            STOCK(ROW, colZAIKO_KIN) = Format(CLng(StrConv(wP_STOCK_REC.ZAIKO_QTY, vbUnicode)) * _
                                        CDbl(StrConv(wP_STOCK_REC.TANKA, vbUnicode)), "#,##0")
        Else
            STOCK(ROW, colZAIKO_KIN) = ""
        End If

        '�O�؎c
        If IsNumeric(StrConv(P_STOCK_REC.MAEGARI_QTY, vbUnicode)) Then
            STOCK(ROW, colMAEGARI_QTY) = Format(CLng(StrConv(P_STOCK_REC.MAEGARI_QTY, vbUnicode)), "#,##0")
        Else
            STOCK(ROW, colMAEGARI_QTY) = ""
        End If

        '�����݌ɐ�
        If IsNumeric(StrConv(P_STOCK_REC.MOTO_ZAIKO_QTY, vbUnicode)) Then
            STOCK(ROW, colMOTO_ZAIKO_QTY) = Format(CLng(StrConv(P_STOCK_REC.MOTO_ZAIKO_QTY, vbUnicode)), "#,##0")
        Else
            STOCK(ROW, colMOTO_ZAIKO_QTY) = ""
        End If

        '���͓�
        STOCK(ROW, colINPUT_DATE) = StrConv(P_STOCK_REC.INPUT_DATE, vbUnicode)


    If IsNumeric(StrConv(P_STOCK_REC.FILLER, vbUnicode)) Then
        STOCK(ROW, 18) = Format(Val(StrConv(P_STOCK_REC.FILLER, vbUnicode)), "#")
    End If
        
        com = BtOpGetNext


    Loop


'    Call UniCode_Conv(K2_tmpP_STOCK.JGYOBU, Save_Jgyobu)
'    Call UniCode_Conv(K2_tmpP_STOCK.NAIGAI, Save_Naigai)
'    Call UniCode_Conv(K2_tmpP_STOCK.HIN_GAI, Save_Hin_Gai)
'    Call UniCode_Conv(K2_tmpP_STOCK.CODE, "zzzzzz")
'    Call UniCode_Conv(K2_tmpP_STOCK.TANKA, "zzzzzzzzzzzz")





    Grid_Set_Proc = False


On Error GoTo 0

End Function

Private Function Print_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'           �������
'
'   mode 1:�W�v�\�̈��
'        2:
'----------------------------------------------------------------------------

Dim Data_Flg        As Boolean


Dim rpt1            As New PR00030F1
Dim rpt2            As New PR00030F2

Dim f               As New PR000302

Dim c               As String * 128     '2006.11.22
Dim FileName        As String           '2006.11.22

Dim sts             As Integer          '2006.11.22

Dim com             As Integer
    

Dim yn              As Integer

    Print_Proc = True

    Select Case Mode
        Case 1

             '�W�v�\���
             Set rpt1 = New PR00030F1

             '���|�[�g��������܂��B�itrue�F����_�C�A���O���� false�F�Ȃ��j
             rpt1.PrintReport False

             Set rpt1 = Nothing

        Case 2


            '2006.11.22 ��
                                        'tmp�I�����t�@�C����荞��
            If GetIni("FILE", tmpP_STOCK_ID, "SYS", c) Then
                Beep
                MsgBox "tmp���ނ̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
                Exit Function
            End If

            FileName = Trim(c)
''            On Error Resume Next
''            Kill (fileName)
''            On Error GoTo 0




            On Error GoTo Print_Proc_Error
            Kill (FileName)




            If tmpP_STOCK_Open(0) Then
                Exit Function
            End If



'2011.02.14 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            com = BtOpGetFirst
            
            
            Do
                DoEvents
            
                sts = BTRV(com, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), K0_tmpP_STOCK, Len(K0_tmpP_STOCK), 0)
                
                Select Case sts
                    Case BtNoErr
                    
                        sts = BTRV(BtOpDelete, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), K0_tmpP_STOCK, Len(K0_tmpP_STOCK), 0)
                        If sts Then
                            Call File_Error(sts, BtOpDelete, "tmp���ޒI���W�v�ް�")
                            Exit Function
                        End If
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, com, "tmp���ޒI���W�v�ް�")
                        Exit Function
                End Select
            
            Loop
    




'2011.02.14 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<



            If tmpP_STOCK_MAKE_Proc() Then
                Exit Function
            End If
            '2006.11.22 ��



             '���ו\���
             Set rpt2 = New PR00030F2

             '���|�[�g��������܂��B�itrue�F����_�C�A���O���� false�F�Ȃ��j
             rpt2.PrintReport False

             Set rpt2 = Nothing


        '     f.RunReport rpt
        '     f.Show

            '2006.11.22 ��
                                                    'tmp���ޒI�����b�k�n�r�d
            sts = BTRV(BtOpClose, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), K0_tmpP_STOCK, Len(K0_tmpP_STOCK), 0)
            If sts Then
                If sts <> BtErrNoOpen Then
                    Call File_Error(sts, BtOpClose, "tmp���ޒI��")
                End If
            End If
        '2006.11.22 ��
    End Select

    Print_Proc = False
    Exit Function

'2010.12.20 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Print_Proc_Error:
    If Err.Number = 70 Then
        yn = MsgBox("���[���Ŏ��ޒI�����W�v���ׁ̈A���s�ł��܂���" & vbCr & vbLf & _
                    "�Ď��s���܂����H", vbOKCancel + vbExclamation, Err.Source)

        If yn = vbOK Then
            Resume
        End If
    Else
        If Err.Number = 53 Then
            Resume Next
        Else
            MsgBox "[" & Err.Number & "] " & Err.Description, vbOKCancel + vbExclamation, Err.Source
        End If
    End If
'2010.12.20 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���ޒI�����ް��쐬
'----------------------------------------------------------------------------
Dim sts                     As Integer
Dim com                     As Integer

Dim Upd_Com                 As Integer

Dim i                       As Integer

Dim wkVal                   As Long
Dim wKEIJYO_YM              As String * 6

Dim SKIP_FLG                As Boolean

Dim wk_Val                  As Double

Dim GK_ZEN_ZAIKO_KIN        As Long
Dim GK_NYUKO_KIN            As Long
Dim GK_SYUKO_KIN            As Long
Dim GK_ZAIKO_KIN            As Long

Dim Save_Jgyobu             As String * 1
Dim Save_Naigai             As String * 1
Dim Save_Hin_Gai            As String * 20
Dim Save_CODE               As String * 5
Dim Save_TANKA              As String * 11

Dim Sum_NYUKA_QTY           As Long

Dim wkZEN_ZAIKO             As Long             '2006.11.22

Dim ZAIKO_F                 As Boolean          '2007.04.26
Dim Save_G_Syushi           As String * 3       '2007.04.26


Dim SYUSHI_ON               As Boolean          '2007.11.13
Dim Fast_Flg                As Boolean          '2007.11.13


Dim wkZaiko_QTY             As Long             '2008.06.21
Dim wkNYUKO_QTY             As Long             '2008.06.21
Dim Syuko_Non_Flg           As Boolean          '2008.06.21


Dim Next_Jgyobu             As String           '2008.06.21
Dim Next_Naigai             As String           '2008.06.21
Dim Next_Hin_Gai            As String           '2008.06.21


Dim Sumi_Zaiko_Qty          As Long
Dim Mi_Zaiko_Qty            As Long

Dim yn                      As Integer          '2016.01.07

Label3(0).Caption = "UPDATE START"

    Update_Proc = True
    PR000301.MousePointer = vbHourglass


'On Error GoTo Error_Proc



    com = BtOpGetFirst


    '�����ް��S���폜
    Do

        DoEvents

        sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

        Select Case sts
            Case BtNoErr


            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "���ޒI�����ް�")
                Exit Function
        End Select

        sts = BTRV(BtOpDelete, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

        Select Case sts
            Case BtNoErr

            Case Else
                Call File_Error(sts, BtOpDelete, "���ޒI�����ް�")
                Exit Function
        End Select


        com = BtOpGetNext

    Loop

    '�W�v�ް��S���폜

    com = BtOpGetFirst

    Do

        DoEvents

        sts = BTRV(com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

        Select Case sts
            Case BtNoErr


            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "���ޒI�����ް�")
                Exit Function
        End Select

        sts = BTRV(BtOpDelete, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

        Select Case sts
            Case BtNoErr

            Case Else
                Call File_Error(sts, BtOpDelete, "���ޒI�����ް�")
                Exit Function
        End Select

        com = BtOpGetNext

    Loop



Label3(0).Caption = "�O���c START"

    If ZenZan_Update_Proc() Then
        Exit Function
    End If




Label3(0).Caption = "�����d�� START"

    If SHIIRE_Update_Proc() Then
        Exit Function
    End If




Label3(0).Caption = "�����݌� START"

    '-------------------------------------  ���݂�蓖���c�݌ɂ��W�v
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, SHIZAI)
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, "")
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, "")
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, "")
    Call UniCode_Conv(K1_ZAIKO.SOKO_NO, "")
    Call UniCode_Conv(K1_ZAIKO.Retu, "")
    Call UniCode_Conv(K1_ZAIKO.Ren, "")
    Call UniCode_Conv(K1_ZAIKO.Dan, "")

    com = BtOpGetGreaterEqual

    Do

        DoEvents

        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)

        Select Case sts
            Case BtNoErr

                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> SHIZAI Then
                    Exit Do
                End If

            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "�݌��ް�")
                Exit Function
        End Select

        SKIP_FLG = False
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(ZAIKOREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(ZAIKOREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ZAIKOREC.HIN_GAI, vbUnicode))


        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

        Select Case sts
            Case BtNoErr

'2012.12.13                If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_OFF Then
                If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_OFF Or _
                    StrConv(ITEMREC.ZAIKO_CLR_F, vbUnicode) = "1" Then                          '2012.12.13
                    SKIP_FLG = True

                Else
                    If Not IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)) Then    '2008.02.13
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA, "00000000000")
                    End If

                    If Trim(StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode)) = "" Then
                        Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
                        '2008.11.24
                        Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, Format(CDbl(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))), "00000000.00"))
                    End If

                End If

Label3(1).Caption = Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode))

                SYUSHI_ON = False               '2007.11.13
                If GLB_SYUSHI_F = "" Then       '2007.11.13
                    SYUSHI_ON = True
                Else
                    SYUSHI_ON = False

                    For i = 0 To UBound(G_SYUSHI_TBL)

                        If Trim(StrConv(ITEMREC.G_SYUSHI, vbUnicode)) = G_SYUSHI_TBL(i) Then
                            SYUSHI_ON = True
                            Exit For
                        End If


                    Next i
                End If

            Case BtErrKeyNotFound
                SKIP_FLG = True

            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                Exit Function
        End Select


        If Not SKIP_FLG And SYUSHI_ON Then      '2007.11.13
            '����ں��ޏ���
            Call UniCode_Conv(K0_P_STOCK.JGYOBU, StrConv(ZAIKOREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_P_STOCK.NAIGAI, StrConv(ZAIKOREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_P_STOCK.HIN_GAI, StrConv(ZAIKOREC.HIN_GAI, vbUnicode))
            Call UniCode_Conv(K0_P_STOCK.CODE, StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode))
            
            If Not IsNumeric(StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode)) Then                '2015.12.28
                Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, Format(CDbl(0), "00000000.00"))    '2015.12.28
Call LOG_OUT(LOG_F, StrConv(ZAIKOREC.HIN_GAI, vbUnicode))
            Else                                                                            '2015.12.28
            '2008.11.24
            Call UniCode_Conv(K0_P_STOCK.TANKA, Format(CDbl(Trim(StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode))), "00000000.00"))
            End If                                                                  '2015.12.28
            sts = BTRV(BtOpGetEqual, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

            Select Case sts
                Case BtNoErr

                    Upd_Com = BtOpUpdate


                Case BtErrKeyNotFound

                    Upd_Com = BtOpInsert


                Case Else

                    Call File_Error(sts, BtOpGetEqual, "���ޒI�����ް�")
                    Exit Function

            End Select



            If Upd_Com = BtOpInsert Then
                Call UniCode_Conv(P_STOCK_REC.JGYOBU, StrConv(ZAIKOREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(P_STOCK_REC.NAIGAI, StrConv(ZAIKOREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(P_STOCK_REC.HIN_GAI, StrConv(ZAIKOREC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(P_STOCK_REC.CODE, StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode))
                '2008.11.24
                Call UniCode_Conv(P_STOCK_REC.TANKA, Format(CDbl(Trim(StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode))), "00000000.00"))
                '2006.11.22
                Call UniCode_Conv(P_STOCK_REC.INPUT_DATE, StrConv(ZAIKOREC.NYUKA_DT, vbUnicode))

                Call UniCode_Conv(P_STOCK_REC.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))
                Call UniCode_Conv(P_STOCK_REC.ZEN_ZAIKO_QTY, "00000000")



                Call UniCode_Conv(P_STOCK_REC.NYUKO_QTY, "00000000")
                Call UniCode_Conv(P_STOCK_REC.SYUKO_QTY, "00000000")
                Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, "00000000")


                Call UniCode_Conv(P_STOCK_REC.LAST_SYUKA_DT, StrConv(ITEMREC.LAST_SYU_DT, vbUnicode))
                Call UniCode_Conv(P_STOCK_REC.LAST_SYUKA_QTY, StrConv(ITEMREC.G_LAST_SYUKA_QTY, vbUnicode))


                Call UniCode_Conv(P_STOCK_REC.MOTO_ZAIKO_QTY, "00000000")
                Call UniCode_Conv(P_STOCK_REC.MAEGARI_QTY, "00000000")

                Call UniCode_Conv(P_STOCK_REC.SYUKA_NON_F, "0")     '2008.06.21


                Call UniCode_Conv(P_STOCK_REC.ZEN_ZAIKO_KIN, "00000000")


                Call UniCode_Conv(P_STOCK_REC.FILLER, "")



            End If
            '2006.11.22
            If StrConv(ZAIKOREC.NYUKA_DT, vbUnicode) < StrConv(P_STOCK_REC.INPUT_DATE, vbUnicode) Then
                Call UniCode_Conv(P_STOCK_REC.INPUT_DATE, StrConv(ZAIKOREC.NYUKA_DT, vbUnicode))
            End If


'2009.08.21            wk_VAL = CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) + _
'                        CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))


            wk_Val = Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) + _
                        Val(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))

            Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, Format(wk_Val, "00000000"))
'            Call UniCode_Conv(P_STOCK_REC.MOTO_ZAIKO_QTY, Format(wk_VAL, "00000000"))
            Call UniCode_Conv(P_STOCK_REC.MOTO_ZAIKO_QTY, StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))

Call UniCode_Conv(P_STOCK_REC.FILLER, Format(wk_Val, "000000"))


            Do
                sts = BTRV(Upd_Com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

                Select Case sts
                    Case BtNoErr, BtErrDuplicates
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        DoEvents
                    Case Else

                        Call File_Error(sts, BtOpGetEqual, "���ޒI�����ް�")
                        Exit Function
                End Select


            Loop


            Do
                sts = BTRV(BtOpUpdate, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)

                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        DoEvents
                    Case Else

                        Call File_Error(sts, BtOpUpdate, "�݌��ް�")
                        Exit Function
                End Select


            Loop

        End If

        com = BtOpGetNext

    Loop

Label3(0).Caption = "�����ȍ~�d���������݌ɂ������"


    '-------------------------------------  �����ȍ~�d���������݌ɂ������
    wKEIJYO_YM = Left(Text1(ptxKEIJYO_YM).Text, 4) & Right(Text1(ptxKEIJYO_YM).Text, 2)

    Call UniCode_Conv(K2_P_SHUKEIRE.KEIJYO_YM, wKEIJYO_YM)
    Call UniCode_Conv(K2_P_SHUKEIRE.UKEIRE_DT, "zzzzzzzz")

    com = BtOpGetGreater


    Do

        DoEvents

        sts = BTRV(com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K2_P_SHUKEIRE, Len(K2_P_SHUKEIRE), 2)

        Select Case sts
            Case BtNoErr


            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "���ގ���ް�")
                Exit Function
        End Select



        '�����ް��ǂݍ���
        Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
        SKIP_FLG = False
        Select Case sts
            Case BtNoErr
                
                
                
                '�i�ڂ̍݌Ɍv���׸ނ��`�F�b�N
                Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))

                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                SKIP_FLG = False
                Select Case sts
                    Case BtNoErr

'2012.12.13                        If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_OFF Then
                        If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_OFF Or _
                            StrConv(ITEMREC.ZAIKO_CLR_F, vbUnicode) = "1" Then                      '2012.12.13
                            SKIP_FLG = True       '�l�����Ȃ��̂Ž����
                        End If


                        SYUSHI_ON = False               '2007.11.13
                        If GLB_SYUSHI_F = "" Then       '2007.11.13
                            SYUSHI_ON = True
                        Else
                            SYUSHI_ON = False

                            For i = 0 To UBound(G_SYUSHI_TBL)

                                If Trim(StrConv(ITEMREC.G_SYUSHI, vbUnicode)) = G_SYUSHI_TBL(i) Then
                                    SYUSHI_ON = True
                                    Exit For
                                End If


                            Next i
                        End If


                    Case BtErrKeyNotFound


                        SKIP_FLG = True       '�l�����Ȃ��̂Ž����


                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                        Exit Function
                End Select




            Case BtErrKeyNotFound


                SKIP_FLG = True       '�����Ȃ��͒ʏ��ް��ł͂Ȃ�


            Case Else
                Call File_Error(sts, BtOpGetEqual, "���ޒ����ް�")
                Exit Function
        End Select


Label3(1).Caption = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))


        If SKIP_FLG Or Not SYUSHI_ON Then       '2007.11.13
        Else

            Call UniCode_Conv(K0_P_STOCK.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_P_STOCK.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_P_STOCK.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
            Call UniCode_Conv(K0_P_STOCK.CODE, StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode))
            '2008.11.24
            Call UniCode_Conv(K0_P_STOCK.TANKA, Format(CDbl(Trim(StrConv(P_SHUKEIRE_REC.UKEIRE_TANKA, vbUnicode))), "00000000.00"))
            sts = BTRV(BtOpGetEqual, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

            Select Case sts
                Case BtNoErr

                    Upd_Com = BtOpUpdate


                Case BtErrKeyNotFound

                    Upd_Com = BtOpInsert


                Case Else

                    Call File_Error(sts, BtOpGetEqual, "���ޒI�����ް�")
                    Exit Function

            End Select

            If Upd_Com = BtOpUpdate Then


'2008.11.13                wk_VAL = CLng(CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) - _
'2008.11.13                                CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)))
'2008.11.13


'2009.08.21                wk_VAL = CLng(CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) - _
'                                CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode)))

                wk_Val = Val(Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) - _
                                Val(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode)))



Call LOG_OUT(PR00030_LOG_F, "�d�� " & StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode) & " " & StrConv(P_SHUKEIRE_REC.KEIJYO_YM, vbUnicode) & " " & StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode))

'2008.11.13
'If Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)) = "C087" Then
'    Debug.Print
'    Call Log_Out(LOG_F, "B " & StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode) & StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode) & " " & StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))
'End If




                If wk_Val < 0 Then
                    Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, Format(wk_Val, "0000000"))
                Else
                    Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, Format(wk_Val, "00000000"))
                End If
                Call UniCode_Conv(P_STOCK_REC.MOTO_ZAIKO_QTY, StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))


                Do
                    sts = BTRV(Upd_Com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

                    Select Case sts
                        Case BtNoErr, BtErrDuplicates
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            DoEvents
                        Case Else

                            Call File_Error(sts, Upd_Com, "���ޒI�����ް�")
                            Exit Function
                    End Select


                Loop
            End If
        End If

        com = BtOpGetNext

    Loop


Label3(0).Caption = "�O�ؕ��𓥂܂��ē����c�݌ɂ��ďW�v"

    '-------------------------------------  �O�ؕ��𓥂܂��ē����c�݌ɂ��ďW�v
    com = BtOpGetFirst

    Fast_Flg = True     '2007.11.13

    Do

        DoEvents

        sts = BTRV(com, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)

        Select Case sts
            Case BtNoErr


                '2007.11.13 ��
                SYUSHI_ON = False

                Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_NYUREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_NYUREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_NYUREC.HIN_GAI, vbUnicode))

                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

                Select Case sts
                    Case BtNoErr

                        If GLB_SYUSHI_F = "" Then       '2007.11.13
                            SYUSHI_ON = True
                        Else
                            SYUSHI_ON = False

                            For i = 0 To UBound(G_SYUSHI_TBL)

                                If Trim(StrConv(ITEMREC.G_SYUSHI, vbUnicode)) = G_SYUSHI_TBL(i) Then
                                    SYUSHI_ON = True
                                    Exit For
                                End If


                            Next i
                        End If


                    Case BtErrKeyNotFound
                        SYUSHI_ON = True

                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                        Exit Function
                End Select
                '2007.11.13 ��


            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "���ޑO���ް�")
                Exit Function
        End Select



        If Not SYUSHI_ON Then       '2007.11.13
        Else

            If Not IsNumeric(StrConv(P_NYUREC.SHIIRE_TANKA, vbUnicode)) Then
                Call UniCode_Conv(P_NYUREC.SHIIRE_TANKA, "00000000.00")
            End If

            If com = BtOpGetFirst Then
                Save_Jgyobu = StrConv(P_NYUREC.JGYOBU, vbUnicode)
                Save_Naigai = StrConv(P_NYUREC.NAIGAI, vbUnicode)
                Save_Hin_Gai = StrConv(P_NYUREC.HIN_GAI, vbUnicode)
                Save_CODE = StrConv(P_NYUREC.SHIIRE_CODE, vbUnicode)
                '2008.11.24
                Save_TANKA = Format(CDbl(Trim(StrConv(P_NYUREC.SHIIRE_TANKA, vbUnicode))), "00000000.00")
                Sum_NYUKA_QTY = 0
            End If
            '2008.11.24


            If Save_Jgyobu <> StrConv(P_NYUREC.JGYOBU, vbUnicode) Or _
                Save_Naigai <> StrConv(P_NYUREC.NAIGAI, vbUnicode) Or _
                Save_Hin_Gai <> StrConv(P_NYUREC.HIN_GAI, vbUnicode) Or _
                Save_CODE <> StrConv(P_NYUREC.SHIIRE_CODE, vbUnicode) Or _
                Save_TANKA <> Format(CDbl(Trim(StrConv(P_NYUREC.SHIIRE_TANKA, vbUnicode))), "00000000.00") Then

                Call UniCode_Conv(K0_P_STOCK.JGYOBU, Save_Jgyobu)
                Call UniCode_Conv(K0_P_STOCK.NAIGAI, Save_Naigai)
                Call UniCode_Conv(K0_P_STOCK.HIN_GAI, Save_Hin_Gai)

                If Trim(Save_CODE) = "" Then
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_NYUREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_NYUREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_NYUREC.HIN_GAI, vbUnicode))

                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

                    Select Case sts
                        Case BtNoErr

                        Case BtErrKeyNotFound
                            '���肦�Ȃ�
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).CODE, "")
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA, "")

                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            Exit Function
                    End Select


                    Call UniCode_Conv(K0_P_STOCK.CODE, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
                    '2008.11.24
                    If IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)) Then
                        Call UniCode_Conv(K0_P_STOCK.TANKA, Format(CDbl(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))), "00000000.00"))
                    Else
                        Call UniCode_Conv(K0_P_STOCK.TANKA, "00000000.00")
                    End If
                Else
                    Call UniCode_Conv(K0_P_STOCK.CODE, Save_CODE)
                    Call UniCode_Conv(K0_P_STOCK.TANKA, Save_TANKA)
                End If


                sts = BTRV(BtOpGetEqual, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

                Select Case sts
                    Case BtNoErr

                        Upd_Com = BtOpUpdate


                    Case BtErrKeyNotFound

                        Upd_Com = BtOpInsert


                    Case Else

                        Call File_Error(sts, BtOpGetEqual, "���ޒI�����ް�")
                        Exit Function

                End Select



                If Upd_Com = BtOpUpdate Then


                    If Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) < Sum_NYUKA_QTY Then
                        Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, "00000000")
                    Else
                        Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, Format(Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) - Sum_NYUKA_QTY, "00000000"))
                    
                    
Call LOG_OUT(PR00030_LOG_F, "�O�� " & StrConv(P_STOCK_REC.HIN_GAI, vbUnicode) & " " & Format(Sum_NYUKA_QTY))
                    
                    
                    End If




                    Call UniCode_Conv(P_STOCK_REC.MOTO_ZAIKO_QTY, StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))


                    Call UniCode_Conv(P_STOCK_REC.MAEGARI_QTY, Format(Sum_NYUKA_QTY, "00000000"))

                    Do

                        sts = BTRV(BtOpUpdate, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

                        Select Case sts
                            Case BtNoErr
                                Exit Do

                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                DoEvents


                            Case Else
                                Call File_Error(sts, BtOpUpdate, "���ޒI���W�v�ް�")
                                Exit Function
                        End Select

                    Loop


'                    Save_Jgyobu = StrConv(P_NYUREC.JGYOBU, vbUnicode)
'                    Save_Naigai = StrConv(P_NYUREC.NAIGAI, vbUnicode)
'                    Save_Hin_Gai = StrConv(P_NYUREC.HIN_GAI, vbUnicode)
'                    Save_CODE = StrConv(P_NYUREC.SHIIRE_CODE, vbUnicode)
'                    '2008.11.24
'                    Save_TANKA = Format(CDbl(Trim(StrConv(P_NYUREC.SHIIRE_TANKA, vbUnicode))), "00000000.00")
'                    Sum_NYUKA_QTY = 0

                End If

                Save_Jgyobu = StrConv(P_NYUREC.JGYOBU, vbUnicode)
                Save_Naigai = StrConv(P_NYUREC.NAIGAI, vbUnicode)
                Save_Hin_Gai = StrConv(P_NYUREC.HIN_GAI, vbUnicode)
                Save_CODE = StrConv(P_NYUREC.SHIIRE_CODE, vbUnicode)
                '2008.11.24
                Save_TANKA = Format(CDbl(Trim(StrConv(P_NYUREC.SHIIRE_TANKA, vbUnicode))), "00000000.00")
                Sum_NYUKA_QTY = 0

            End If


            If IsNumeric(StrConv(P_NYUREC.SOUSAI_QTY, vbUnicode)) Then
'2009.08.21                Sum_NYUKA_QTY = Sum_NYUKA_QTY + CLng(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode)) - CLng(StrConv(P_NYUREC.SOUSAI_QTY, vbUnicode))


                Sum_NYUKA_QTY = Sum_NYUKA_QTY + Val(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode)) - Val(StrConv(P_NYUREC.SOUSAI_QTY, vbUnicode))

            Else
'2009.08.21                Sum_NYUKA_QTY = Sum_NYUKA_QTY + CLng(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode))

                Sum_NYUKA_QTY = Sum_NYUKA_QTY + Val(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode))

            End If

            Fast_Flg = False '2007.11.13

        End If      '2007.11.13

        com = BtOpGetNext

    Loop

'    If com <> BtOpGetFirst Then        2007.11.13
    If Not Fast_Flg Then                '2007.11.13
        Call UniCode_Conv(K0_P_STOCK.JGYOBU, Save_Jgyobu)
        Call UniCode_Conv(K0_P_STOCK.NAIGAI, Save_Naigai)
        Call UniCode_Conv(K0_P_STOCK.HIN_GAI, Save_Hin_Gai)

        If Trim(Save_CODE) = "" Then
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_NYUREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_NYUREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_NYUREC.HIN_GAI, vbUnicode))

            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

            Select Case sts
                Case BtNoErr

                Case BtErrKeyNotFound
                    '���肦�Ȃ�
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).CODE, "")
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Function
            End Select


            Call UniCode_Conv(K0_P_STOCK.CODE, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
            '2008.11.24

            If IsNumeric(ITEMREC.G_SHIIRE_TBL(0).TANKA) Then

                Call UniCode_Conv(K0_P_STOCK.TANKA, Format(CDbl(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))), "00000000.00"))

            Else

                Call UniCode_Conv(K0_P_STOCK.TANKA, "00000000.00")
            End If

        Else
            Call UniCode_Conv(K0_P_STOCK.CODE, Save_CODE)
            Call UniCode_Conv(K0_P_STOCK.TANKA, Save_TANKA)
        End If

        sts = BTRV(BtOpGetEqual, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

        Select Case sts
            Case BtNoErr

                Upd_Com = BtOpUpdate


            Case BtErrKeyNotFound

                Upd_Com = BtOpInsert


            Case Else

                Call File_Error(sts, BtOpGetEqual, "���ޒI�����ް�")
                Exit Function

        End Select



        If Upd_Com = BtOpUpdate Then
            'Clng --> Val 2016.01.08
            If Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) < Sum_NYUKA_QTY Then
                Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, "00000000")
            Else
'2009.08.21                Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, Format(CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) - Sum_NYUKA_QTY, "00000000"))
                Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, Format(Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) - Sum_NYUKA_QTY, "00000000"))
            End If


            Call UniCode_Conv(P_STOCK_REC.MAEGARI_QTY, Format(Sum_NYUKA_QTY, "00000000"))

            Call UniCode_Conv(P_STOCK_REC.MOTO_ZAIKO_QTY, StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))


            Do

                sts = BTRV(BtOpUpdate, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

                Select Case sts
                    Case BtNoErr
                        Exit Do

                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        DoEvents

                    Case Else
                        Call File_Error(sts, BtOpUpdate, "���ޒI���W�v�ް�")
                        Exit Function
                End Select

            Loop
        End If

    End If


    '----
    '-------------------------------------  �o�ɐ��̌v�Z

Label3(0).Caption = "�o�א��e�Z�b�g"

    If Syuka_F_Update_Proc() Then
        Exit Function
    End If


'------------------------------------------------------�@�o�א��e�Z�b�g 2008.06.21


 Label3(0).Caption = "�i�ڏW�v����"

    If Hin_Sum_Update_Proc() Then
        Exit Function
    End If


 Label3(0).Caption = "�o�א��v�Z"

    If Syuka_Update_Proc() Then
        Exit Function
    End If



    '�݌ɋ��z�ďW�v

Label3(0).Caption = "�݌ɋ��z�ďW�v"
    If Total_Update_Proc() Then
        Exit Function
    End If


On Error GoTo 0


                                    '�h�m�h�������t�o��
                                                        '2016.01.07 P_SYS.INI -- > PR00030.INI
    If WriteIni(App.EXEName, "START_DT" & Trim(GLB_SYUSHI_F), App.EXEName, Text1(ptxSTART_DT).Text) Then
'    If WriteIni(App.EXEName, "START_DT" & Trim(GLB_SYUSHI_F), "p_sys", Text1(ptxSTART_DT).Text) Then
        Beep
        MsgBox ("INI�t�@�C���̏������݂Ɏ��s���܂����B" & App.EXEName & "START_DT")
        Unload Me
    End If

    Text1(ptxLAST_START_DT).Text = Text1(ptxSTART_DT).Text


    START_DT = Text1(ptxSTART_DT).Text



                                                        '2016.01.07 P_SYS.INI -- > PR00030.INI
    If WriteIni(App.EXEName, "LAST_START_DT" & Trim(GLB_SYUSHI_F), App.EXEName, START_DT) Then
'    If WriteIni(App.EXEName, "LAST_START_DT" & Trim(GLB_SYUSHI_F), "p_sys", START_DT) Then
        Beep
        MsgBox ("INI�t�@�C���̏������݂Ɏ��s���܂����B" & App.EXEName & "LAST_START_DT")
        Unload Me
    End If
    LAST_START_DT = Text1(ptxSTART_DT).Text


    PR000301.MousePointer = vbDefault

    Update_Proc = False

    Exit Function



Error_Proc:

    If Err.Number = 13 Then
        MsgBox "Err.number= " & Err.Number
        Resume Next
    End If

End Function

Private Function RE_Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���ޒI�����ް��쐬
'----------------------------------------------------------------------------
Dim sts                     As Integer
Dim com                     As Integer

Dim Upd_Com                 As Integer

Dim i                       As Integer

Dim wk_Val                  As Long
Dim wKEIJYO_YM              As String * 6

Dim SKIP_FLG                As Boolean

Dim Check_Flg               As Boolean

Dim ZAIKO_KIN               As Long

Dim GK_ZEN_ZAIKO_KIN        As Long
Dim GK_NYUKO_KIN            As Long
Dim GK_SYUKO_KIN            As Long
Dim GK_ZAIKO_KIN            As Long

Dim Save_Jgyobu             As String * 1
Dim Save_Naigai             As String * 1
Dim Save_Hin_Gai            As String * 20
Dim Save_CODE               As String * 5
Dim Save_TANKA              As String * 11


Dim Sum_NYUKA_QTY           As Long

Dim wkZEN_ZAIKO             As Long             '2006.11.22

Dim ZAIKO_F                 As Boolean          '2007.04.26
Dim Save_G_Syushi           As String * 3       '2007.04.26


Dim SYUSHI_ON               As Boolean          '2007.11.13
Dim Fast_Flg                As Boolean          '2007.11.13


Dim wkZaiko_QTY             As Long             '2008.06.21
Dim wkNYUKO_QTY             As Long             '2008.06.21
Dim Syuko_Non_Flg           As Boolean          '2008.06.21


Dim Next_Jgyobu             As String           '2008.06.21
Dim Next_Naigai             As String           '2008.06.21
Dim Next_Hin_Gai            As String           '2008.06.21


    RE_Update_Proc = True
    PR000301.MousePointer = vbHourglass

    Check_Flg = False

'''GoTo L0             '2007.01.31
    com = BtOpGetFirst
    '�����ް��S���ر�
    Do

        DoEvents

        sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

        Select Case sts
            Case BtNoErr


            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "���ޒI�����ް�")
                Exit Function
        End Select
        Call UniCode_Conv(P_STOCK_REC.ZEN_ZAIKO_QTY, "00000000")    '�O���݌ɐ���


        Call UniCode_Conv(P_STOCK_REC.NYUKO_QTY, "00000000")        '���ɐ�
        Call UniCode_Conv(P_STOCK_REC.SYUKO_QTY, "00000000")        '�o�ɐ�
'        Call UniCode_Conv(P_STOCK_REC.MAEGARI_QTY, "00000000")      '�O��


        Call UniCode_Conv(P_STOCK_REC.ZEN_ZAIKO_KIN, "00000000")    '�O���݌ɐ���


        sts = BTRV(BtOpUpdate, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

        Select Case sts
            Case BtNoErr

            Case Else
                Call File_Error(sts, BtOpUpdate, "���ޒI�����ް�")
                Exit Function
        End Select


        com = BtOpGetNext

    Loop

    '�W�v�ް��S���폜
L0:
    com = BtOpGetFirst

    Do

        DoEvents

        sts = BTRV(com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

        Select Case sts
            Case BtNoErr


            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "���ޒI�����ް�")
                Exit Function
        End Select

        sts = BTRV(BtOpDelete, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

        Select Case sts
            Case BtNoErr

            Case Else
                Call File_Error(sts, BtOpDelete, "���ޒI�����ް�")
                Exit Function
        End Select


        com = BtOpGetNext

    Loop

''GoTo L1             '2007.01.31

    '-------------------------------------  �i�ڃ}�X�^���O���c�L�蕪���W�v
    If ZenZan_Update_Proc() Then
        Exit Function
    End If


    '-------------------------------------  ���ގ����蓖�����ɂ��W�v
    If SHIIRE_Update_Proc() Then
        Exit Function
    End If


    '-------------------------------------  �ďW�v�O�̒l�ɖ߂�
    com = BtOpGetFirst

    Do

        DoEvents

        sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

        Select Case sts
            Case BtNoErr


            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "���ޒI�����ް�")
                Exit Function
        End Select


        Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, StrConv(P_STOCK_REC.MOTO_ZAIKO_QTY, vbUnicode))



'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.12.13
        If RE_UPDATE_F = 1 Then
        
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_STOCK_REC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_STOCK_REC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))
        
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    If StrConv(ITEMREC.ZAIKO_CLR_F, vbUnicode) = "1" Then
                        Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, "00000000")
                    End If
                Case BtErrKeyNotFound
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                    Exit Function
        
            End Select
        End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.12.13



        sts = BTRV(BtOpUpdate, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

        Select Case sts
            Case BtNoErr

            Case Else
                Call File_Error(sts, BtOpUpdate, "���ޒI�����ް�")
                Exit Function
        End Select


''Call Log_Out("c:\yoshi.txt", StrConv(P_STOCK_REC.HIN_GAI, vbUnicode) & "=" & Format(CLng(StrConv(P_STOCK_REC.ZAIKO_QTY)), "#0"))
        com = BtOpGetNext

    Loop

'    GoTo L2

    '-------------------------------------  �����ȍ~�d���������݌ɂ������
''    wKEIJYO_YM = Left(Text1(ptxKEIJYO_YM).Text, 4) & Right(Text1(ptxKEIJYO_YM).Text, 2)
''
''    Call UniCode_Conv(K2_P_SHUKEIRE.KEIJYO_YM, wKEIJYO_YM)
''    Call UniCode_Conv(K2_P_SHUKEIRE.UKEIRE_DT, "zzzzzzzz")
''
''    com = BtOpGetGreater
''
''
''    Do
''
''        DoEvents
''
''        sts = BTRV(com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K2_P_SHUKEIRE, Len(K2_P_SHUKEIRE), 2)
''
''        Select Case sts
''            Case BtNoErr
''
''
''            Case BtErrEOF
''
''                Exit Do
''
''
''            Case Else
''                Call File_Error(sts, com, "���ގ���ް�")
''                Exit Function
''        End Select
''
''
''
''        '�����ް��ǂݍ���
''        Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode))
''        sts = BTRV(BtOpGetEqual, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
''        Skip_Flg = False
''        Select Case sts
''            Case BtNoErr
''                '�i�ڂ̍݌Ɍv���׸ނ��`�F�b�N
''                Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
''                Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
''                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
''
''                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
''                Skip_Flg = False
''                Select Case sts
''                    Case BtNoErr
''
''                        If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_OFF Then
''                            Skip_Flg = True       '�l�����Ȃ��̂Ž����
''                        End If
''
''
''
''
''                    Case BtErrKeyNotFound
''
''
''                        Skip_Flg = True       '�l�����Ȃ��̂Ž����
''
''
''                    Case Else
''                        Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
''                        Exit Function
''                End Select
''
''
''
''
''
''            Case BtErrKeyNotFound
''
''
''                Skip_Flg = True       '�����Ȃ��͒ʏ��ް��ł͂Ȃ�
''
''
''            Case Else
''                Call File_Error(sts, BtOpGetEqual, "���ޒ����ް�")
''                Exit Function
''        End Select
''
''
''        If Skip_Flg Then
''        Else
''
''            Call UniCode_Conv(K0_P_STOCK.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
''            Call UniCode_Conv(K0_P_STOCK.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
''            Call UniCode_Conv(K0_P_STOCK.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
''            Call UniCode_Conv(K0_P_STOCK.CODE, StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode))
''            Call UniCode_Conv(K0_P_STOCK.TANKA, StrConv(P_SHUKEIRE_REC.UKEIRE_TANKA, vbUnicode))
''            sts = BTRV(BtOpGetEqual, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
''
''            Select Case sts
''                Case BtNoErr
''
''                    upd_com = BtOpUpdate
''
''
''                Case BtErrKeyNotFound
''
''                    upd_com = BtOpInsert
''
''
''                Case Else
''
''                    Call File_Error(sts, BtOpGetEqual, "���ޒI�����ް�")
''                    Exit Function
''
''            End Select
''
''            If upd_com = BtOpUpdate Then
''
''
''                wk_VAL = CLng(CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) - _
''                CDbl(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)))
''
''                Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, Format(wk_VAL, "00000000"))
''
''
''                Do
''                    sts = BTRV(upd_com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
''
''                    Select Case sts
''                        Case BtNoErr
''                            Exit Do
''                        Case BtErrDuplicates
'''                            sts = BTRV(BtOpUpdate, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
'''                            If sts <> BtNoErr Then
'''                                Call File_Error(sts, BtOpUpdate, "���ޒI�����ް�")
'''                                Exit Function
'''                            End If
''                            Exit Do
''                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
''                            DoEvents
''                        Case Else
''
''                            Call File_Error(sts, upd_com, "���ޒI�����ް�")
''                            Exit Function
''                    End Select
''
''
''                Loop
''            End If
''        End If
''
''        com = BtOpGetNext
''
''    Loop


    '-------------------------------------  �O�ؕ��𓥂܂��ē����c�݌ɂ��ďW�v
''    com = BtOpGetFirst
''
''    Do
''
''        DoEvents
''
''        sts = BTRV(com, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
''
''        Select Case sts
''            Case BtNoErr
''
''
''            Case BtErrEOF
''
''                Exit Do
''
''
''            Case Else
''                Call File_Error(sts, com, "���ޑO���ް�")
''                Exit Function
''        End Select
''
''
''
''
''        If com = BtOpGetFirst Then
''            Save_Jgyobu = StrConv(P_NYUREC.JGYOBU, vbUnicode)
''            Save_Naigai = StrConv(P_NYUREC.NAIGAI, vbUnicode)
''            Save_Hin_Gai = StrConv(P_NYUREC.HIN_GAI, vbUnicode)
''            Save_CODE = StrConv(P_NYUREC.SHIIRE_CODE, vbUnicode)
''            Save_TANKA = StrConv(P_NYUREC.SHIIRE_TANKA, vbUnicode)
''            Sum_NYUKA_QTY = 0
''        End If
''
''        If Save_Jgyobu <> StrConv(P_NYUREC.JGYOBU, vbUnicode) Or _
''            Save_Naigai <> StrConv(P_NYUREC.NAIGAI, vbUnicode) Or _
''            Save_Hin_Gai <> StrConv(P_NYUREC.HIN_GAI, vbUnicode) Or _
''            Save_CODE <> StrConv(P_NYUREC.SHIIRE_CODE, vbUnicode) Or _
''            Save_TANKA <> StrConv(P_NYUREC.SHIIRE_TANKA, vbUnicode) Then
''
''            Call UniCode_Conv(K0_P_STOCK.JGYOBU, Save_Jgyobu)
''            Call UniCode_Conv(K0_P_STOCK.NAIGAI, Save_Naigai)
''            Call UniCode_Conv(K0_P_STOCK.HIN_GAI, Save_Hin_Gai)
''
''            If Trim(Save_CODE) = "" Then
''                Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_NYUREC.JGYOBU, vbUnicode))
''                Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_NYUREC.NAIGAI, vbUnicode))
''                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_NYUREC.HIN_GAI, vbUnicode))
''
''                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
''
''                Select Case sts
''                    Case BtNoErr
''
''
''                    Case BtErrKeyNotFound
''                        '���肦�Ȃ�
''                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).CODE, "")
''                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA, "")
''                    Case Else
''                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
''                        Exit Function
''                End Select
''
''
''                Call UniCode_Conv(K0_P_STOCK.CODE, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
''                Call UniCode_Conv(K0_P_STOCK.TANKA, StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))
''            Else
''                Call UniCode_Conv(K0_P_STOCK.CODE, Save_CODE)
''                Call UniCode_Conv(K0_P_STOCK.TANKA, Save_TANKA)
''            End If
''
''            sts = BTRV(BtOpGetEqual, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
''
''            Select Case sts
''                Case BtNoErr
''
''                    upd_com = BtOpUpdate
''
''
''                Case BtErrKeyNotFound
''
''                    upd_com = BtOpInsert
''
''
''                Case Else
''
''                    Call File_Error(sts, BtOpGetEqual, "���ޒI�����ް�")
''                    Exit Function
''
''            End Select
''
''
''
''            If upd_com = BtOpUpdate Then
''
''
''                If CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) < Sum_NYUKA_QTY Then
''                    Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, "00000000")
''                Else
''                    Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, Format(CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) - Sum_NYUKA_QTY, "00000000"))
''                End If
''
''                Call UniCode_Conv(P_STOCK_REC.MAEGARI_QTY, Format(Sum_NYUKA_QTY, "00000000"))
''
''                Do
''
''                    sts = BTRV(BtOpUpdate, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
''
''                    Select Case sts
''                        Case BtNoErr
''                            Exit Do
''
''                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
''                            DoEvents
''
''
''                        Case Else
''                            Call File_Error(sts, BtOpUpdate, "���ޒI���W�v�ް�")
''                            Exit Function
''                    End Select
''
''                Loop
''
''
''                Save_Jgyobu = StrConv(P_NYUREC.JGYOBU, vbUnicode)
''                Save_Naigai = StrConv(P_NYUREC.NAIGAI, vbUnicode)
''                Save_Hin_Gai = StrConv(P_NYUREC.HIN_GAI, vbUnicode)
''                Save_CODE = StrConv(P_NYUREC.SHIIRE_CODE, vbUnicode)
''                Save_TANKA = StrConv(P_NYUREC.SHIIRE_TANKA, vbUnicode)
''                Sum_NYUKA_QTY = 0
''
''
''            End If
''
''        End If
''
''        If IsNumeric(StrConv(P_NYUREC.SOUSAI_QTY, vbUnicode)) Then
''            Sum_NYUKA_QTY = Sum_NYUKA_QTY + CLng(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode)) - CLng(StrConv(P_NYUREC.SOUSAI_QTY, vbUnicode))
''        Else
''            Sum_NYUKA_QTY = Sum_NYUKA_QTY + CLng(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode))
''        End If
''        com = BtOpGetNext
''
''    Loop
''
''    If com <> BtOpGetFirst Then
''        Call UniCode_Conv(K0_P_STOCK.JGYOBU, Save_Jgyobu)
''        Call UniCode_Conv(K0_P_STOCK.NAIGAI, Save_Naigai)
''        Call UniCode_Conv(K0_P_STOCK.HIN_GAI, Save_Hin_Gai)
''
''        If Trim(Save_CODE) = "" Then
''            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_NYUREC.JGYOBU, vbUnicode))
''            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_NYUREC.NAIGAI, vbUnicode))
''            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_NYUREC.HIN_GAI, vbUnicode))
''
''            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
''
''            Select Case sts
''                Case BtNoErr
''
''
''                Case BtErrKeyNotFound
''                    '���肦�Ȃ�
''                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).CODE, "")
''                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA, "")
''                Case Else
''                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
''                    Exit Function
''            End Select
''
''
''            Call UniCode_Conv(K0_P_STOCK.CODE, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
''            Call UniCode_Conv(K0_P_STOCK.TANKA, StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))
''        Else
''            Call UniCode_Conv(K0_P_STOCK.CODE, Save_CODE)
''            Call UniCode_Conv(K0_P_STOCK.TANKA, Save_TANKA)
''        End If
''        sts = BTRV(BtOpGetEqual, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
''
''        Select Case sts
''            Case BtNoErr
''
''                upd_com = BtOpUpdate
''
''
''            Case BtErrKeyNotFound
''
''                upd_com = BtOpInsert
''
''
''            Case Else
''
''                Call File_Error(sts, BtOpGetEqual, "���ޒI�����ް�")
''                Exit Function
''
''        End Select
''
''
''
''        If upd_com = BtOpUpdate Then
''
''
''            If CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) < Sum_NYUKA_QTY Then
''                Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, "00000000")
''            Else
''                Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, Format(CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) - Sum_NYUKA_QTY, "00000000"))
''            End If
''
''
''            Call UniCode_Conv(P_STOCK_REC.MAEGARI_QTY, Format(Sum_NYUKA_QTY, "00000000"))
''
''
''            Do
''
''                sts = BTRV(BtOpUpdate, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
''
''                Select Case sts
''                    Case BtNoErr
''                        Exit Do
''
''                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
''                        DoEvents
''
''
''                    Case Else
''                        Call File_Error(sts, BtOpUpdate, "���ޒI���W�v�ް�")
''                        Exit Function
''                End Select
''
''            Loop
''        End If
''    End If


    '-------------------------------------  �ړ������I�������Ԓ��̕����W�v
L2:
    Call UniCode_Conv(K0_IDO.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_IDO.JITU_DT, Format(Text1(ptxS_INPUT_DT).Text, "YYYYMMDD"))
    Call UniCode_Conv(K0_IDO.JITU_TM, "")

    com = BtOpGetGreaterEqual

    Do

        DoEvents

        sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)

        Select Case sts
            Case BtNoErr

                If StrConv(IDOREC.JGYOBU, vbUnicode) <> SHIZAI Then
                    Exit Do
                End If

                If StrConv(IDOREC.JITU_DT, vbUnicode) > Format(Text1(ptxE_INPUT_DT).Text, "YYYYMMDD") Then
                    Exit Do
                End If
            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "�݌Ɉړ���")
                Exit Function
        End Select




        '�i�ڂ̍݌Ɍv���׸� & ���x���`�F�b�N        2007.11.13  ��
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(IDOREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(IDOREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(IDOREC.HIN_GAI, vbUnicode))

        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        SKIP_FLG = False
        Select Case sts
            Case BtNoErr

                If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_OFF Then
                    SKIP_FLG = True       '�l�����Ȃ��̂Ž����
'                    Call Log_Out(LOG_F, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
                End If


                SYUSHI_ON = False               '2007.11.13
                If GLB_SYUSHI_F = "" Then       '2007.11.13
                    SYUSHI_ON = True
                Else
                    SYUSHI_ON = False

                    For i = 0 To UBound(G_SYUSHI_TBL)

                        If Trim(StrConv(ITEMREC.G_SYUSHI, vbUnicode)) = G_SYUSHI_TBL(i) Then
                            SYUSHI_ON = True
                            Exit For
                        End If


                    Next i
                End If


            Case BtErrKeyNotFound


                SKIP_FLG = True       '�l�����Ȃ��̂Ž����


            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                Exit Function
        End Select
        '2007.11.13     ��




        For i = 0 To UBound(IN_YOIN)
            If IN_YOIN(i) = StrConv(IDOREC.RIRK_ID, vbUnicode) Then
                Exit For
            End If
        Next i

        If i <= UBound(IN_YOIN) And Not SKIP_FLG And SYUSHI_ON Then     '2007.11.13
            '�I����+
            Check_Flg = True
            SKIP_FLG = False



            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(IDOREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(IDOREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(IDOREC.HIN_GAI, vbUnicode))


            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

            Select Case sts
                Case BtNoErr
                    If Trim(StrConv(IDOREC.SHIIRE_CODE, vbUnicode)) = "" Then
                        Call UniCode_Conv(IDOREC.SHIIRE_CODE, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
                        Call UniCode_Conv(IDOREC.SHIIRE_TANKA, StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))
                    End If
                Case BtErrKeyNotFound

                    SKIP_FLG = True


                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                    Exit Function
            End Select

            If Not SKIP_FLG Then
                '2007.01.31 �O���c��GET
                Call UniCode_Conv(K0_P_STOCK.JGYOBU, StrConv(IDOREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_P_STOCK.NAIGAI, StrConv(IDOREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_P_STOCK.HIN_GAI, StrConv(IDOREC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(K0_P_STOCK.CODE, "")
                Call UniCode_Conv(K0_P_STOCK.TANKA, "")
                sts = BTRV(BtOpGetEqual, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

                Select Case sts
                    Case BtNoErr

                        If IsNumeric(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)) Then
                            wkZEN_ZAIKO = CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode))
                        Else
                            wkZEN_ZAIKO = 0
                        End If

                    Case BtErrKeyNotFound
                        wkZEN_ZAIKO = 0

                    Case Else

                        Call File_Error(sts, BtOpGetEqual, "���ޒI�����ް�")
                        Exit Function

                End Select
                '2007.01.31


                Call UniCode_Conv(K0_P_STOCK.JGYOBU, StrConv(IDOREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_P_STOCK.NAIGAI, StrConv(IDOREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_P_STOCK.HIN_GAI, StrConv(IDOREC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(K0_P_STOCK.CODE, StrConv(IDOREC.SHIIRE_CODE, vbUnicode))
                Call UniCode_Conv(K0_P_STOCK.TANKA, StrConv(IDOREC.SHIIRE_TANKA, vbUnicode))
                sts = BTRV(BtOpGetEqual, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

                Select Case sts
                    Case BtNoErr

                        Upd_Com = BtOpUpdate


                    Case BtErrKeyNotFound

                        Upd_Com = BtOpInsert


                    Case Else

                        Call File_Error(sts, BtOpGetEqual, "���ޒI�����ް�")
                        Exit Function

                End Select



                If Upd_Com = BtOpInsert Then
                    Call UniCode_Conv(P_STOCK_REC.JGYOBU, StrConv(IDOREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(P_STOCK_REC.NAIGAI, StrConv(IDOREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(P_STOCK_REC.HIN_GAI, StrConv(IDOREC.HIN_GAI, vbUnicode))
                    Call UniCode_Conv(P_STOCK_REC.CODE, StrConv(IDOREC.SHIIRE_CODE, vbUnicode))
                    Call UniCode_Conv(P_STOCK_REC.TANKA, StrConv(IDOREC.SHIIRE_TANKA, vbUnicode))




                    '2006.11.22
                    Call UniCode_Conv(P_STOCK_REC.INPUT_DATE, StrConv(IDOREC.NYUKA_DT, vbUnicode))


                    Call UniCode_Conv(P_STOCK_REC.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))
                    Call UniCode_Conv(P_STOCK_REC.ZEN_ZAIKO_QTY, "00000000")

                    Call UniCode_Conv(P_STOCK_REC.ZEN_ZAIKO_KIN, "00000000")



                    Call UniCode_Conv(P_STOCK_REC.NYUKO_QTY, "00000000")
                    Call UniCode_Conv(P_STOCK_REC.SYUKO_QTY, "00000000")
                    Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, "00000000")


                    Call UniCode_Conv(P_STOCK_REC.LAST_SYUKA_DT, StrConv(ITEMREC.LAST_SYU_DT, vbUnicode))
                    Call UniCode_Conv(P_STOCK_REC.LAST_SYUKA_QTY, StrConv(ITEMREC.G_LAST_SYUKA_QTY, vbUnicode))

                    Call UniCode_Conv(P_STOCK_REC.SYUKA_NON_F, "0")     '2008.06.21


                    Call UniCode_Conv(P_STOCK_REC.FILLER, "")



                End If

                '2006.11.22
                If StrConv(IDOREC.NYUKA_DT, vbUnicode) < StrConv(P_STOCK_REC.INPUT_DATE, vbUnicode) Then
                    Call UniCode_Conv(P_STOCK_REC.INPUT_DATE, StrConv(IDOREC.NYUKA_DT, vbUnicode))
                End If

                'Clng --> Val 2016.01.08
                wk_Val = Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) + _
                            Val(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + Val(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))

                If wk_Val < 0 Then
                    Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, Format(wk_Val, "0000000"))
                Else
                    Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, Format(wk_Val, "00000000"))
                End If


'2007.01.31                wk_VAL = CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)) + CLng(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)) - (CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)))
                'Clng --> Val 2016.01.08
                wk_Val = wkZEN_ZAIKO + CLng(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)) - (Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)))

                If wk_Val < 0 Then
                    Call UniCode_Conv(P_STOCK_REC.SYUKO_QTY, Format(wk_Val, "0000000"))
                Else
                    Call UniCode_Conv(P_STOCK_REC.SYUKO_QTY, Format(wk_Val, "00000000"))
                End If




                Do
                    sts = BTRV(Upd_Com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrDuplicates
'                            sts = BTRV(BtOpUpdate, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
'                            If sts <> BtNoErr Then
'                                Call File_Error(sts, BtOpUpdate, "���ޒI���ް�")
'                                Exit Function
'                            End If
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            DoEvents
                        Case Else

                            Call File_Error(sts, Upd_Com, "���ޒI���ް�")
                            Exit Function
                    End Select


                Loop


                Do
                    sts = BTRV(BtOpUpdate, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)

                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            DoEvents
                        Case Else

                            Call File_Error(sts, BtOpUpdate, "�݌Ɉړ���")
                            Exit Function
                    End Select


                Loop

            End If
        End If

        '�i�ڂ̍݌Ɍv���׸� & ���x���`�F�b�N        2007.11.13  ��
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(IDOREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(IDOREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(IDOREC.HIN_GAI, vbUnicode))

        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        SKIP_FLG = False
        Select Case sts
            Case BtNoErr

                If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_OFF Then
                    SKIP_FLG = True       '�l�����Ȃ��̂Ž����
'                    Call Log_Out(LOG_F, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
                End If


                SYUSHI_ON = False               '2007.11.13
                If GLB_SYUSHI_F = "" Then       '2007.11.13
                    SYUSHI_ON = True
                Else
                    SYUSHI_ON = False

                    For i = 0 To UBound(G_SYUSHI_TBL)

                        If Trim(StrConv(ITEMREC.G_SYUSHI, vbUnicode)) = G_SYUSHI_TBL(i) Then
                            SYUSHI_ON = True
                            Exit For
                        End If


                    Next i
                End If


            Case BtErrKeyNotFound


                SKIP_FLG = True       '�l�����Ȃ��̂Ž����


            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                Exit Function
        End Select




        For i = 0 To UBound(OUT_YOIN)
            If OUT_YOIN(i) = StrConv(IDOREC.RIRK_ID, vbUnicode) Then
                Exit For
            End If
        Next i



        If i <= UBound(OUT_YOIN) And Not SKIP_FLG And SYUSHI_ON Then        '2007.11.13
            '�I����-
            Check_Flg = True
            SKIP_FLG = False

            Call UniCode_Conv(ITEMREC.JGYOBU, StrConv(IDOREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(ITEMREC.NAIGAI, StrConv(IDOREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(IDOREC.HIN_GAI, vbUnicode))


            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

            Select Case sts
                Case BtNoErr
                    If Trim(StrConv(IDOREC.SHIIRE_CODE, vbUnicode)) = "" Then
                        Call UniCode_Conv(IDOREC.SHIIRE_CODE, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
                        Call UniCode_Conv(IDOREC.SHIIRE_TANKA, StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))
                    End If
                Case BtErrKeyNotFound

                    SKIP_FLG = True


                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                    Exit Function
            End Select

            If Not SKIP_FLG Then


                '2007.01.31 �O���c��GET
                Call UniCode_Conv(K0_P_STOCK.JGYOBU, StrConv(IDOREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_P_STOCK.NAIGAI, StrConv(IDOREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_P_STOCK.HIN_GAI, StrConv(IDOREC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(K0_P_STOCK.CODE, "")
                Call UniCode_Conv(K0_P_STOCK.TANKA, "")
                sts = BTRV(BtOpGetEqual, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

                Select Case sts
                    Case BtNoErr

                        If IsNumeric(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)) Then
                            wkZEN_ZAIKO = CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode))
                        Else
                            wkZEN_ZAIKO = 0
                        End If

                    Case BtErrKeyNotFound
                        wkZEN_ZAIKO = 0

                    Case Else

                        Call File_Error(sts, BtOpGetEqual, "���ޒI�����ް�")
                        Exit Function

                End Select
                '2007.01.31


                Call UniCode_Conv(K0_P_STOCK.JGYOBU, StrConv(IDOREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_P_STOCK.NAIGAI, StrConv(IDOREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_P_STOCK.HIN_GAI, StrConv(IDOREC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(K0_P_STOCK.CODE, StrConv(IDOREC.SHIIRE_CODE, vbUnicode))
                Call UniCode_Conv(K0_P_STOCK.TANKA, StrConv(IDOREC.SHIIRE_TANKA, vbUnicode))
                sts = BTRV(BtOpGetEqual, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

                Select Case sts
                    Case BtNoErr

                        Upd_Com = BtOpUpdate


                    Case BtErrKeyNotFound

                        Upd_Com = BtOpInsert


                    Case Else

                        Call File_Error(sts, BtOpGetEqual, "���ޒI�����ް�")
                        Exit Function

                End Select



                If Upd_Com = BtOpInsert Then
                    Call UniCode_Conv(P_STOCK_REC.JGYOBU, StrConv(IDOREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(P_STOCK_REC.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(P_STOCK_REC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                    Call UniCode_Conv(P_STOCK_REC.CODE, StrConv(IDOREC.SHIIRE_CODE, vbUnicode))
                    Call UniCode_Conv(P_STOCK_REC.TANKA, StrConv(IDOREC.SHIIRE_TANKA, vbUnicode))

                    '2006.11.22
                    Call UniCode_Conv(P_STOCK_REC.INPUT_DATE, StrConv(IDOREC.NYUKA_DT, vbUnicode))

                    Call UniCode_Conv(P_STOCK_REC.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))
                    Call UniCode_Conv(P_STOCK_REC.ZEN_ZAIKO_QTY, "00000000")
                    Call UniCode_Conv(P_STOCK_REC.ZEN_ZAIKO_KIN, "00000000")



                    Call UniCode_Conv(P_STOCK_REC.NYUKO_QTY, "00000000")
                    Call UniCode_Conv(P_STOCK_REC.SYUKO_QTY, "00000000")
                    Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, "00000000")


                    Call UniCode_Conv(P_STOCK_REC.LAST_SYUKA_DT, StrConv(ITEMREC.LAST_SYU_DT, vbUnicode))
                    Call UniCode_Conv(P_STOCK_REC.LAST_SYUKA_QTY, StrConv(ITEMREC.G_LAST_SYUKA_QTY, vbUnicode))


                    Call UniCode_Conv(P_STOCK_REC.MAEGARI_QTY, "00000000")

                    Call UniCode_Conv(P_STOCK_REC.SYUKA_NON_F, "0")     '2008.06.21

                    Call UniCode_Conv(P_STOCK_REC.FILLER, "")



                End If

                '2006.11.22
                If StrConv(IDOREC.NYUKA_DT, vbUnicode) < StrConv(P_STOCK_REC.INPUT_DATE, vbUnicode) Then
                    Call UniCode_Conv(P_STOCK_REC.INPUT_DATE, StrConv(IDOREC.NYUKA_DT, vbUnicode))
                End If
                'Clng --> Val 2016.01.08
                wk_Val = Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) - _
                           (Val(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + Val(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)))

                If wk_Val < 0 Then
                    Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, Format(wk_Val, "0000000"))
                Else
                    Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, Format(wk_Val, "00000000"))
                End If

'2007.01.31                wk_VAL = CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)) + CLng(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)) - (CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)))
                'Clng --> Val 2016.01.08
                wk_Val = wkZEN_ZAIKO + Val(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)) - (Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)))

                If wk_Val < 0 Then
                    Call UniCode_Conv(P_STOCK_REC.SYUKO_QTY, Format(wk_Val, "0000000"))
                Else
                    Call UniCode_Conv(P_STOCK_REC.SYUKO_QTY, Format(wk_Val, "00000000"))
                End If


                Do
                    sts = BTRV(Upd_Com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrDuplicates
'                            sts = BTRV(BtOpUpdate, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
'                            If sts <> BtNoErr Then
'                                Call File_Error(sts, BtOpUpdate, "���ޒI�����W�v�ް�")
'                                Exit Function
'                            End If
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            DoEvents
                        Case Else

                            Call File_Error(sts, Upd_Com, "���ޒI�����W�v�ް�")
                            Exit Function
                    End Select


                Loop

''2006.11.22                sts = BTRV(BtOpGetEqual, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
''2006.11.22
''2006.11.22                Select Case sts
''2006.11.22                    Case BtNoErr
''2006.11.22
''2006.11.22                    Case BtErrKeyNotFound
''2006.11.22
''2006.11.22                        Call UniCode_Conv(P_STOCK_REC.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
''2006.11.22                        Call UniCode_Conv(P_STOCK_REC.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
''2006.11.22                        Call UniCode_Conv(P_STOCK_REC.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
''2006.11.22                        Call UniCode_Conv(P_STOCK_REC.CODE, "")
''2006.11.22                        Call UniCode_Conv(P_STOCK_REC.TANKA, "")
''2006.11.22
''2006.11.22                        Call UniCode_Conv(P_STOCK_REC.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))
''2006.11.22                        Call UniCode_Conv(P_STOCK_REC.ZEN_ZAIKO_QTY, "00000000")
''2006.11.22
''2006.11.22
''2006.11.22
''2006.11.22                        Call UniCode_Conv(P_STOCK_REC.NYUKO_QTY, "00000000")
''2006.11.22                        Call UniCode_Conv(P_STOCK_REC.SYUKO_QTY, "00000000")
''2006.11.22                        Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, "00000000")
''2006.11.22
''2006.11.22
''2006.11.22                        Call UniCode_Conv(P_STOCK_REC.LAST_SYUKA_DT, StrConv(ITEMREC.LAST_SYU_DT, vbUnicode))
''2006.11.22                        Call UniCode_Conv(P_STOCK_REC.LAST_SYUKA_QTY, StrConv(ITEMREC.G_LAST_SYUKA_QTY, vbUnicode))
''2006.11.22
''2006.11.22                        Call UniCode_Conv(P_STOCK_REC.MAEGARI_QTY, "00000000")
''2006.11.22
''2006.11.22
''2006.11.22                        Call UniCode_Conv(P_STOCK_REC.FILLER, "")
''2006.11.22
''2006.11.22
''2006.11.22                        Do
''2006.11.22                            sts = BTRV(BtOpInsert, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
''2006.11.22
''2006.11.22                            Select Case sts
''2006.11.22                                Case BtNoErr, BtErrDuplicates
''2006.11.22                                    Exit Do
''2006.11.22                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
''2006.11.22                                    DoEvents
''2006.11.22                                Case Else
''2006.11.22
''2006.11.22                                    Call File_Error(sts, BtOpGetEqual, "���ޒI���ް�")
''2006.11.22                                    Exit Function
''2006.11.22                            End Select
''2006.11.22
''2006.11.22
''2006.11.22                        Loop
''2006.11.22
''2006.11.22
''2006.11.22
''2006.11.22                    Case Else
''2006.11.22
''2006.11.22                        Call File_Error(sts, BtOpGetEqual, "���ޒI�����ް�")
''2006.11.22                        Exit Function
''2006.11.22
''2006.11.22                End Select
''2006.11.22
''2006.11.22
''2006.11.22
                Do
                    sts = BTRV(BtOpUpdate, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)

                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            DoEvents
                        Case Else

                            Call File_Error(sts, BtOpUpdate, "�݌Ɉړ���")
                            Exit Function
                    End Select


                Loop
            End If
        End If

        com = BtOpGetNext

    Loop

    '-------------------------------------  �o�ɐ��̌v�Z


'------------------------------------------------------�@�o�א��e�Z�b�g 2008.06.21
    If Syuka_F_Update_Proc() Then
        Exit Function
    End If
'------------------------------------------------------�@�o�א��e�Z�b�g 2008.06.21
 Label3(0).Caption = "�i�ڏW�v����"

    If Hin_Sum_Update_Proc() Then
        Exit Function
    End If


 Label3(0).Caption = "�o�א��v�Z"

    If Syuka_Update_Proc() Then
        Exit Function
    End If
    
    





    '�W�v�ް��݌ɋ��z�ر�

    com = BtOpGetFirst

    Do

        DoEvents

        sts = BTRV(com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

        Select Case sts
            Case BtNoErr


            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "���ޒI�����ް�")
                Exit Function
        End Select

        Call UniCode_Conv(P_STOCKSUM_REC.ZAIKO_KIN, "00000000000")

        sts = BTRV(BtOpUpdate, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

        Do
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    DoEvents
                Case Else
                    Call File_Error(sts, BtOpDelete, "���ޒI�����ް�")
                    Exit Function
            End Select
        Loop

        com = BtOpGetNext

    Loop

    '�݌ɋ��z�ďW�v
    If Total_Update_Proc() Then
        Exit Function
    End If

    PR000301.MousePointer = vbDefault

    RE_Update_Proc = False

End Function

Public Function wP_STOCK_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���ޒI�����ް�  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String


Dim ret         As Long


    wP_STOCK_Open = True
                                            '���ޒI���f�[�^�t���p�X�捞��
    sts = GetIni("FILE", tmpP_STOCK_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_STOCK]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = Trim(c)



    '2007.11.13
'    FullPath = Trim(c)
    ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), ret) & GLB_SYUSHI_F & Right(Trim(c), Len(Trim(c)) - ret)
    '2007.11.13





    Do
        sts = BTRV(BtOpOpen, wP_STOCK_POS, wP_STOCK_REC, Len(wP_STOCK_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "w���ޒI�����ް�")
                Exit Function
        End Select
    Loop

    wP_STOCK_Open = False

End Function
Public Function wkP_STOCK_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���ޒI�����ް�  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*      2010.01.14
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String


Dim ret         As Long


    wkP_STOCK_Open = True
                                            '���ޒI���f�[�^�t���p�X�捞��
    sts = GetIni("FILE", P_STOCK_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_STOCK]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = Trim(c)

    ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), ret) & GLB_SYUSHI_F & Right(Trim(c), Len(Trim(c)) - ret)


    Do
        sts = BTRV(BtOpOpen, wkP_STOCK_POS, wkP_STOCK_REC, Len(wkP_STOCK_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "wk���ޒI�����ް�")
                Exit Function
        End Select
    Loop

    wkP_STOCK_Open = False

End Function


Private Function Next_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���ތJ�z����
'----------------------------------------------------------------------------
Dim sts                     As Integer
Dim com                     As Integer


Dim i                       As Integer

Dim wk_Val                  As Long

Dim SKIP_FLG                As Boolean

Dim SYUSHI_ON               As Boolean          '2007.11.13

Dim Sum_Zen_Zaiko           As Long

Dim Sum_Zen_Zaiko_KIN       As Long             '2011.10.18


Dim Sum_Zaiko               As Long



Dim Sum_Nyuko               As Long
Dim Sum_Syuko               As Long

Dim Sum_Zaiko_KIN           As Long




Dim svJGYOBU                As String * 1
Dim svNAIGAI                As String * 1
Dim svHIN_GAI               As String * 20




    Next_Proc = True
    PR000301.MousePointer = vbHourglass


    '-------------------------------------  �i�ڃ}�X�^�̓��e��ر�����


    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_ITEM.NAIGAI, "")
    Call UniCode_Conv(K0_ITEM.HIN_GAI, "")

    com = BtOpGetGreater


    Do

        DoEvents

        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

        Select Case sts
            Case BtNoErr

                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> SHIZAI Then
                    Exit Do
                End If

            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "�i�ڃ}�X�^")
                Exit Function
        End Select

        SYUSHI_ON = False               '2007.11.13
        If GLB_SYUSHI_F = "" Then       '2007.11.13
            SYUSHI_ON = True
        Else
            SYUSHI_ON = False

            For i = 0 To UBound(G_SYUSHI_TBL)

                If Trim(StrConv(ITEMREC.G_SYUSHI, vbUnicode)) = G_SYUSHI_TBL(i) Then
                    SYUSHI_ON = True
                    Exit For
                End If


            Next i
        End If




        If SYUSHI_ON Then

            '2009.10.30
            Call UniCode_Conv(K0_P_STOCK.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_P_STOCK.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_P_STOCK.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))

            Call UniCode_Conv(K0_P_STOCK.CODE, "")
            Call UniCode_Conv(K0_P_STOCK.TANKA, "")



            Sum_Zen_Zaiko = 0
            
            
            Sum_Zen_Zaiko_KIN = 0           '2011.10.18
            
            
            Sum_Zaiko = 0
            Sum_Nyuko = 0
            Sum_Syuko = 0


            Sum_Zaiko_KIN = 0           '2018.01.24

            com = BtOpGetGreaterEqual



            Do
                DoEvents


                sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

                Select Case sts
                    Case BtNoErr

                        If StrConv(ITEMREC.JGYOBU, vbUnicode) <> StrConv(P_STOCK_REC.JGYOBU, vbUnicode) Or _
                            StrConv(ITEMREC.NAIGAI, vbUnicode) <> StrConv(P_STOCK_REC.NAIGAI, vbUnicode) Or _
                            StrConv(ITEMREC.HIN_GAI, vbUnicode) <> StrConv(P_STOCK_REC.HIN_GAI, vbUnicode) Then
                            Exit Do
                        End If

                    Case BtErrEOF

                        Exit Do


                    Case Else
                        Call File_Error(sts, com, "�i�ڃ}�X�^")
                        Exit Function
                End Select


                If IsNumeric(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)) Then
                    Sum_Zen_Zaiko = Sum_Zen_Zaiko + CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode))
                End If
                
                
''''''''''''''''''''2011.10.18
                If IsNumeric(StrConv(P_STOCK_REC.ZEN_ZAIKO_KIN, vbUnicode)) Then
                    Sum_Zen_Zaiko_KIN = Sum_Zen_Zaiko_KIN + CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_KIN, vbUnicode))
                End If
''''''''''''''''''''2011.10.18
                
                
                
                
                If Trim(StrConv(P_STOCK_REC.TANKA, vbUnicode)) <> "" Or Trim(StrConv(P_STOCK_REC.CODE, vbUnicode)) <> "" Then
                
                    If IsNumeric(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) Then
                        'Clng --> Val 2016.01.08
                        Sum_Zaiko = Sum_Zaiko + Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))
                    End If
                    If IsNumeric(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) Then
                        'Clng --> Val 2016.01.08
                        Sum_Nyuko = Sum_Nyuko + Val(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode))
                    End If
                    If IsNumeric(StrConv(P_STOCK_REC.SYUKO_QTY, vbUnicode)) Then
                        Sum_Syuko = Sum_Syuko + CLng(StrConv(P_STOCK_REC.SYUKO_QTY, vbUnicode))
                    End If
                
                
                    
''''''''''''''''''''2011.10.18
                    If IsNumeric(StrConv(P_STOCK_REC.TANKA, vbUnicode)) Then
                        Sum_Zaiko_KIN = Sum_Zaiko_KIN + ToRoundUp(CCur(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) * _
                                                    CCur(StrConv(P_STOCK_REC.TANKA, vbUnicode)), 0)
                    End If
''''''''''''''''''''2011.10.18
                
                
                
                End If


                com = BtOpGetNext



            Loop


''''''''''''''''''''2011.10.18
            '''If Sum_Zen_Zaiko <> Sum_Zaiko Then
            If Sum_Zen_Zaiko <> Sum_Zaiko Or Sum_Zen_Zaiko_KIN <> Sum_Zaiko_KIN Then
''''''''''''''''''''2011.10.18



                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "00000000000")
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "00000000")

                Do
                    sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            DoEvents
                        Case Else

                            Call File_Error(sts, BtOpUpdate, "���ޒI���ް�")
                            Exit Function
                    End Select


                Loop

            Else
                If Sum_Nyuko = 0 And Sum_Syuko = 0 Then

'Call Log_Out(LOG_F, "CLR=" & StrConv(ITEMREC.HIN_GAI, vbUnicode))
                Else






                    Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "00000000000")
                    Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "00000000")

                    Do
                        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                DoEvents
                            Case Else

                                Call File_Error(sts, BtOpUpdate, "���ޒI���ް�")
                                Exit Function
                        End Select


                    Loop



                End If


            End If
        End If
        com = BtOpGetNext

    Loop

    
    
    
    
    '-------------------------------------  �������c��Ă���

    svJGYOBU = ""
    svNAIGAI = ""
    svHIN_GAI = ""


    com = BtOpGetFirst

    Do

        DoEvents

        sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

        Select Case sts
            Case BtNoErr

            Case BtErrEOF
                Exit Do

            Case Else
                Call File_Error(sts, com, "���ޒI���ް�")
                Exit Function

        End Select


        If Trim(svJGYOBU) = "" Then
            svJGYOBU = StrConv(P_STOCK_REC.JGYOBU, vbUnicode)
            svNAIGAI = StrConv(P_STOCK_REC.NAIGAI, vbUnicode)
            svHIN_GAI = StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)

            Sum_Zen_Zaiko = 0
            
            
            Sum_Zen_Zaiko_KIN = 0       '2011.10.18
            
            Sum_Zaiko = 0
            Sum_Nyuko = 0
            Sum_Syuko = 0

            Sum_Zaiko_KIN = 0
        End If




        If svJGYOBU <> StrConv(P_STOCK_REC.JGYOBU, vbUnicode) Or _
            svNAIGAI <> StrConv(P_STOCK_REC.NAIGAI, vbUnicode) Or _
            svHIN_GAI <> StrConv(P_STOCK_REC.HIN_GAI, vbUnicode) Then

''''''''''''''''''''2011.10.18
            '''If Sum_Zen_Zaiko = Sum_Zaiko And Sum_Nyuko = 0 And Sum_Syuko = 0 Then
          
          
          
          
          
          
            
            If Sum_Zen_Zaiko = Sum_Zaiko And Sum_Zen_Zaiko_KIN = Sum_Zaiko_KIN And Sum_Nyuko = 0 And Sum_Syuko = 0 Then
''''''''''''''''''''2011.10.18


            Else
                Call UniCode_Conv(K0_ITEM.JGYOBU, svJGYOBU)
                Call UniCode_Conv(K0_ITEM.NAIGAI, svNAIGAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, svHIN_GAI)

                SKIP_FLG = False
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

                Select Case sts
                    Case BtNoErr

                    Case BtErrKeyNotFound
                        SKIP_FLG = True

                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                        Exit Function

                End Select


                If Not SKIP_FLG Then

                    If Not IsNumeric(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode)) Then
                        Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "00000000000")
                    End If

                    wk_Val = CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode))

                    wk_Val = wk_Val + Sum_Zaiko_KIN

                    If wk_Val < 0 Then
                        wk_Val = 0
                    End If
                    Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, Format(wk_Val, "00000000000"))

                    If Not IsNumeric(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode)) Then
                        Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "000000000")
                    End If

                    wk_Val = CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode))
                    wk_Val = wk_Val + Sum_Zaiko

                    If wk_Val < 0 Then
                        wk_Val = 0
                    End If
                    Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, Format(wk_Val, "00000000"))

                    Do
                        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

                        Select Case sts
                            Case BtNoErr
                                Exit Do

                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                DoEvents

                            Case Else
                                Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^")
                                Exit Function

                        End Select
                    Loop

                End If

            End If

            svJGYOBU = StrConv(P_STOCK_REC.JGYOBU, vbUnicode)
            svNAIGAI = StrConv(P_STOCK_REC.NAIGAI, vbUnicode)
            svHIN_GAI = StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)


            Sum_Zen_Zaiko = 0
            Sum_Zen_Zaiko_KIN = 0   '2018.01.24
            Sum_Zaiko = 0
            Sum_Nyuko = 0
            Sum_Syuko = 0


            Sum_Zaiko_KIN = 0

        End If


        If Trim(StrConv(P_STOCK_REC.TANKA, vbUnicode)) = "" And Trim(StrConv(P_STOCK_REC.CODE, vbUnicode)) = "" Then

            If IsNumeric(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)) Then
                Sum_Zen_Zaiko = Sum_Zen_Zaiko + CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode))

            End If

''''''''''''''''''''2011.10.18
            If IsNumeric(StrConv(P_STOCK_REC.ZEN_ZAIKO_KIN, vbUnicode)) Then
                Sum_Zen_Zaiko_KIN = Sum_Zen_Zaiko_KIN + CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_KIN, vbUnicode))

            End If
''''''''''''''''''''2011.10.18


        Else
            If IsNumeric(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)) Then
                Sum_Nyuko = Sum_Nyuko + CLng(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode))
            End If

            If IsNumeric(StrConv(P_STOCK_REC.SYUKO_QTY, vbUnicode)) Then
                Sum_Syuko = Sum_Syuko + CLng(StrConv(P_STOCK_REC.SYUKO_QTY, vbUnicode))
            End If


            If IsNumeric(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) Then
                Sum_Zaiko = Sum_Zaiko + CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))
            End If



            If IsNumeric(StrConv(P_STOCK_REC.TANKA, vbUnicode)) And IsNumeric(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) Then

                Sum_Zaiko_KIN = Sum_Zaiko_KIN + ToRoundUp(CCur(StrConv(P_STOCK_REC.TANKA, vbUnicode)) * _
                                CCur(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)), 0)
            End If

        End If



'        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_STOCK_REC.JGYOBU, vbUnicode))
'        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_STOCK_REC.NAIGAI, vbUnicode))
'        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))
'
'
'        Skip_Flg = False
'        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'
'
'        Select Case sts
'            Case BtNoErr
'            Case BtErrKeyNotFound
'
'
'                Skip_Flg = True
'            Case Else
'
'                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
'                Exit Function
'        End Select
'
'
'        If Not Skip_Flg Then
'
'If Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) = "D900" Then
'    Debug.Print
'End If
'            wk_VAL = CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode))
'
'            If IsNumeric(StrConv(P_STOCK_REC.TANKA, vbUnicode)) Then
'            '2009.09.30 0.5->0.9
''                wk_VAL = wk_VAL + Int(CDbl(CDbl(StrConv(P_STOCK_REC.TANKA, vbUnicode)) * CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) + 0.5))
'                wk_VAL = wk_VAL + Format(ToRoundUp(CCur(StrConv(P_STOCK_REC.TANKA, vbUnicode)) * _
'                                    CCur(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)), 0), "#,##0")
'
'
'
'
'
'                If wk_VAL < 0 Then
'                    wk_VAL = 0
'                End If
'                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, Format(wk_VAL, "00000000000"))
'            End If
'
'            wk_VAL = CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode))
'            wk_VAL = wk_VAL + CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))
'
'            If wk_VAL < 0 Then
'                wk_VAL = 0
'            End If
'            Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, Format(wk_VAL, "00000000"))
'
'
'            Do
'                sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'
'                Select Case sts
'                    Case BtNoErr
'                        Exit Do
'                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                        DoEvents
'                    Case Else
'
'                        Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^")
'                        Exit Function
'                End Select
'
'
'            Loop
'
'        End If



        com = BtOpGetNext

    Loop



    If Trim(svJGYOBU) <> "" Then

        
''''''''''''''''''''2011.10.18
'''        If Sum_Zen_Zaiko = Sum_Zaiko And Sum_Nyuko = 0 And Sum_Syuko = 0 Then
        If Sum_Zen_Zaiko = Sum_Zaiko And Sum_Zen_Zaiko_KIN = Sum_Zaiko_KIN And Sum_Nyuko = 0 And Sum_Syuko = 0 Then
''''''''''''''''''''2011.10.18

        Else
            Call UniCode_Conv(K0_ITEM.JGYOBU, svJGYOBU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, svNAIGAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, svHIN_GAI)


            SKIP_FLG = False
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)


            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound


                    SKIP_FLG = True
                Case Else

                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Function
            End Select


            If Not SKIP_FLG Then


                If Not IsNumeric(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "00000000000")
                End If

                wk_Val = CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode))

                wk_Val = wk_Val + Sum_Zaiko_KIN

                If wk_Val < 0 Then
                    wk_Val = 0
                End If
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, Format(wk_Val, "00000000000"))


                If Not IsNumeric(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "000000000")
                End If


                wk_Val = CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode))
                wk_Val = wk_Val + Sum_Zaiko

                If wk_Val < 0 Then
                    wk_Val = 0
                End If
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, Format(wk_Val, "00000000"))




                Do
                    sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            DoEvents
                        Case Else

                            Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^")
                            Exit Function
                    End Select


                Loop

            End If

        End If
    End If



                                    '�h�m�h�������t�o��
                                                                        '2016.01.07 P_SYS.INI --> PR00030.INI
    If WriteIni(App.EXEName, "LAST_SHIME_DT" & Trim(GLB_SYUSHI_F), App.EXEName, Format(Now, "YYYY/MM/DD")) Then
'    If WriteIni(App.EXEName, "LAST_SHIME_DT" & Trim(GLB_SYUSHI_F), "p_sys", Format(Now, "YYYY/MM/DD")) Then
        Beep
        MsgBox ("INI�t�@�C���̏������݂Ɏ��s���܂����B" & App.EXEName & "LAST_SHIME_DT")
        Unload Me
    End If

    LAST_SHIME_DT = Format(Now, "YYYY/MM/DD")


    PR000301.MousePointer = vbDefault

    MsgBox "�J�z�������I�����܂����B"

    Next_Proc = False

End Function

Private Function Data_Out_Proc() As Integer
'----------------------------------------------------------------------------
'           EXCEL�ɏo�͂���
'----------------------------------------------------------------------------

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>2011.10.18
'Dim excelApplication    As Excel.Application
'Dim excelApplication    As Excel.Workbook
'Dim excelApplication    As Excel.Worksheet

Dim excelApplication    As Object
Dim excelWorkBook       As Object
Dim excelSheet          As Object
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>2011.10.18



Dim com                 As Integer
Dim sts                 As Integer

Dim FSW                 As Boolean
Dim Lcnt                As Integer
Dim ZAIKO_F             As Boolean

Dim c                   As String * 128
Dim FileName            As String

Dim yn                  As Integer




    Data_Out_Proc = True


    Call Input_Lock


    If GetIni("FILE", tmpP_STOCK_ID, "SYS", c) Then


        Call Input_UnLock

        Beep
        MsgBox "tmp���ނ̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        Exit Function
    End If

    FileName = Trim(c)
''    On Error Resume Next
''    Kill (fileName)
''    On Error GoTo 0


    On Error GoTo Data_Out_Proc_Error
    Kill (FileName)



    If tmpP_STOCK_Open(0) Then
        Call Input_UnLock
        Exit Function
    End If




'2011.02.14 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    com = BtOpGetFirst
    
    
    Do
        DoEvents
    
        sts = BTRV(com, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), K0_tmpP_STOCK, Len(K0_tmpP_STOCK), 0)
        
        Select Case sts
            Case BtNoErr
            
                sts = BTRV(BtOpDelete, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), K0_tmpP_STOCK, Len(K0_tmpP_STOCK), 0)
                If sts Then
                    Call File_Error(sts, BtOpDelete, "tmp���ޒI���W�v�ް�")
                    Exit Function
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "tmp���ޒI���W�v�ް�")
                Exit Function
        End Select
    
    Loop
    




'2011.02.14 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<





    If tmpP_STOCK_MAKE_Proc() Then
        Call Input_UnLock
        Exit Function
    End If




    'Excel���ع���ݵ�޼ު�Ď擾
    Set excelApplication = CreateObject("Excel.Application")
'    excelApplication.Visible = True                2016.04.27

    Set excelWorkBook = excelApplication.Workbooks.Open(exSheet)    '����Ώ��ޯ����J��
    Set excelSheet = excelWorkBook.Worksheets(1)                    '�P��Ėڂ�I��




    Lcnt = LStart


    FSW = True


    '�����N����
    excelSheet.Application.Cells(1, 1).Value = StrConv(Left(PR000301.Text1(1).Text, 4), vbWide) & "�N" & _
                                        StrConv(Right(PR000301.Text1(1).Text, 2), vbWide) & "��"
    '���s����
''    excelSheet.Application.Cells(1, 12).Value = Format(Now, "YYYY/MM/SS HH:MM:SS")
    excelSheet.Application.Cells(1, 12).Value = Format(Now, "YYYY/MM/DD HH:NN:SS")      '2010/10/21 upd





    com = BtOpGetFirst

    Do

        DoEvents

        sts = BTRV(com, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), K1_tmpP_STOCK, Len(K1_tmpP_STOCK), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "���ޒI���ް�")
                Exit Function
        End Select


        If FSW Then
            FSW = False
        Else
            excelSheet.Application.Range(Lcnt - 1 & ":" & Lcnt - 1).Copy
            excelSheet.Application.Range(Lcnt & ":" & Lcnt).Insert
        End If


        '�i��
        excelSheet.Application.Cells(Lcnt, exHin_Gai).Value = Trim(StrConv(tmpP_STOCK_REC.HIN_GAI, vbUnicode))
        '�i��
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(tmpP_STOCK_REC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(tmpP_STOCK_REC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(tmpP_STOCK_REC.HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                excelSheet.Application.Cells(Lcnt, exHin_Name).Value = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
            Case BtErrKeyNotFound
                excelSheet.Application.Cells(Lcnt, exHin_Name).Value = ""
            Case Else
                Call Input_UnLock
                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                Exit Function
        End Select

        '�݌Ɍ�
        excelSheet.Application.Cells(Lcnt, exG_SYUSHI).Value = StrConv(tmpP_STOCK_REC.G_SYUSHI, vbUnicode)
        '�O���݌ɐ�
        If IsNumeric(StrConv(tmpP_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)) Then
            excelSheet.Application.Cells(Lcnt, exZEN_ZAIKO_QTY).Value = CLng(StrConv(tmpP_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode))
        Else
            excelSheet.Application.Cells(Lcnt, exZEN_ZAIKO_QTY).Value = ""
        End If
        '���ɐ�
        If IsNumeric(StrConv(tmpP_STOCK_REC.NYUKO_QTY, vbUnicode)) Then
            excelSheet.Application.Cells(Lcnt, exNYUKO_QTY).Value = CLng(StrConv(tmpP_STOCK_REC.NYUKO_QTY, vbUnicode))
        Else
            excelSheet.Application.Cells(Lcnt, exNYUKO_QTY).Value = ""
        End If

        '�o�ɐ�
        If IsNumeric(StrConv(tmpP_STOCK_REC.SYUKO_QTY, vbUnicode)) Then
            excelSheet.Application.Cells(Lcnt, exSYUKO_QTY).Value = CLng(StrConv(tmpP_STOCK_REC.SYUKO_QTY, vbUnicode))
        Else
            excelSheet.Application.Cells(Lcnt, exSYUKO_QTY).Value = ""
        End If
        '�����݌�
        If IsNumeric(StrConv(tmpP_STOCK_REC.ZAIKO_QTY, vbUnicode)) Then
            excelSheet.Application.Cells(Lcnt, exZAIKO_QTY).Value = CLng(StrConv(tmpP_STOCK_REC.ZAIKO_QTY, vbUnicode))
        Else
            excelSheet.Application.Cells(Lcnt, exZAIKO_QTY).Value = ""
        End If
        '�d���P��
        If IsNumeric(StrConv(tmpP_STOCK_REC.TANKA, vbUnicode)) Then
            excelSheet.Application.Cells(Lcnt, exSHI_TANKA).Value = CDbl(StrConv(tmpP_STOCK_REC.TANKA, vbUnicode))
        Else
            excelSheet.Application.Cells(Lcnt, exSHI_TANKA).Value = ""
        End If
        
'>>>>>>>>   2018.07.25
        '�����݌ɋ��z
'        If IsNumeric(StrConv(tmpP_STOCK_REC.TANKA, vbUnicode)) And IsNumeric(StrConv(tmpP_STOCK_REC.ZAIKO_QTY, vbUnicode)) Then
'            excelSheet.Application.Cells(Lcnt, exZAIKO_KIN).Value = CLng(CDbl(StrConv(tmpP_STOCK_REC.TANKA, vbUnicode)) * CLng(StrConv(tmpP_STOCK_REC.ZAIKO_QTY, vbUnicode)))
'        Else
'            excelSheet.Application.Cells(Lcnt, exZAIKO_KIN).Value = ""
'        End If

'>>>>>>>>>>>>
    If IsNumeric(StrConv(tmpP_STOCK_REC.TANKA, vbUnicode)) Then
'        STOCK(ROW, colZAIKO_KIN) = Format(CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) * _
'                                    CDbl(StrConv(P_STOCK_REC.TANKA, vbUnicode)), "#,##0")

        If Not IsNumeric(StrConv(tmpP_STOCK_REC.ZAIKO_QTY, vbUnicode)) Then
            Call UniCode_Conv(tmpP_STOCK_REC.ZAIKO_QTY, "00000000")
        End If
       If Not IsNumeric(StrConv(tmpP_STOCK_REC.TANKA, vbUnicode)) Then
            Call UniCode_Conv(P_STOCK_REC.TANKA, "00000000")
       End If

        excelSheet.Application.Cells(Lcnt, exZAIKO_KIN).Value = Format(ToRoundUp(CCur(StrConv(tmpP_STOCK_REC.ZAIKO_QTY, vbUnicode)) * _
                                    CCur(StrConv(tmpP_STOCK_REC.TANKA, vbUnicode)), 0), "#,##0")


    Else
        excelSheet.Application.Cells(Lcnt, exZAIKO_KIN).Value = ""
    End If




'>>>>>>>>   2018.07.25

        '�d����
        excelSheet.Application.Cells(Lcnt, exSHI_CODE).Value = StrConv(tmpP_STOCK_REC.CODE, vbUnicode)
        '�ŏI�o�ד�
        If Trim(StrConv(tmpP_STOCK_REC.LAST_SYUKA_DT, vbUnicode)) = "" Then
            excelSheet.Application.Cells(Lcnt, exLAST_SYUKA_DT).Value = ""
        Else
            excelSheet.Application.Cells(Lcnt, exLAST_SYUKA_DT).Value = Mid(StrConv(tmpP_STOCK_REC.LAST_SYUKA_DT, vbUnicode), 1, 4) & "/" & _
                                                                        Mid(StrConv(tmpP_STOCK_REC.LAST_SYUKA_DT, vbUnicode), 5, 2) & "/" & _
                                                                        Mid(StrConv(tmpP_STOCK_REC.LAST_SYUKA_DT, vbUnicode), 7, 2)
        End If
        '�ŏI�o�א�
        If IsNumeric(StrConv(tmpP_STOCK_REC.LAST_SYUKA_QTY, vbUnicode)) Then
            excelSheet.Application.Cells(Lcnt, exLAST_SYUKA_QTY).Value = CLng(StrConv(tmpP_STOCK_REC.LAST_SYUKA_QTY, vbUnicode))
        Else
            excelSheet.Application.Cells(Lcnt, exLAST_SYUKA_QTY).Value = ""
        End If
        '�m�F(�O��)
        If IsNumeric(StrConv(tmpP_STOCK_REC.MAEGARI_QTY, vbUnicode)) Then
            excelSheet.Application.Cells(Lcnt, exMAEGARI_QTY).Value = CLng(StrConv(tmpP_STOCK_REC.MAEGARI_QTY, vbUnicode))
        Else
            excelSheet.Application.Cells(Lcnt, exMAEGARI_QTY).Value = ""
        End If


        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(tmpP_STOCK_REC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(tmpP_STOCK_REC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(tmpP_STOCK_REC.HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr

            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                Call UniCode_Conv(ITEMREC.ST_RETU, "")
                Call UniCode_Conv(ITEMREC.ST_REN, "")
                Call UniCode_Conv(ITEMREC.ST_DAN, "")
            Case Else
                Call Input_UnLock
                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                Exit Function
        End Select






        Call UniCode_Conv(K1_ZAIKO.JGYOBU, StrConv(tmpP_STOCK_REC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K1_ZAIKO.NAIGAI, StrConv(tmpP_STOCK_REC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K1_ZAIKO.HIN_GAI, StrConv(tmpP_STOCK_REC.HIN_GAI, vbUnicode))

        Call UniCode_Conv(K1_ZAIKO.GOODS_ON, "")
        Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, "")
        Call UniCode_Conv(K1_ZAIKO.SOKO_NO, "")
        Call UniCode_Conv(K1_ZAIKO.Retu, "")
        Call UniCode_Conv(K1_ZAIKO.Ren, "")
        Call UniCode_Conv(K1_ZAIKO.Dan, "")

        com = BtOpGetGreater


        ZAIKO_F = False

        Do

            DoEvents

            sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
            Select Case sts
                Case BtNoErr

                    If StrConv(ZAIKOREC.SOKO_NO, vbUnicode) = StrConv(ITEMREC.ST_SOKO, vbUnicode) And _
                        StrConv(ZAIKOREC.Retu, vbUnicode) = StrConv(ITEMREC.ST_RETU, vbUnicode) And _
                        StrConv(ZAIKOREC.Ren, vbUnicode) = StrConv(ITEMREC.ST_REN, vbUnicode) And _
                        StrConv(ZAIKOREC.Dan, vbUnicode) = StrConv(ITEMREC.ST_DAN, vbUnicode) Then
                    Else
                        ZAIKO_F = True
                        Exit Do
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call Input_UnLock
                    Call File_Error(sts, com, "�݌��ް�")
                    Exit Function
            End Select


        Loop


        If ZAIKO_F Then
            excelSheet.Application.Cells(Lcnt, exLOCATION).Value = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
        Else
            excelSheet.Application.Cells(Lcnt, exLOCATION).Value = ""

        End If





        Lcnt = Lcnt + 1


        com = BtOpGetNext


    Loop




                                            'tmp���ޒI�����b�k�n�r�d    2011.02.14
    sts = BTRV(BtOpClose, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), K0_tmpP_STOCK, Len(K0_tmpP_STOCK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "tmp���ޒI��")
        End If
    End If






    excelApplication.Visible = True                '2016.04.27





'    excelApplication.DisplayAlerts = False


'    excelApplication.Visible = False

'    ExcelApp.Workbooks.Close                                       '�ۑ��m�F�����ŕ���

'    excelApplication.Quit



    Set excelSheet = Nothing                                        'ܰ���ĊJ��
    Set excelWorkBook = Nothing                                     'ܰ��ޯ��J��
    Set excelApplication = Nothing                                  'Excel���ع���� Close

    Call Input_UnLock


    MsgBox "EXCEL�o�͏I���I�I"


    Data_Out_Proc = False
    Exit Function

'2010.12.20 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Data_Out_Proc_Error:
    If Err.Number = 70 Then
        yn = MsgBox("���[���Ŏ��ޒI�����W�v���ׁ̈A���s�ł��܂���" & vbCr & vbLf & _
                    "�Ď��s���܂����H", vbOKCancel + vbExclamation, Err.Source)

        If yn = vbOK Then
            Resume
        End If
    Else
        If Err.Number = 53 Then
            Resume Next
        Else
            MsgBox "[" & Err.Number & "] " & Err.Description, vbOKCancel + vbExclamation, Err.Source
        End If
    End If
'2010.12.20 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function

Private Function ZenZan_Update_Proc() As Integer
'-------------------------------------  �i�ڃ}�X�^���O���c�L�蕪���W�v
Dim sts                     As Integer
Dim com                     As Integer
Dim SYUSHI_ON               As Boolean


Dim i                       As Integer
Dim Upd_Com                 As Integer

Dim wk_Val                  As Double



    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_ITEM.NAIGAI, "")
    Call UniCode_Conv(K0_ITEM.HIN_GAI, "")

    com = BtOpGetGreaterEqual


    Do

        DoEvents
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> SHIZAI Then
                    Exit Do
                End If

            Case BtErrEOF
                Exit Do

            Case Else
                Call File_Error(sts, com, "�i�ڃ}�X�^")
                Exit Function

        End Select

Label3(1).Caption = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))

        SYUSHI_ON = False               '2007.11.13
        If GLB_SYUSHI_F = "" Then       '2007.11.13
            SYUSHI_ON = True
        Else
            SYUSHI_ON = False

            For i = 0 To UBound(G_SYUSHI_TBL)

                If Trim(StrConv(ITEMREC.G_SYUSHI, vbUnicode)) = G_SYUSHI_TBL(i) Then
                    SYUSHI_ON = True
                    Exit For
                End If


            Next i
        End If




        If StrConv(ITEMREC.ZAIKO_F, vbUnicode) <> P_ZAIKO_F_ON Or _
            Not SYUSHI_ON Then                                          '2007.11.13
        Else


            '�W�vں��ޏ���
            Call UniCode_Conv(K0_P_STOCKSUM.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))

            sts = BTRV(BtOpGetEqual, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

            Select Case sts
                Case BtNoErr

                    Upd_Com = BtOpUpdate


                Case BtErrKeyNotFound

                    Upd_Com = BtOpInsert


                Case Else

                    Call File_Error(sts, BtOpGetEqual, "���ޒI�����W�v�ް�")
                    Exit Function

            End Select



            If Upd_Com = BtOpInsert Then
                Call UniCode_Conv(P_STOCKSUM_REC.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))
                Call UniCode_Conv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, "00000000000")
                Call UniCode_Conv(P_STOCKSUM_REC.NYUKO_KIN, "00000000000")
                Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, "00000000000")
                Call UniCode_Conv(P_STOCKSUM_REC.ZAIKO_KIN, "00000000000")
                Call UniCode_Conv(P_STOCKSUM_REC.FILLER, "")



            End If

            If IsNumeric(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode)) Then
                wk_Val = Val(StrConv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, vbUnicode)) + _
                            Val(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode))
            End If

            Call UniCode_Conv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, Format(wk_Val, "00000000000"))

If Upd_Com = BtOpInsert Then
    Call LOG_OUT(LOG_F, "3=" & StrConv(ITEMREC.G_SYUSHI, vbUnicode))
End If
            Do
                sts = BTRV(Upd_Com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

                Select Case sts
                    Case BtNoErr, BtErrDuplicates
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        DoEvents
                    Case Else

                        Call File_Error(sts, BtOpGetEqual, "���ޒI�����W�v�ް�")
                        Exit Function
                End Select


            Loop




            If Not IsNumeric(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode)) Then
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "00000000000")
            End If


            If Not IsNumeric(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode)) Then
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "00000000000")
            End If



            If Not IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)) Then    '2008.02.13
                Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA, "00000000000")
            End If

'Debug.Print StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)


'2006.11.22            If CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode)) = 0 And _
'2006.11.22                CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode)) = 0 Then
'2006.11.22            Else

                Call UniCode_Conv(K0_P_STOCK.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_P_STOCK.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_P_STOCK.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
''2006.11.22                Call UniCode_Conv(K0_P_STOCK.CODE, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
''2006.11.22                Call UniCode_Conv(K0_P_STOCK.TANKA, StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))
                Call UniCode_Conv(K0_P_STOCK.CODE, "")          '2006.11.22
                Call UniCode_Conv(K0_P_STOCK.TANKA, "")         '2006.11.22



                sts = BTRV(BtOpGetEqual, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

                Select Case sts
                    Case BtNoErr

                        Upd_Com = BtOpUpdate


                    Case BtErrKeyNotFound

                        Upd_Com = BtOpInsert


                    Case Else

                        Call File_Error(sts, BtOpGetEqual, "���ޒI�����ް�")
                        Exit Function

                End Select



                If Upd_Com = BtOpInsert Then
                    Call UniCode_Conv(P_STOCK_REC.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(P_STOCK_REC.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(P_STOCK_REC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
''2006.11.22                    Call UniCode_Conv(P_STOCK_REC.CODE, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
''2006.11.22                    Call UniCode_Conv(P_STOCK_REC.TANKA, StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))

                    Call UniCode_Conv(P_STOCK_REC.INPUT_DATE, "")   '2006.11.22
                    Call UniCode_Conv(P_STOCK_REC.CODE, "")         '2006.11.22
                    Call UniCode_Conv(P_STOCK_REC.TANKA, "")        '2006.11.22



                    Call UniCode_Conv(P_STOCK_REC.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))
                    Call UniCode_Conv(P_STOCK_REC.ZEN_ZAIKO_QTY, "00000000")



                    Call UniCode_Conv(P_STOCK_REC.NYUKO_QTY, "00000000")
                    Call UniCode_Conv(P_STOCK_REC.SYUKO_QTY, "00000000")
                    Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, "00000000")


                    Call UniCode_Conv(P_STOCK_REC.LAST_SYUKA_DT, StrConv(ITEMREC.LAST_SYU_DT, vbUnicode))
                    Call UniCode_Conv(P_STOCK_REC.LAST_SYUKA_QTY, StrConv(ITEMREC.G_LAST_SYUKA_QTY, vbUnicode))

                    Call UniCode_Conv(P_STOCK_REC.MOTO_ZAIKO_QTY, "00000000")
                    Call UniCode_Conv(P_STOCK_REC.MAEGARI_QTY, "00000000")

                    Call UniCode_Conv(P_STOCK_REC.SYUKA_NON_F, "0")     '2008.06.21






                    Call UniCode_Conv(P_STOCK_REC.ZEN_ZAIKO_KIN, "00000000")



                    Call UniCode_Conv(P_STOCK_REC.FILLER, "")



                End If

'2009.08.21                wk_VAL = CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)) + _
'                            CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode))


                wk_Val = Val(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)) + _
                            Val(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode))


                Call UniCode_Conv(P_STOCK_REC.ZEN_ZAIKO_QTY, Format(wk_Val, "0000000"))




                wk_Val = Val(StrConv(P_STOCK_REC.ZEN_ZAIKO_KIN, vbUnicode)) + _
                            Val(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode))






                Call UniCode_Conv(P_STOCK_REC.ZEN_ZAIKO_KIN, Format(wk_Val, "0000000"))





                Do
                    sts = BTRV(Upd_Com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

                    Select Case sts
                        Case BtNoErr, BtErrDuplicates
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            DoEvents
                        Case Else

                            Call File_Error(sts, Upd_Com, "���ޒI���ް�")
                            Exit Function
                    End Select


                Loop

            End If

'2006.11.22        End If

        com = BtOpGetNext

    Loop

End Function

Private Function SHIIRE_Update_Proc() As Integer
'-------------------------------------  ���ގ����蓖�����ɂ��W�v
Dim wKEIJYO_YM              As String * 6

Dim com                     As Integer
Dim Upd_Com                 As Integer
Dim sts                     As Integer

Dim SKIP_FLG                As Boolean

Dim SYUSHI_ON               As Boolean

Dim i                       As Integer

Dim wk_Val                  As Double


    wKEIJYO_YM = Left(Text1(ptxKEIJYO_YM).Text, 4) & Right(Text1(ptxKEIJYO_YM).Text, 2)

    Call UniCode_Conv(K2_P_SHUKEIRE.KEIJYO_YM, wKEIJYO_YM)
    Call UniCode_Conv(K2_P_SHUKEIRE.UKEIRE_DT, "")

    com = BtOpGetGreaterEqual


    Do

        DoEvents

        sts = BTRV(com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K2_P_SHUKEIRE, Len(K2_P_SHUKEIRE), 2)

        Select Case sts
            Case BtNoErr

                If StrConv(P_SHUKEIRE_REC.KEIJYO_YM, vbUnicode) <> wKEIJYO_YM Then
                    Exit Do
                End If

            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "���ގ���ް�")
                Exit Function
        End Select



        '�����ް��ǂݍ���
        Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
        SKIP_FLG = False
        Select Case sts
            Case BtNoErr
                '�i�ڂ̍݌Ɍv���׸ނ��`�F�b�N
                Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))

                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                SKIP_FLG = False
                Select Case sts
                    Case BtNoErr



                        If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_OFF Then
                            SKIP_FLG = True       '�l�����Ȃ��̂Ž����
                        End If

                        SYUSHI_ON = False               '2007.11.13
                        If GLB_SYUSHI_F = "" Then       '2007.11.13
                            SYUSHI_ON = True
                        Else
                            SYUSHI_ON = False

                            For i = 0 To UBound(G_SYUSHI_TBL)

                                If Trim(StrConv(ITEMREC.G_SYUSHI, vbUnicode)) = G_SYUSHI_TBL(i) Then
                                    SYUSHI_ON = True
                                    Exit For
                                End If


                            Next i
                        End If

 
                    Case BtErrKeyNotFound


                        SKIP_FLG = True       '�l�����Ȃ��̂Ž����


                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                        Exit Function
                End Select





            Case BtErrKeyNotFound


                SKIP_FLG = True       '�����Ȃ��͒ʏ��ް��ł͂Ȃ�


            Case Else
                Call File_Error(sts, BtOpGetEqual, "���ޒ����ް�")
                Exit Function
        End Select

Label3(1).Caption = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))


        If StrConv(P_SHORDER_REC.G_SHIIRE_KBN, vbUnicode) = ZEI_SHIIRE_KBN Then
            SKIP_FLG = True       '����Ŏd��
        End If

        If Not SKIP_FLG And SYUSHI_ON Then      '2007.11.13
Call LOG_OUT(LOG_F, StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode) & " itemrec.G_SYUSHI=" & StrConv(ITEMREC.G_SYUSHI, vbUnicode) & "P_SHORDER_REC=" & StrConv(P_SHORDER_REC.G_SYUSHI, vbUnicode))

            '�W�vں��ޏ���
            Call UniCode_Conv(K0_P_STOCKSUM.G_SYUSHI, StrConv(P_SHORDER_REC.G_SYUSHI, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

            Select Case sts
                Case BtNoErr

                    Upd_Com = BtOpUpdate


                Case BtErrKeyNotFound

                    Upd_Com = BtOpInsert


                Case Else

                    Call File_Error(sts, BtOpGetEqual, "���ޒI�����W�v�ް�")
                    Exit Function

            End Select


            If Upd_Com = BtOpInsert Then
                Call UniCode_Conv(P_STOCKSUM_REC.G_SYUSHI, StrConv(P_SHORDER_REC.G_SYUSHI, vbUnicode))
                Call UniCode_Conv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, "00000000000")
                Call UniCode_Conv(P_STOCKSUM_REC.NYUKO_KIN, "00000000000")
                Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, "00000000000")
                Call UniCode_Conv(P_STOCKSUM_REC.ZAIKO_KIN, "00000000000")
                Call UniCode_Conv(P_STOCKSUM_REC.FILLER, "")
            End If

'2009.08.21            wk_VAL = CLng(StrConv(P_STOCKSUM_REC.NYUKO_KIN, vbUnicode)) + _
                    CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))


            wk_Val = Val(StrConv(P_STOCKSUM_REC.NYUKO_KIN, vbUnicode)) + _
                    Val(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode))






            If wk_Val > 0 Then

                Call UniCode_Conv(P_STOCKSUM_REC.NYUKO_KIN, Format(wk_Val, "0000000000"))
            Else
                Call UniCode_Conv(P_STOCKSUM_REC.NYUKO_KIN, Format(wk_Val, "000000000"))
            End If



            Do
                sts = BTRV(Upd_Com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

                Select Case sts
                    Case BtNoErr, BtErrDuplicates
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        DoEvents
                    Case Else

                        Call File_Error(sts, Upd_Com, "���ޒI�����W�v�ް�")
                        Exit Function
                End Select


            Loop


            Call UniCode_Conv(K0_P_STOCK.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_P_STOCK.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_P_STOCK.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))



            Call UniCode_Conv(K0_P_STOCK.CODE, StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode))
            '2008.11.24
            Call UniCode_Conv(K0_P_STOCK.TANKA, Format(CDbl(Trim(StrConv(P_SHUKEIRE_REC.UKEIRE_TANKA, vbUnicode))), "00000000.00"))


            sts = BTRV(BtOpGetEqual, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

            Select Case sts
                Case BtNoErr

                    Upd_Com = BtOpUpdate


                Case BtErrKeyNotFound

                    Upd_Com = BtOpInsert


                Case Else

                    Call File_Error(sts, BtOpGetEqual, "���ޒI�����ް�")
                    Exit Function

            End Select


            If Upd_Com = BtOpInsert Then
                Call UniCode_Conv(P_STOCK_REC.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
                Call UniCode_Conv(P_STOCK_REC.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
                Call UniCode_Conv(P_STOCK_REC.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))


                Call UniCode_Conv(P_STOCK_REC.CODE, StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode))
                '2008.11.24
                Call UniCode_Conv(P_STOCK_REC.TANKA, Format(CDbl(Trim(StrConv(P_SHUKEIRE_REC.UKEIRE_TANKA, vbUnicode))), "00000000.00"))

                '2006.11.22
                Call UniCode_Conv(P_STOCK_REC.INPUT_DATE, StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode))


                'Call UniCode_Conv(P_STOCK_REC.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))              '2018.03.30
                Call UniCode_Conv(P_STOCK_REC.G_SYUSHI, StrConv(P_SHORDER_REC.G_SYUSHI, vbUnicode))         '2018.03.30
                Call UniCode_Conv(P_STOCK_REC.ZEN_ZAIKO_QTY, "00000000")



                Call UniCode_Conv(P_STOCK_REC.NYUKO_QTY, "00000000")
                Call UniCode_Conv(P_STOCK_REC.SYUKO_QTY, "00000000")
                Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, "00000000")


                Call UniCode_Conv(P_STOCK_REC.LAST_SYUKA_DT, StrConv(ITEMREC.LAST_SYU_DT, vbUnicode))
                Call UniCode_Conv(P_STOCK_REC.LAST_SYUKA_QTY, StrConv(ITEMREC.G_LAST_SYUKA_QTY, vbUnicode))

                Call UniCode_Conv(P_STOCK_REC.MOTO_ZAIKO_QTY, "00000000")

                Call UniCode_Conv(P_STOCK_REC.MAEGARI_QTY, "00000000")

                Call UniCode_Conv(P_STOCK_REC.SYUKA_NON_F, "0")     '2008.06.21



                Call UniCode_Conv(P_STOCK_REC.ZEN_ZAIKO_KIN, "00000000")

                Call UniCode_Conv(P_STOCK_REC.FILLER, "")



            End If

            '2006.11.22
            If StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode) < StrConv(P_STOCK_REC.INPUT_DATE, vbUnicode) Then
                Call UniCode_Conv(P_STOCK_REC.INPUT_DATE, StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode))
            End If

'If Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)) = "C087" Then
'    Debug.Print
'    Call Log_Out(LOG_F, "A " & StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode) & Val(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode)))
'End If


'2009.08.21            wk_VAL = CLng(CLng(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)) + _
'                        CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode)))


            wk_Val = Val(Val(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)) + _
                        Val(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode)))


            '2010.02.02
            If wk_Val >= 0 Then
                Call UniCode_Conv(P_STOCK_REC.NYUKO_QTY, Format(wk_Val, "00000000"))
            Else
                Call UniCode_Conv(P_STOCK_REC.NYUKO_QTY, Format(wk_Val, "0000000"))
            End If



            Do
                sts = BTRV(Upd_Com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

                Select Case sts
                    Case BtNoErr, BtErrDuplicates
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        DoEvents
                    Case Else

                        Call File_Error(sts, Upd_Com, "���ޒI���ް�")
                        Exit Function
                End Select

            Loop


        End If

        com = BtOpGetNext

    Loop

End Function

Private Function Syuka_F_Update_Proc() As Integer
'------------------------------------------------------�@�o�א��e�Z�b�g 2008.06.21
Dim com                     As Integer
Dim sts                     As Integer

Dim Save_Jgyobu             As String * 1
Dim Save_Naigai             As String * 1
Dim Save_Hin_Gai            As String * 20
Dim Save_CODE               As String * 5
Dim Save_TANKA              As String * 11

Dim wkZaiko_QTY             As Long
Dim wkNYUKO_QTY             As Long

Dim wkZEN_ZAIKO             As Long

Dim Next_Jgyobu             As String
Dim Next_Naigai             As String
Dim Next_Hin_Gai            As String


    Save_Hin_Gai = ""

    com = BtOpGetFirst


    Do

        DoEvents

        sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

        Select Case sts
            Case BtNoErr


            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "���ޒI���ް�")
                Exit Function
        End Select


        If Trim(Save_Hin_Gai) = "" Then


            Save_Jgyobu = StrConv(P_STOCK_REC.JGYOBU, vbUnicode)
            Save_Naigai = StrConv(P_STOCK_REC.NAIGAI, vbUnicode)
            Save_Hin_Gai = StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)

            wkZaiko_QTY = 0
            wkNYUKO_QTY = 0

            If Trim(StrConv(P_STOCK_REC.CODE, vbUnicode)) = "" And _
                Trim(StrConv(P_STOCK_REC.TANKA, vbUnicode)) = "" Then

                If IsNumeric(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)) Then

                    wkZEN_ZAIKO = CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode))

                Else
                    wkZEN_ZAIKO = 0
                End If

            End If

        End If

Label3(1).Caption = Trim(StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))

        If Trim(Save_Jgyobu) <> Trim(StrConv(P_STOCK_REC.JGYOBU, vbUnicode)) Or _
            Trim(Save_Naigai) <> Trim(StrConv(P_STOCK_REC.NAIGAI, vbUnicode)) Or _
            Trim(Save_Hin_Gai) <> Trim(StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)) Then

            Next_Jgyobu = StrConv(P_STOCK_REC.JGYOBU, vbUnicode)
            Next_Naigai = StrConv(P_STOCK_REC.NAIGAI, vbUnicode)
            Next_Hin_Gai = StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)

'Debug.Print "Save_Hin_Gai =  " & Save_Hin_Gai & " Next_Hin_Gai =  " & Next_Hin_Gai


            Call UniCode_Conv(K0_P_STOCK.JGYOBU, Save_Jgyobu)
            Call UniCode_Conv(K0_P_STOCK.NAIGAI, Save_Naigai)

            Call UniCode_Conv(K0_P_STOCK.HIN_GAI, Save_Hin_Gai)

            Call UniCode_Conv(K0_P_STOCK.CODE, "")
            Call UniCode_Conv(K0_P_STOCK.TANKA, "")

            sts = BTRV(BtOpGetEqual, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)




            Select Case sts
                Case BtNoErr

                    If wkNYUKO_QTY = 0 And _
                        wkZaiko_QTY = wkZEN_ZAIKO Then
                        Call UniCode_Conv(P_STOCK_REC.SYUKA_NON_F, "0")
                    Else
                        Call UniCode_Conv(P_STOCK_REC.SYUKA_NON_F, "1")
                    End If

                    sts = BTRV(BtOpUpdate, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
                    If sts <> BtNoErr Then

                        Call File_Error(sts, BtOpUpdate, "���ޒI���ް�")
                        Exit Function

                    End If

                Case BtErrKeyNotFound


                Case Else
                    Call File_Error(sts, com, "���ޒI���ް�")
                    Exit Function
            End Select





            Call UniCode_Conv(K0_P_STOCK.JGYOBU, Next_Jgyobu)
            Call UniCode_Conv(K0_P_STOCK.NAIGAI, Next_Naigai)

            Call UniCode_Conv(K0_P_STOCK.HIN_GAI, Next_Hin_Gai)

            Call UniCode_Conv(K0_P_STOCK.CODE, "")
            Call UniCode_Conv(K0_P_STOCK.TANKA, "")

            sts = BTRV(BtOpGetEqual, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)




            Select Case sts
                Case BtNoErr

                Case Else
                    Call File_Error(sts, BtOpGetEqual, "���ޒI���ް�")
                    Exit Function
            End Select

            If Trim(StrConv(P_STOCK_REC.CODE, vbUnicode)) = "" And _
                Trim(StrConv(P_STOCK_REC.TANKA, vbUnicode)) = "" Then



                If IsNumeric(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)) Then
                    wkZEN_ZAIKO = CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode))
                Else

                    wkZEN_ZAIKO = 0
                End If

            End If


            Save_Hin_Gai = StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)

            wkZaiko_QTY = 0
            wkNYUKO_QTY = 0

        End If



        If Trim(StrConv(P_STOCK_REC.CODE, vbUnicode)) <> "" And _
            Trim(StrConv(P_STOCK_REC.TANKA, vbUnicode)) <> "" Then


            If IsNumeric(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) Then

                wkZaiko_QTY = wkZaiko_QTY + CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))

            End If


            If IsNumeric(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)) Then
                wkNYUKO_QTY = wkNYUKO_QTY + CLng(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode))
            End If

        End If




        com = BtOpGetNext

    Loop


    If Trim(Save_Hin_Gai) <> "" Then
        Call UniCode_Conv(K0_P_STOCK.JGYOBU, Save_Jgyobu)
        Call UniCode_Conv(K0_P_STOCK.NAIGAI, Save_Naigai)

        Call UniCode_Conv(K0_P_STOCK.HIN_GAI, Save_Hin_Gai)

        Call UniCode_Conv(K0_P_STOCK.CODE, "")
        Call UniCode_Conv(K0_P_STOCK.TANKA, "")

        sts = BTRV(BtOpGetEqual, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)




        Select Case sts
            Case BtNoErr

                If wkNYUKO_QTY = 0 And _
                    wkZaiko_QTY = wkZEN_ZAIKO Then
                    Call UniCode_Conv(P_STOCK_REC.SYUKA_NON_F, "0")
                Else
                    Call UniCode_Conv(P_STOCK_REC.SYUKA_NON_F, "1")
                End If

                sts = BTRV(BtOpUpdate, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
                If sts <> BtNoErr Then

                    Call File_Error(sts, BtOpUpdate, "���ޒI���ް�")
                    Exit Function

                End If

            Case BtErrKeyNotFound


            Case Else
                Call File_Error(sts, com, "���ޒI���ް�")
                Exit Function
        End Select
    End If

End Function

Private Function OLD_Syuka_Update_Proc() As Integer

Dim sts                     As Integer
Dim com                     As Integer

Dim wkZEN_ZAIKO             As Long
Dim ZAIKO_F                 As Boolean

Dim Syuko_Non_Flg           As Boolean

Dim wk_Val                  As Double

Dim Save_Hin_Gai            As String * 20
Dim Save_G_Syushi           As String * 3

Dim SKIP_FLG                As Boolean


Dim AFT_Hin_Zaiko_Qty       As Long     '2011.02.22
Dim BEF_Hin_Zaiko_Qty       As Long     '2011.02.22
Dim BEF_Hin_SYUKO_Qty       As Long     '2011.02.22
Dim TOP_Hin_ZENZAN_Qty      As Long     '2011.02.22
Dim BEF_Hin_GAI             As String * 20     '2011.02.22



Dim wkZENZAN_QTY            As Long
Dim wkNYUKO_QTY             As Long
Dim wkSYUKO_QTY             As Long
Dim wkZaiko_QTY             As Long



    com = BtOpGetFirst
    ZAIKO_F = False             '2007.04.26
    wkZEN_ZAIKO = 0             '2007.04.26

    Do
        DoEvents

        '2009.09.30 K0--K1
'        sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
        sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K1_P_STOCK, Len(K1_P_STOCK), 1)

        Select Case sts
            Case BtNoErr


            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "���ޒI���ް�")
                Exit Function
        End Select

'''''''''''''''''   2008.06.21
        If Trim(StrConv(P_STOCK_REC.CODE, vbUnicode)) = "" And _
            Trim(StrConv(P_STOCK_REC.TANKA, vbUnicode)) = "" Then

            If StrConv(P_STOCK_REC.SYUKA_NON_F, vbUnicode) = "0" Then
                Syuko_Non_Flg = False
            Else
                Syuko_Non_Flg = True
            End If
        End If
'''''''''''''''''   2008.06.21






        If Syuko_Non_Flg Then           '2008.06.21


Label3(1).Caption = Trim(StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))

            If Trim(StrConv(P_STOCK_REC.CODE, vbUnicode)) <> "" Or _
                Trim(StrConv(P_STOCK_REC.TANKA, vbUnicode)) <> "" Then
    '2006.11.22            wk_VAL = CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)) + CLng(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)) - (CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)))
    '2006.11.22
                
                
                
''''''''''''''''''''''''''''''''' 2011.02.22
                Call UniCode_Conv(K1_wkP_STOCK.JGYOBU, StrConv(P_STOCK_REC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K1_wkP_STOCK.NAIGAI, StrConv(P_STOCK_REC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K1_wkP_STOCK.HIN_GAI, StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(K1_wkP_STOCK.INPUT_DATE, StrConv(P_STOCK_REC.INPUT_DATE, vbUnicode))
                
                Call UniCode_Conv(K1_wkP_STOCK.CODE, StrConv(P_STOCK_REC.CODE, vbUnicode))
                Call UniCode_Conv(K1_wkP_STOCK.TANKA, StrConv(P_STOCK_REC.TANKA, vbUnicode))
                
                AFT_Hin_Zaiko_Qty = 0
                
                sts = BTRV(BtOpGetGreater, wkP_STOCK_POS, wkP_STOCK_REC, Len(wkP_STOCK_REC), K1_wkP_STOCK, Len(K1_wkP_STOCK), 1)
        
                Select Case sts
                    Case BtNoErr
        
                                
                        If StrConv(P_STOCK_REC.JGYOBU, vbUnicode) = StrConv(wkP_STOCK_REC.JGYOBU, vbUnicode) And _
                            StrConv(P_STOCK_REC.NAIGAI, vbUnicode) = StrConv(wkP_STOCK_REC.NAIGAI, vbUnicode) And _
                            StrConv(P_STOCK_REC.HIN_GAI, vbUnicode) = StrConv(wkP_STOCK_REC.HIN_GAI, vbUnicode) Then
                            

                            
                            
                            
                            
                            
''                            If Val(StrConv(wkP_STOCK_REC.NYUKO_QTY, vbUnicode)) <> 0 Then
''                                If Val(StrConv(wkP_STOCK_REC.ZAIKO_QTY, vbUnicode)) <= Val(StrConv(wkP_STOCK_REC.NYUKO_QTY, vbUnicode)) Then
''                                Else
                            
                                    AFT_Hin_Zaiko_Qty = Val(StrConv(wkP_STOCK_REC.ZAIKO_QTY, vbUnicode))
''                                End If
                                
''                            End If
                        End If
                                
                    Case BtErrEOF
        
        
        
                    Case Else
                        Call File_Error(sts, com, "���ޒI���ް�")
                        Exit Function
                End Select
                
                
                
                
                
                Call UniCode_Conv(K1_wkP_STOCK.JGYOBU, StrConv(P_STOCK_REC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K1_wkP_STOCK.NAIGAI, StrConv(P_STOCK_REC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K1_wkP_STOCK.HIN_GAI, StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(K1_wkP_STOCK.INPUT_DATE, StrConv(P_STOCK_REC.INPUT_DATE, vbUnicode))
                Call UniCode_Conv(K1_wkP_STOCK.CODE, StrConv(P_STOCK_REC.CODE, vbUnicode))
                Call UniCode_Conv(K1_wkP_STOCK.TANKA, StrConv(P_STOCK_REC.TANKA, vbUnicode))
                
                BEF_Hin_Zaiko_Qty = 0
                BEF_Hin_SYUKO_Qty = 0
''                TOP_Hin_ZENZAN_Qty = 0
                                
                BEF_Hin_GAI = ""
                
                sts = BTRV(BtOpGetLess, wkP_STOCK_POS, wkP_STOCK_REC, Len(wkP_STOCK_REC), K1_wkP_STOCK, Len(K1_wkP_STOCK), 1)
        
                Select Case sts
                    Case BtNoErr
        
                                
                        If StrConv(P_STOCK_REC.JGYOBU, vbUnicode) = StrConv(wkP_STOCK_REC.JGYOBU, vbUnicode) And _
                            StrConv(P_STOCK_REC.NAIGAI, vbUnicode) = StrConv(wkP_STOCK_REC.NAIGAI, vbUnicode) And _
                            StrConv(P_STOCK_REC.HIN_GAI, vbUnicode) = StrConv(wkP_STOCK_REC.HIN_GAI, vbUnicode) Then
                            

                                
                            If AFT_Hin_Zaiko_Qty = 0 Then
                                
                                If Trim(StrConv(wkP_STOCK_REC.CODE, vbUnicode)) <> "" Or _
                                    Trim(StrConv(wkP_STOCK_REC.TANKA, vbUnicode)) <> "" Then
                                
                                
                                        BEF_Hin_Zaiko_Qty = Val(StrConv(wkP_STOCK_REC.ZAIKO_QTY, vbUnicode))
                                        BEF_Hin_SYUKO_Qty = Val(StrConv(wkP_STOCK_REC.SYUKO_QTY, vbUnicode))
                                        
                                        AFT_Hin_Zaiko_Qty = Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) * -1
                                
                                                    
                                
                                Else
                                        TOP_Hin_ZENZAN_Qty = Val(StrConv(wkP_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode))
                                End If
                            
                            
                                
                                AFT_Hin_Zaiko_Qty = (BEF_Hin_Zaiko_Qty * -1) + TOP_Hin_ZENZAN_Qty + Val(StrConv(wkP_STOCK_REC.NYUKO_QTY, vbUnicode)) - Val(StrConv(wkP_STOCK_REC.SYUKO_QTY, vbUnicode))
                                TOP_Hin_ZENZAN_Qty = 0
                            End If
                        
                        End If
                        BEF_Hin_GAI = StrConv(wkP_STOCK_REC.HIN_GAI, vbUnicode)
                                
                    Case BtErrEOF
        
        
        
                    Case Else
                        Call File_Error(sts, com, "���ޒI���ް�")
                        Exit Function
                End Select
                
                
                If wkZEN_ZAIKO > (Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))) + AFT_Hin_Zaiko_Qty Then
                
                
                        wk_Val = wkZEN_ZAIKO + Val(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)) - (Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))) - AFT_Hin_Zaiko_Qty
                    
                        If BEF_Hin_Zaiko_Qty > BEF_Hin_SYUKO_Qty Then
                            wk_Val = wkZEN_ZAIKO + Val(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)) - (Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)))
                    
                        End If
                Else
                    wk_Val = wkZEN_ZAIKO + Val(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)) - (Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)))
                
                    If BEF_Hin_Zaiko_Qty > BEF_Hin_SYUKO_Qty Then
                        wk_Val = wkZEN_ZAIKO + Val(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)) - (Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))) - AFT_Hin_Zaiko_Qty
                    End If
                End If
                If Val(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)) - (Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))) = 0 Then
                    wk_Val = 0
                End If


 '               wk_VAL = wkZEN_ZAIKO + Val(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)) - (Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)))
                
                
                
'                If Trim(BEF_Hin_GAI) = Trim(StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)) Then
''                    If Trim(StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)) = "D001" Then
''                    wk_VAL = wkZEN_ZAIKO + Val(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)) - (Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))) - AFT_Hin_Zaiko_Qty
''                Else
''                    wk_VAL = wkZEN_ZAIKO + Val(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)) - (Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)))
''                End If

''''''''''''''''''''''''''''''''' 2011.02.22
                
                wkZEN_ZAIKO = 0




                If wk_Val < 0 Then
                    Call UniCode_Conv(P_STOCK_REC.SYUKO_QTY, Format(wk_Val, "0000000"))
                Else
                    Call UniCode_Conv(P_STOCK_REC.SYUKO_QTY, Format(wk_Val, "00000000"))
                End If
                Do



        '2009.09.30 K0--K1
'                    sts = BTRV(BtOpUpdate, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
                    sts = BTRV(BtOpUpdate, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K1_P_STOCK, Len(K1_P_STOCK), 1)

                    Select Case sts
                        Case BtNoErr
                            Exit Do

                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            DoEvents


                        Case Else
                            Call File_Error(sts, BtOpUpdate, "���ޒI���ް�")
                            Exit Function
                    End Select

                Loop

'                ZAIKO_F = True              '2007.04.26
                '2009.09.30 True--False
                ZAIKO_F = False              '2007.04.26


            Else        '2006.11.22

                If wkZEN_ZAIKO <> 0 And Not ZAIKO_F Then


                    Call UniCode_Conv(P_STOCK_REC.JGYOBU, SHIZAI)
                    Call UniCode_Conv(P_STOCK_REC.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(P_STOCK_REC.HIN_GAI, Save_Hin_Gai)


                    SKIP_FLG = False


                    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, Save_Hin_Gai)

                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    SKIP_FLG = False
                    Select Case sts
                        Case BtNoErr

                        Case BtErrKeyNotFound
                            SKIP_FLG = True       '�l�����Ȃ��̂Ž����

                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                            Exit Function
                    End Select




                    Call UniCode_Conv(P_STOCK_REC.CODE, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))


                    If Not IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)) Then
                        Call UniCode_Conv(P_STOCK_REC.TANKA, "00000000.00")
                    Else
                        Call UniCode_Conv(P_STOCK_REC.TANKA, StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))
                    End If

                    Call UniCode_Conv(P_STOCK_REC.INPUT_DATE, "00000000")

                    Call UniCode_Conv(P_STOCK_REC.G_SYUSHI, Save_G_Syushi)

                    Call UniCode_Conv(P_STOCK_REC.ZEN_ZAIKO_QTY, "00000000")

                    Call UniCode_Conv(P_STOCK_REC.NYUKO_QTY, "00000000")

                    Call UniCode_Conv(P_STOCK_REC.SYUKO_QTY, Format(wkZEN_ZAIKO, "00000000"))

                    Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, "00000000")

                    Call UniCode_Conv(P_STOCK_REC.LAST_SYUKA_DT, StrConv(ITEMREC.LAST_SYU_DT, vbUnicode))
                    Call UniCode_Conv(P_STOCK_REC.LAST_SYUKA_QTY, StrConv(ITEMREC.G_LAST_SYUKA_QTY, vbUnicode))

                    Call UniCode_Conv(P_STOCK_REC.MOTO_ZAIKO_QTY, "00000000")
                    Call UniCode_Conv(P_STOCK_REC.MAEGARI_QTY, "00000000")




                    Call UniCode_Conv(P_STOCK_REC.ZEN_ZAIKO_KIN, "00000000")

                    Call UniCode_Conv(P_STOCK_REC.FILLER, "")


                    Do

                        '2009.09.30 K0--K1
'                        sts = BTRV(BtOpInsert, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
                        sts = BTRV(BtOpInsert, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K1_P_STOCK, Len(K1_P_STOCK), 1)

                        Select Case sts
                            Case BtNoErr, BtErrDuplicates
                                Exit Do

                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                DoEvents


'                            Case BtErrDuplicates




                            Case Else


                                Call File_Error(sts, BtOpInsert, "���ޒI���ް�")
                                Exit Function
                        End Select

                    Loop


                End If

                wkZEN_ZAIKO = Val(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode))
                ZAIKO_F = False             '2007.04.26

                Save_Hin_Gai = StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)
                Save_G_Syushi = StrConv(P_STOCK_REC.G_SYUSHI, vbUnicode)

                '2011.02.22
                TOP_Hin_ZENZAN_Qty = Val(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode))


            End If

        End If
        
        BEF_Hin_GAI = StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)
        
        
        
        
        com = BtOpGetNext
    Loop
'----------------------------------------------- 2011.02.22
    com = BtOpGetFirst

    BEF_Hin_GAI = ""

    Do
        
        DoEvents
        sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K1_P_STOCK, Len(K1_P_STOCK), 1)

        Select Case sts
            Case BtNoErr

            Case BtErrEOF
                Exit Do

            Case Else
                Call File_Error(sts, BtOpGetEqual, "�I���ް�")
                Exit Function
        End Select

        If Trim(StrConv(P_STOCK_REC.CODE, vbUnicode)) = "" And _
            Trim(StrConv(P_STOCK_REC.TANKA, vbUnicode)) = "" Then
        
            BEF_Hin_GAI = StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)

            wkZENZAN_QTY = Val(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode))
            wkNYUKO_QTY = 0
            wkSYUKO_QTY = 0
            wkZaiko_QTY = 0

        Else
                    
        End If
        com = BtOpGetNext
    
    Loop
End Function

Private Function Total_Update_Proc() As Integer
'
'�݌ɋ��z�ďW�v
'
Dim GK_ZEN_ZAIKO_KIN        As Long
Dim GK_NYUKO_KIN            As Long
Dim GK_SYUKO_KIN            As Long
Dim GK_ZAIKO_KIN            As Long

Dim com                     As Integer
Dim sts                     As Integer
Dim Upd_Com                 As Integer

Dim wk_Val                  As Double


Dim GK_ZEN_ZAIKO_QTY        As Long
Dim GK_ZAIKO_QTY            As Long


'2010.01.14
Dim Sum_Zen_Zaiko           As Long
Dim Sum_Zaiko               As Long
Dim Sum_Nyuko               As Long
Dim Sum_Syuko               As Long

Dim Sum_Zaiko_KIN           As Long


Dim svJGYOBU                As String * 1
Dim svNAIGAI                As String * 1
Dim svHIN_GAI               As String * 20

Dim svG_SYUSHI              As String * 3

'2010.01.14




    com = BtOpGetFirst


Label3(0).Caption = "�݌ɋ��z�ďW�v"


        '2009.01.14
    svJGYOBU = ""


    Do

        DoEvents

        sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

        Select Case sts
            Case BtNoErr

            Case BtErrEOF
                Exit Do

            Case Else
                Call File_Error(sts, com, "���ޒI�����ް�")
                Exit Function
        End Select



        If Trim(svJGYOBU) = "" Then
            svJGYOBU = StrConv(P_STOCK_REC.JGYOBU, vbUnicode)
            svNAIGAI = StrConv(P_STOCK_REC.NAIGAI, vbUnicode)
            svHIN_GAI = StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)

            svG_SYUSHI = StrConv(P_STOCK_REC.G_SYUSHI, vbUnicode)

            Sum_Zen_Zaiko = 0
            Sum_Zaiko = 0
            Sum_Nyuko = 0
            Sum_Syuko = 0
            Sum_Zaiko_KIN = 0
        End If




        If Trim(svJGYOBU) <> Trim(StrConv(P_STOCK_REC.JGYOBU, vbUnicode)) Or _
           Trim(svNAIGAI) <> Trim(StrConv(P_STOCK_REC.NAIGAI, vbUnicode)) Or _
           Trim(svHIN_GAI) <> Trim(StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)) Then


            If Sum_Zen_Zaiko = Sum_Zaiko And _
                (Sum_Nyuko = 0 And Sum_Syuko = 0) Then


                Call UniCode_Conv(K0_ITEM.JGYOBU, svJGYOBU)
                Call UniCode_Conv(K0_ITEM.NAIGAI, svNAIGAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, svHIN_GAI)

                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

                Select Case sts
                    Case BtNoErr
                        If IsNumeric(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode)) Then
                            Sum_Zaiko_KIN = CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode))
                        Else
                            Sum_Zaiko_KIN = 0
                        End If

                    Case BtErrKeyNotFound
                        Sum_Zaiko_KIN = 0

                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                        Exit Function

                End Select

                Call UniCode_Conv(K0_P_STOCKSUM.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))




                sts = BTRV(BtOpGetEqual, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

                Select Case sts
                    Case BtNoErr

                        Upd_Com = BtOpUpdate


                    Case BtErrKeyNotFound

                        Upd_Com = BtOpInsert


                    Case Else

                        Call File_Error(sts, BtOpGetEqual, "���ޒI�����W�v�ް�")
                        Exit Function

                End Select



                If Upd_Com = BtOpInsert Then
                    
                    
                    
                    Call UniCode_Conv(P_STOCKSUM_REC.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))

                    Call UniCode_Conv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, "00000000000")
                    Call UniCode_Conv(P_STOCKSUM_REC.NYUKO_KIN, "00000000000")
                    Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, "00000000000")
                    Call UniCode_Conv(P_STOCKSUM_REC.ZAIKO_KIN, "00000000000")
                    Call UniCode_Conv(P_STOCKSUM_REC.FILLER, "")



                End If

                wk_Val = Val(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode)) + Sum_Zaiko_KIN

                Call UniCode_Conv(P_STOCKSUM_REC.ZAIKO_KIN, Format(wk_Val, "00000000000"))


                wk_Val = Val(StrConv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, vbUnicode)) + _
                         Val(StrConv(P_STOCKSUM_REC.NYUKO_KIN, vbUnicode)) - _
                        (Val(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode)))
                If wk_Val < 0 Then
                    Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, Format(wk_Val, "0000000"))
                Else
                    Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, Format(wk_Val, "00000000"))
                End If

If Upd_Com = BtOpInsert Then
    Call LOG_OUT(LOG_F, "1=" & StrConv(P_STOCKSUM_REC.G_SYUSHI, vbUnicode))
End If
                Do
                    sts = BTRV(Upd_Com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

                    Select Case sts
                        Case BtNoErr, BtErrDuplicates
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            DoEvents
                        Case Else

                            Call File_Error(sts, Upd_Com, "���ޒI�����W�v�ް�")
                            Exit Function
                    End Select


                Loop


            Else

                Call UniCode_Conv(K0_wkP_STOCK.JGYOBU, svJGYOBU)
                Call UniCode_Conv(K0_wkP_STOCK.NAIGAI, svNAIGAI)
                Call UniCode_Conv(K0_wkP_STOCK.HIN_GAI, svHIN_GAI)

                Call UniCode_Conv(K0_wkP_STOCK.CODE, "")
                Call UniCode_Conv(K0_wkP_STOCK.TANKA, "")


                com = BtOpGetGreater

                Do

                    DoEvents

                    sts = BTRV(com, wkP_STOCK_POS, wkP_STOCK_REC, Len(wkP_STOCK_REC), K0_wkP_STOCK, Len(K0_wkP_STOCK), 0)

                    Select Case sts
                        Case BtNoErr


                        Case BtErrEOF

                            Exit Do


                        Case Else
                            Call File_Error(sts, com, "���ޒI�����ް�")
                            Exit Function
                    End Select



                    If Trim(svJGYOBU) <> Trim(StrConv(wkP_STOCK_REC.JGYOBU, vbUnicode)) Or _
                        Trim(svNAIGAI) <> Trim(StrConv(wkP_STOCK_REC.NAIGAI, vbUnicode)) Or _
                        Trim(svHIN_GAI) <> Trim(StrConv(wkP_STOCK_REC.HIN_GAI, vbUnicode)) Then

                        Exit Do

                    End If



                    Call UniCode_Conv(K0_P_STOCKSUM.G_SYUSHI, StrConv(wkP_STOCK_REC.G_SYUSHI, vbUnicode))



                    sts = BTRV(BtOpGetEqual, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

                    Select Case sts
                        Case BtNoErr

                            Upd_Com = BtOpUpdate


                        Case BtErrKeyNotFound

                            Upd_Com = BtOpInsert


                        Case Else

                            Call File_Error(sts, BtOpGetEqual, "���ޒI�����W�v�ް�")
                            Exit Function

                    End Select



                    If Upd_Com = BtOpInsert Then
            
            
            
            '            Call UniCode_Conv(P_STOCKSUM_REC.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))
                        Call UniCode_Conv(P_STOCKSUM_REC.G_SYUSHI, StrConv(wkP_STOCK_REC.G_SYUSHI, vbUnicode))

                        Call UniCode_Conv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, "00000000000")
                        Call UniCode_Conv(P_STOCKSUM_REC.NYUKO_KIN, "00000000000")
                        Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, "00000000000")
                        Call UniCode_Conv(P_STOCKSUM_REC.ZAIKO_KIN, "00000000000")
                        Call UniCode_Conv(P_STOCKSUM_REC.FILLER, "")



                    End If




'>>>>>>>    2018.07.25
                    If IsNumeric(StrConv(wkP_STOCK_REC.TANKA, vbUnicode)) Then



                        wk_Val = Val(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode)) + _
                                Int(CDbl(CDbl(StrConv(wkP_STOCK_REC.TANKA, vbUnicode)) * _
                                          Val(StrConv(wkP_STOCK_REC.ZAIKO_QTY, vbUnicode)) + 0.99))



                        Call UniCode_Conv(P_STOCKSUM_REC.ZAIKO_KIN, Format(wk_Val, "00000000000"))
                    End If



'    If IsNumeric(StrConv(P_STOCK_REC.TANKA, vbUnicode)) Then
''        STOCK(Row, colZAIKO_KIN) = Format(CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) * _
''                                    CDbl(StrConv(P_STOCK_REC.TANKA, vbUnicode)), "#,##0")
'
'2018.08.21        If Not IsNumeric(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) Then
'2018.08.21            Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, "00000000")
'2018.08.21        End If
'2018.08.21        If Not IsNumeric(StrConv(P_STOCK_REC.TANKA, vbUnicode)) Then
'2018.08.21            Call UniCode_Conv(P_STOCK_REC.TANKA, "00000000")
'2018.08.21        End If
'2018.08.21
'2018.08.21        wk_Val = Format(ToRoundUp(CCur(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) * _
                                    CCur(StrConv(P_STOCK_REC.TANKA, vbUnicode)), 0), "#,##0")





'2018.08.21        If Not IsNumeric(StrConv(wkP_STOCK_REC.ZAIKO_QTY, vbUnicode)) Then
'2018.08.21            Call UniCode_Conv(wkP_STOCK_REC.ZAIKO_QTY, "00000000")
'2018.08.21        End If
'2018.08.21        If Not IsNumeric(StrConv(wkP_STOCK_REC.TANKA, vbUnicode)) Then
'2018.08.21            Call UniCode_Conv(wkP_STOCK_REC.TANKA, "00000000")
'2018.08.21        End If

'2018.08.21        wk_Val = Format(ToRoundUp(CCur(StrConv(wkP_STOCK_REC.ZAIKO_QTY, vbUnicode)) * _
                                    CCur(StrConv(wkP_STOCK_REC.TANKA, vbUnicode)), 0), "#,##0")




'2018.08.21       Call UniCode_Conv(P_STOCKSUM_REC.ZAIKO_KIN, Format(wk_Val, "00000000000"))
'    End If







'>>>>>>>    2018.07.25



                    wk_Val = Val(StrConv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, vbUnicode)) + _
                             Val(StrConv(P_STOCKSUM_REC.NYUKO_KIN, vbUnicode)) - _
                            (Val(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode)))
                    If wk_Val < 0 Then
                        Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, Format(wk_Val, "0000000"))
                    Else
                        Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, Format(wk_Val, "00000000"))
                    End If

If Upd_Com = BtOpInsert Then
    Call LOG_OUT(LOG_F, "777=" & StrConv(P_STOCKSUM_REC.G_SYUSHI, vbUnicode))
End If


                    Do
                        sts = BTRV(Upd_Com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

                        Select Case sts
                            Case BtNoErr, BtErrDuplicates
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                DoEvents
                            Case Else

                                Call File_Error(sts, Upd_Com, "���ޒI�����W�v�ް�")
                                Exit Function
                        End Select


                    Loop


                    com = BtOpGetNext

                Loop



            End If


            svJGYOBU = StrConv(P_STOCK_REC.JGYOBU, vbUnicode)
            svNAIGAI = StrConv(P_STOCK_REC.NAIGAI, vbUnicode)
            svHIN_GAI = StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)

            svG_SYUSHI = StrConv(P_STOCK_REC.G_SYUSHI, vbUnicode)


            Sum_Zen_Zaiko = 0
            Sum_Zaiko = 0
            Sum_Nyuko = 0
            Sum_Syuko = 0
            Sum_Zaiko_KIN = 0



        End If



        Sum_Zen_Zaiko = Sum_Zen_Zaiko + CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode))
        
        Sum_Zaiko = Sum_Zaiko + CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))
        Sum_Nyuko = Sum_Nyuko + CLng(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode))
        Sum_Syuko = Sum_Syuko + CLng(StrConv(P_STOCK_REC.SYUKO_QTY, vbUnicode))

        com = BtOpGetNext


    Loop



    If Trim(svJGYOBU) <> "" Then


        If Sum_Zen_Zaiko = Sum_Zaiko And _
            (Sum_Nyuko = 0 And Sum_Syuko = 0) Then


            Call UniCode_Conv(K0_ITEM.JGYOBU, svJGYOBU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, svNAIGAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, svHIN_GAI)

            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

            Select Case sts
                Case BtNoErr
                    If IsNumeric(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode)) Then
                        Sum_Zaiko_KIN = CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode))
                    Else
                        Sum_Zaiko_KIN = 0
                    End If

                Case BtErrKeyNotFound
                    Sum_Zaiko_KIN = 0

                Case Else

                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                    Exit Function

            End Select

            Call UniCode_Conv(K0_P_STOCKSUM.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))

            sts = BTRV(BtOpGetEqual, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

            Select Case sts
                Case BtNoErr

                    Upd_Com = BtOpUpdate


                Case BtErrKeyNotFound

                    Upd_Com = BtOpInsert


                Case Else

                    Call File_Error(sts, BtOpGetEqual, "���ޒI�����W�v�ް�")
                    Exit Function

            End Select



            If Upd_Com = BtOpInsert Then
                Call UniCode_Conv(P_STOCKSUM_REC.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))

                Call UniCode_Conv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, "00000000000")
                Call UniCode_Conv(P_STOCKSUM_REC.NYUKO_KIN, "00000000000")
                Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, "00000000000")
                Call UniCode_Conv(P_STOCKSUM_REC.ZAIKO_KIN, "00000000000")
                Call UniCode_Conv(P_STOCKSUM_REC.FILLER, "")
            End If

            wk_Val = Val(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode)) + Sum_Zaiko_KIN

            Call UniCode_Conv(P_STOCKSUM_REC.ZAIKO_KIN, Format(wk_Val, "00000000000"))


            wk_Val = Val(StrConv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, vbUnicode)) + Val(StrConv(P_STOCKSUM_REC.NYUKO_KIN, vbUnicode)) - (Val(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode)))
            If wk_Val < 0 Then
                Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, Format(wk_Val, "0000000"))
            Else
                Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, Format(wk_Val, "00000000"))
            End If
If Upd_Com = BtOpInsert Then
    Call LOG_OUT(LOG_F, "99=" & StrConv(P_STOCKSUM_REC.G_SYUSHI, vbUnicode))
End If


            Do
                sts = BTRV(Upd_Com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

                Select Case sts
                    Case BtNoErr, BtErrDuplicates
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        DoEvents
                    Case Else

                        Call File_Error(sts, Upd_Com, "���ޒI�����W�v�ް�")
                        Exit Function
                End Select


            Loop


        Else

            Call UniCode_Conv(K0_wkP_STOCK.JGYOBU, svJGYOBU)
            Call UniCode_Conv(K0_wkP_STOCK.NAIGAI, svNAIGAI)
            Call UniCode_Conv(K0_wkP_STOCK.HIN_GAI, svHIN_GAI)
            Call UniCode_Conv(K0_wkP_STOCK.CODE, "")
            Call UniCode_Conv(K0_wkP_STOCK.TANKA, "")

            com = BtOpGetGreater

            Do

                DoEvents

                sts = BTRV(com, wkP_STOCK_POS, wkP_STOCK_REC, Len(wkP_STOCK_REC), K0_wkP_STOCK, Len(K0_wkP_STOCK), 0)

                Select Case sts
                    Case BtNoErr


                    Case BtErrEOF

                        Exit Do


                    Case Else
                        Call File_Error(sts, com, "���ޒI�����ް�")
                        Exit Function
                End Select


                If Trim(svJGYOBU) <> Trim(StrConv(wkP_STOCK_REC.JGYOBU, vbUnicode)) Or _
                    Trim(svNAIGAI) <> Trim(StrConv(wkP_STOCK_REC.NAIGAI, vbUnicode)) Or _
                    Trim(svHIN_GAI) <> Trim(StrConv(wkP_STOCK_REC.HIN_GAI, vbUnicode)) Then

                    Exit Do

                End If


                Call UniCode_Conv(K0_P_STOCKSUM.G_SYUSHI, StrConv(wkP_STOCK_REC.G_SYUSHI, vbUnicode))

                sts = BTRV(BtOpGetEqual, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

                Select Case sts
                    Case BtNoErr
                        Upd_Com = BtOpUpdate

                    Case BtErrKeyNotFound
                        Upd_Com = BtOpInsert

                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "���ޒI�����W�v�ް�")
                        Exit Function

                End Select
Call LOG_OUT(LOG_F, "2=" & StrConv(P_STOCKSUM_REC.G_SYUSHI, vbUnicode))


                If Upd_Com = BtOpInsert Then
        '            Call UniCode_Conv(P_STOCKSUM_REC.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))
                    Call UniCode_Conv(P_STOCKSUM_REC.G_SYUSHI, StrConv(wkP_STOCK_REC.G_SYUSHI, vbUnicode))

                    Call UniCode_Conv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, "00000000000")
                    Call UniCode_Conv(P_STOCKSUM_REC.NYUKO_KIN, "00000000000")
                    Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, "00000000000")
                    Call UniCode_Conv(P_STOCKSUM_REC.ZAIKO_KIN, "00000000000")
                    Call UniCode_Conv(P_STOCKSUM_REC.FILLER, "")
                End If


                If IsNumeric(StrConv(wkP_STOCK_REC.TANKA, vbUnicode)) Then

                    wk_Val = Val(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode)) + _
                            Int(CDbl(CDbl(StrConv(wkP_STOCK_REC.TANKA, vbUnicode)) * Val(StrConv(wkP_STOCK_REC.ZAIKO_QTY, vbUnicode)) + 0.99))

                    Call UniCode_Conv(P_STOCKSUM_REC.ZAIKO_KIN, Format(wk_Val, "00000000000"))
                End If

                wk_Val = Val(StrConv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, vbUnicode)) + Val(StrConv(P_STOCKSUM_REC.NYUKO_KIN, vbUnicode)) - (Val(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode)))
                If wk_Val < 0 Then
                    Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, Format(wk_Val, "0000000"))
                Else
                    Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, Format(wk_Val, "00000000"))
                End If

If Upd_Com = BtOpInsert Then
    Call LOG_OUT(LOG_F, "100=" & StrConv(P_STOCKSUM_REC.G_SYUSHI, vbUnicode))
End If

                Do
                    sts = BTRV(Upd_Com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

                    Select Case sts
                        Case BtNoErr, BtErrDuplicates
                            Exit Do

                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            DoEvents

                        Case Else
                            Call File_Error(sts, Upd_Com, "���ޒI�����W�v�ް�")
                            Exit Function

                    End Select

                Loop


                com = BtOpGetNext

            Loop


        End If


        svJGYOBU = StrConv(P_STOCK_REC.JGYOBU, vbUnicode)
        svNAIGAI = StrConv(P_STOCK_REC.NAIGAI, vbUnicode)
        svHIN_GAI = StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)

        svG_SYUSHI = StrConv(P_STOCK_REC.G_SYUSHI, vbUnicode)


        Sum_Zen_Zaiko = 0
        Sum_Zaiko = 0
        Sum_Nyuko = 0
        Sum_Syuko = 0
        Sum_Zaiko_KIN = 0



    End If






        '2009.01.14



'        Call UniCode_Conv(K0_P_STOCKSUM.G_SYUSHI, StrConv(P_STOCK_REC.G_SYUSHI, vbUnicode))
'
'        sts = BTRV(BtOpGetEqual, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)
'
'        Select Case sts
'            Case BtNoErr
'
'                upd_com = BtOpUpdate
'
'
'            Case BtErrKeyNotFound
'
'                upd_com = BtOpInsert
'
'
'            Case Else
'
'                Call File_Error(sts, BtOpGetEqual, "���ޒI�����W�v�ް�")
'                Exit Function
'
'        End Select
'
'
'
'        If upd_com = BtOpInsert Then
'            Call UniCode_Conv(P_STOCKSUM_REC.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))
'            Call UniCode_Conv(P_STOCKSUM_REC.G_SYUSHI, StrConv(P_STOCK_REC.G_SYUSHI, vbUnicode))
'
'            Call UniCode_Conv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, "00000000000")
'            Call UniCode_Conv(P_STOCKSUM_REC.NYUKO_KIN, "00000000000")
'            Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, "00000000000")
'            Call UniCode_Conv(P_STOCKSUM_REC.ZAIKO_KIN, "00000000000")
'            Call UniCode_Conv(P_STOCKSUM_REC.FILLER, "")
'
'
'
'        End If
'
'
'
'
'
'        If IsNumeric(StrConv(P_STOCK_REC.TANKA, vbUnicode)) Then
'
'
'
'            '0.5-->0.9 2009.08.25
''            wk_VAL = Val(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode)) + _
''                    Int(CDbl(CDbl(StrConv(P_STOCK_REC.TANKA, vbUnicode)) * Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) + 0.9))
'
'            wk_VAL = Val(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode)) + _
'                    Int(CDbl(CDbl(StrConv(P_STOCK_REC.TANKA, vbUnicode)) * Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) + 0.99))
'
'
'
'            Call UniCode_Conv(P_STOCKSUM_REC.ZAIKO_KIN, Format(wk_VAL, "00000000000"))
'        End If
'
'        wk_VAL = Val(StrConv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, vbUnicode)) + Val(StrConv(P_STOCKSUM_REC.NYUKO_KIN, vbUnicode)) - (Val(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode)))
'        If wk_VAL < 0 Then
'            Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, Format(wk_VAL, "0000000"))
'        Else
'            Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, Format(wk_VAL, "00000000"))
'        End If
'
'
'
'        Do
'            sts = BTRV(upd_com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)
'
'            Select Case sts
'                Case BtNoErr, BtErrDuplicates
'                    Exit Do
'                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                    DoEvents
'                Case Else
'
'                    Call File_Error(sts, upd_com, "���ޒI�����W�v�ް�")
'                    Exit Function
'            End Select
'
'
'        Loop
'
'        com = BtOpGetNext
'
'    Loop




    '-------------------------------------  �I�[���[�����R�[�h�폜  2006.11.22
    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

        Select Case sts
            Case BtNoErr


            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "���ޒI�����ް�")
                Exit Function
        End Select


        If Trim(StrConv(P_STOCK_REC.CODE, vbUnicode)) <> "" And _
            Trim(StrConv(P_STOCK_REC.TANKA, vbUnicode)) <> "" Then
            'Clng --> Val 2016.01.08
            If Val(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)) = 0 And _
                Val(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)) = 0 And _
                Val(StrConv(P_STOCK_REC.SYUKO_QTY, vbUnicode)) = 0 And _
                Val(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) = 0 Then

                Do
                    sts = BTRV(BtOpDelete, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            DoEvents
                        Case Else

                            Call File_Error(sts, Upd_Com, "���ޒI�����ް�")
                            Exit Function
                    End Select


                Loop
            End If
        End If

        com = BtOpGetNext

    Loop

'>>>>>>>>   2018.01.22

    '-------------------------------------  ���ٕ������Z�b�g


1    com = BtOpGetEqual             '2016.01.07
'    com = BtOpGetFirst              '2016.01.07
    Do
        DoEvents

        sts = BTRV(com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

        Select Case sts
            Case BtNoErr


            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "���ޒI���W�v�ް�")
                Exit Function
        End Select






        GK_ZEN_ZAIKO_QTY = 0
        GK_ZAIKO_QTY = 0


        com = BtOpGetFirst

        Do
            DoEvents


            sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

            Select Case sts
                Case BtNoErr


                Case BtErrEOF

                    Exit Do


                Case Else
                    Call File_Error(sts, com, "���ޒI���ް�")
                    Exit Function
            End Select


            If StrConv(P_STOCK_REC.G_SYUSHI, vbUnicode) = StrConv(P_STOCKSUM_REC.G_SYUSHI, vbUnicode) Then

                If IsNumeric(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) Then

                    GK_ZEN_ZAIKO_QTY = GK_ZEN_ZAIKO_QTY + CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode))


                End If


                If IsNumeric(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) Then


'>>>>>>>>>>>>>>>>   2018.01.22
                    If Trim(StrConv(P_STOCK_REC.CODE, vbUnicode)) <> "" Then        '2018.01.22


                    GK_ZAIKO_QTY = GK_ZAIKO_QTY + CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))
                    End If                                                          '2018.01.22
'>>>>>>>>>>>>>>>>   2018.01.22


                End If

            End If

            com = BtOpGetNext


        Loop



        If GK_ZEN_ZAIKO_QTY = GK_ZAIKO_QTY Then
            Call UniCode_Conv(P_STOCKSUM_REC.ZAIKO_KIN, StrConv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, vbUnicode))





            sts = BTRV(BtOpUpdate, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

            Select Case sts
                Case BtNoErr




                Case Else
                    Call File_Error(sts, BtOpUpdate, "���ޒI���W�v�ް�")
                    Exit Function
            End Select


        End If



        com = BtOpGetNext



    Loop





    '-------------------------------------  ���v�ް��̍ďW�v
    GK_ZEN_ZAIKO_KIN = 0
    GK_NYUKO_KIN = 0
    GK_SYUKO_KIN = 0
    GK_ZAIKO_KIN = 0



    com = BtOpGetFirst

    Do
        DoEvents

        sts = BTRV(com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

        Select Case sts
            Case BtNoErr


            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "���ޒI���W�v�ް�")
                Exit Function
        End Select


        wk_Val = Val(StrConv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, vbUnicode)) + Val(StrConv(P_STOCKSUM_REC.NYUKO_KIN, vbUnicode)) - (Val(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode)))

        If wk_Val < 0 Then
            Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, Format(wk_Val, "0000000"))
        Else
            Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, Format(wk_Val, "00000000"))
        End If


        'Clng --> Val 2016.01.08
        GK_ZEN_ZAIKO_KIN = GK_ZEN_ZAIKO_KIN + Val(StrConv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, vbUnicode))
        'Clng --> Val 2016.01.08
        GK_NYUKO_KIN = GK_NYUKO_KIN + Val(StrConv(P_STOCKSUM_REC.NYUKO_KIN, vbUnicode))
        'Clng --> Val 2016.01.08
        GK_SYUKO_KIN = GK_SYUKO_KIN + Val(StrConv(P_STOCKSUM_REC.SYUKO_KIN, vbUnicode))
        'Clng --> Val 2016.01.08
        GK_ZAIKO_KIN = GK_ZAIKO_KIN + Val(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode))


        Do

'            sts = BTRV(BtOpUpdate, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)            '2018.10.26
            sts = BTRV(BtOpUpdate, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)       '2018.01.26

            Select Case sts
                Case BtNoErr
                    Exit Do

                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    DoEvents


                Case Else
                    Call File_Error(sts, BtOpUpdate, "���ޒI���W�v�ް�")
                    Exit Function
            End Select

        Loop




        com = BtOpGetNext

    Loop

    '���vں��ޏo��
    Call UniCode_Conv(K0_P_STOCKSUM.G_SYUSHI, P_StokSum_Key)
    sts = BTRV(BtOpGetEqual, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)

    Select Case sts
        Case BtNoErr
            Upd_Com = BtOpUpdate

        Case BtErrKeyNotFound

            Upd_Com = BtOpInsert


        Case Else
            Call File_Error(sts, BtOpGetEqual, "���ޒI���W�v�ް�")
            Exit Function
    End Select


    If Upd_Com = BtOpInsert Then
        Call UniCode_Conv(P_STOCKSUM_REC.G_SYUSHI, P_StokSum_Key)
        Call UniCode_Conv(P_STOCKSUM_REC.FILLER, "")
    End If

    Call UniCode_Conv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, Format(GK_ZEN_ZAIKO_KIN, "00000000000"))
    Call UniCode_Conv(P_STOCKSUM_REC.NYUKO_KIN, Format(GK_NYUKO_KIN, "00000000000"))
    Call UniCode_Conv(P_STOCKSUM_REC.SYUKO_KIN, Format(GK_SYUKO_KIN, "0000000000"))
    Call UniCode_Conv(P_STOCKSUM_REC.ZAIKO_KIN, Format(GK_ZAIKO_KIN, "00000000000"))

If Upd_Com = BtOpInsert Then
    Call LOG_OUT(LOG_F, "1000=" & StrConv(P_STOCKSUM_REC.G_SYUSHI, vbUnicode))
End If


    Do

'        sts = BTRV(Upd_Com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)           '2018.10.26
        sts = BTRV(Upd_Com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)      '2018.01.26

        Select Case sts
            Case BtNoErr, BtErrDuplicates
                Exit Do

            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                DoEvents


            Case Else
                Call File_Error(sts, Upd_Com, "���ޒI���W�v�ް�")
                Exit Function
        End Select

    Loop



End Function

Private Function Gyo_Update_Proc(GYO As Integer) As Integer

Dim sts As Integer

    Gyo_Update_Proc = True


    Set TDBGrid1(1).Array = STOCK
    TDBGrid1(1).Refresh

    TDBGrid1(1).Update



    Call UniCode_Conv(K0_P_STOCK.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_P_STOCK.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_P_STOCK.HIN_GAI, STOCK(GYO, colHIN_GAI))
    Call UniCode_Conv(K0_P_STOCK.CODE, STOCK(GYO, colSHI_CODE))

    If IsNumeric(STOCK(GYO, colSHI_TANKA)) Then
        Call UniCode_Conv(K0_P_STOCK.TANKA, Format(CDbl(STOCK(GYO, colSHI_TANKA)), "00000000.00"))
    Else
        Call UniCode_Conv(K0_P_STOCK.TANKA, STOCK(GYO, colSHI_TANKA))

    End If



    sts = BTRV(BtOpGetEqual, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

    Select Case sts
        Case BtNoErr



            Call UniCode_Conv(P_STOCK_REC.MOTO_ZAIKO_QTY, Format(CDbl(STOCK(GYO, colMOTO_ZAIKO_QTY)), "00000000"))


            sts = BTRV(BtOpUpdate, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

            Select Case sts
                Case BtNoErr

                Case BtErrKeyNotFound

                Case Else
                    Call File_Error(BtOpGetEqual, BtOpGetEqual, "���ޒI���ް�")
                    Exit Function
            End Select




        Case BtErrKeyNotFound

        Case Else
            Call File_Error(BtOpGetEqual, BtOpGetEqual, "���ޒI���ް�")
            Exit Function
    End Select



    Gyo_Update_Proc = False


End Function

Private Sub STANA_ErrLogPut(LogData As String)
'*************************************************************************
'*�@�@�@���ޒI���G���[���O�@�o�͏���                2010.10.28
'*
'*�@���@���FLogData  : �o�͓��e
'*
'*�@�߂�l�F�Ȃ�
'*************************************************************************
Dim stream      As Integer              '�t�@�C���ԍ�

On Error Resume Next

    If STANA_LOG_F = "" Then    '���ޒI���װ۸�̧�ٖ��̖��� �� ۸ޏo�͖���
        Exit Sub
    End If

    stream = FreeFile
    Open STANA_LOG_F For Append As stream
    Print #stream, LogData
    Close stream

End Sub
Private Function Hin_Sum_Update_Proc() As Integer
'-------------------------------------  �i�ڃ}�X�^���O���c�L�蕪���W�v
Dim sts                     As Integer
Dim com                     As Integer

Dim Sum_Nyuko_Qty           As Long
Dim Sum_Zaiko_Qty           As Long
Dim Sum_Syuko_Qty           As Long

    com = BtOpGetFirst


    Do

        DoEvents
        sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

        Select Case sts
            Case BtNoErr

            Case BtErrEOF
                Exit Do

            Case Else
                Call File_Error(sts, com, "�I���ް�")
                Exit Function

        End Select

Label3(1).Caption = Trim(StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))




        If Trim(StrConv(P_STOCK_REC.CODE, vbUnicode)) <> "" Or _
            Trim(StrConv(P_STOCK_REC.TANKA, vbUnicode)) <> "" Then
        Else

                

            Sum_Nyuko_Qty = 0
            Sum_Zaiko_Qty = 0


            Call UniCode_Conv(K0_wkP_STOCK.JGYOBU, StrConv(P_STOCK_REC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_wkP_STOCK.NAIGAI, StrConv(P_STOCK_REC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_wkP_STOCK.HIN_GAI, StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))

            Call UniCode_Conv(K0_wkP_STOCK.CODE, "")
            Call UniCode_Conv(K0_wkP_STOCK.TANKA, "")

            com = BtOpGetGreater

            
            Do
                
                DoEvents
                
                sts = BTRV(com, wkP_STOCK_POS, wkP_STOCK_REC, Len(wkP_STOCK_REC), K0_wkP_STOCK, Len(K0_wkP_STOCK), 0)
        
                Select Case sts
                    Case BtNoErr
        
                    Case BtErrEOF
                        Exit Do
        
                    Case Else
                        Call File_Error(sts, com, "�I���ް�")
                        Exit Function
        
                End Select

                If Trim(StrConv(wkP_STOCK_REC.HIN_GAI, vbUnicode)) <> Trim(StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)) Then
                    Exit Do
                End If
                
                Sum_Nyuko_Qty = Sum_Nyuko_Qty + Val(StrConv(wkP_STOCK_REC.NYUKO_QTY, vbUnicode))
                Sum_Zaiko_Qty = Sum_Zaiko_Qty + Val(StrConv(wkP_STOCK_REC.ZAIKO_QTY, vbUnicode))
                

                com = BtOpGetNext



            Loop

        
            Call UniCode_Conv(P_STOCK_REC.NYUKO_QTY, Format(Sum_Nyuko_Qty, "0000000"))
            Call UniCode_Conv(P_STOCK_REC.ZAIKO_QTY, Format(Sum_Zaiko_Qty, "0000000"))

        
            Sum_Syuko_Qty = Val(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)) + Sum_Nyuko_Qty - Sum_Zaiko_Qty
        
            Call UniCode_Conv(P_STOCK_REC.SYUKO_QTY, Format(Sum_Syuko_Qty, "0000000"))
        
        
            sts = BTRV(BtOpUpdate, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
    
            Select Case sts
                Case BtNoErr
    
                Case Else
                    Call File_Error(sts, com, "�I���ް�")
                    Exit Function
    
            End Select
        
        
        
        End If
        
        com = BtOpGetNext
        
    Loop

End Function

Private Function Syuka_Update_Proc() As Integer
'-------------------------------------  �i�ڃ}�X�^���O���c�L�蕪���W�v
Dim sts                     As Integer
Dim com                     As Integer



Dim Sum_Nyuko_Qty           As Long
Dim Sum_Zaiko_Qty           As Long
Dim Sum_Syuko_Qty           As Long


Dim wkZaiko_QTY             As Long
'2011.03.28
Dim FSW                     As Boolean

    com = BtOpGetFirst


    Do

        DoEvents
        sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

        Select Case sts
            Case BtNoErr

            Case BtErrEOF
                Exit Do

            Case Else
                Call File_Error(sts, com, "�I���ް�")
                Exit Function

        End Select

Label3(1).Caption = Trim(StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))


        If Trim(StrConv(P_STOCK_REC.CODE, vbUnicode)) <> "" Or _
            Trim(StrConv(P_STOCK_REC.TANKA, vbUnicode)) <> "" Then
        Else

            FSW = True

            Sum_Syuko_Qty = Val(StrConv(P_STOCK_REC.SYUKO_QTY, vbUnicode))


            Call UniCode_Conv(K1_wkP_STOCK.JGYOBU, StrConv(P_STOCK_REC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K1_wkP_STOCK.NAIGAI, StrConv(P_STOCK_REC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K1_wkP_STOCK.HIN_GAI, StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))
            
            Call UniCode_Conv(K1_wkP_STOCK.INPUT_DATE, "")


            Call UniCode_Conv(K1_wkP_STOCK.CODE, "")
            Call UniCode_Conv(K1_wkP_STOCK.TANKA, "")

            com = BtOpGetGreater

            
            Do
                
                DoEvents
                
                sts = BTRV(com, wkP_STOCK_POS, wkP_STOCK_REC, Len(wkP_STOCK_REC), K1_wkP_STOCK, Len(K1_wkP_STOCK), 1)
                
                
                Select Case sts
                    Case BtNoErr
        
                    Case BtErrEOF
                        Exit Do
        
                    Case Else
                        Call File_Error(sts, com, "�I���ް�")
                        Exit Function
        
                End Select

                If Trim(StrConv(wkP_STOCK_REC.HIN_GAI, vbUnicode)) <> Trim(StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)) Then
                    
                    
                    If FSW And Sum_Syuko_Qty <> 0 Then
                        FSW = False
                    
                    
                    
                    
                    
                        Call UniCode_Conv(wkP_STOCK_REC.JGYOBU, StrConv(P_STOCK_REC.JGYOBU, vbUnicode))
                        Call UniCode_Conv(wkP_STOCK_REC.NAIGAI, StrConv(P_STOCK_REC.NAIGAI, vbUnicode))
                        Call UniCode_Conv(wkP_STOCK_REC.HIN_GAI, StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))
                    
                    
                    
                        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_STOCK_REC.JGYOBU, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_STOCK_REC.NAIGAI, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))
        
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
        
        
                                If Not IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)) Then    '2008.02.13
                                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA, "00000000000")
                                End If
                                
                                Call UniCode_Conv(wkP_STOCK_REC.TANKA, StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))
                                
                                Call UniCode_Conv(wkP_STOCK_REC.CODE, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
                        
        
        
        
        
        
        
        
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                                Exit Function
                        End Select
                    
                    
                    
                    
                    
                    
                    
                        Call UniCode_Conv(wkP_STOCK_REC.INPUT_DATE, StrConv(ITEMREC.LAST_NYU_DT, vbUnicode))
        
        
                        Call UniCode_Conv(wkP_STOCK_REC.G_SYUSHI, StrConv(ITEMREC.G_SYUSHI, vbUnicode))
                        Call UniCode_Conv(wkP_STOCK_REC.ZEN_ZAIKO_QTY, "00000000")
        
        
        
                        Call UniCode_Conv(wkP_STOCK_REC.NYUKO_QTY, "00000000")
                        Call UniCode_Conv(wkP_STOCK_REC.SYUKO_QTY, Format(Sum_Syuko_Qty, "0000000"))
                        Call UniCode_Conv(wkP_STOCK_REC.ZAIKO_QTY, "00000000")
        
        
                        Call UniCode_Conv(wkP_STOCK_REC.LAST_SYUKA_DT, StrConv(ITEMREC.LAST_SYU_DT, vbUnicode))
                        Call UniCode_Conv(wkP_STOCK_REC.LAST_SYUKA_QTY, StrConv(ITEMREC.G_LAST_SYUKA_QTY, vbUnicode))
        
                        Call UniCode_Conv(wkP_STOCK_REC.MOTO_ZAIKO_QTY, "00000000")
        
                        Call UniCode_Conv(wkP_STOCK_REC.MAEGARI_QTY, "00000000")
        
                        Call UniCode_Conv(wkP_STOCK_REC.SYUKA_NON_F, "0")     '2008.06.21
        
        
        
                        Call UniCode_Conv(wkP_STOCK_REC.ZEN_ZAIKO_KIN, "00000000")
        
                        Call UniCode_Conv(wkP_STOCK_REC.FILLER, "")
                    
                    
                    
                    
                        sts = BTRV(BtOpInsert, wkP_STOCK_POS, wkP_STOCK_REC, Len(wkP_STOCK_REC), K1_wkP_STOCK, Len(K1_wkP_STOCK), 1)
                    
                    
                        Select Case sts
                            Case BtNoErr, BtErrDuplicates
                
                            Case Else
                                Call File_Error(sts, com, "�I���ް�")
                                Exit Function
                
                        End Select
                    
                    
                    End If
                    
                    
                    
                    
                    Exit Do
                End If
                
                
    '----------------------------------------   ���o�ɐ����O�@�������Ȃ��@----------------------------------------
                If Sum_Syuko_Qty = 0 Then
                    Exit Do
                End If


    '----------------------------------------   ���o�ɐ����O�@��s�ڂɑS���Z�b�g ---------------------------------
                
                If Sum_Syuko_Qty > 0 Then
                
                    FSW = False
                    Call UniCode_Conv(wkP_STOCK_REC.SYUKO_QTY, Format(Sum_Syuko_Qty, "0000000"))
    
    
                    
                    sts = BTRV(BtOpUpdate, wkP_STOCK_POS, wkP_STOCK_REC, Len(wkP_STOCK_REC), K1_wkP_STOCK, Len(K1_wkP_STOCK), 1)
            
                    Select Case sts
                        Case BtNoErr
            
                        Case Else
                            Call File_Error(sts, com, "�I���ް�")
                            Exit Function
            
                    End Select
                
                
                    Exit Do
                
                End If

    '----------------------------------------   ���o�ɐ����O�@----------------------------------------------------



                If Val(StrConv(wkP_STOCK_REC.ZAIKO_QTY, vbUnicode)) >= Abs(Sum_Syuko_Qty) Then
                    


                    FSW = False

                    Call UniCode_Conv(wkP_STOCK_REC.SYUKO_QTY, Format(Sum_Syuko_Qty, "0000000"))
    
    
                    
                    sts = BTRV(BtOpUpdate, wkP_STOCK_POS, wkP_STOCK_REC, Len(wkP_STOCK_REC), K1_wkP_STOCK, Len(K1_wkP_STOCK), 1)
            
                    Select Case sts
                        Case BtNoErr
            
                        Case Else
                            Call File_Error(sts, com, "�I���ް�")
                            Exit Function
            
                    End Select
                
                
                
                
                    Exit Do

                End If

                
                
'''''''''                Call UniCode_Conv(wkP_STOCK_REC.SYUKO_QTY, Format(Val(StrConv(wkP_STOCK_REC.ZAIKO_QTY, vbUnicode)) * -1, "0000000"))
                
                wkZaiko_QTY = Val(StrConv(wkP_STOCK_REC.ZAIKO_QTY, vbUnicode)) * -1
'''''''''                Sum_Syuko_Qty = Sum_Syuko_Qty + Val(StrConv(wkP_STOCK_REC.ZAIKO_QTY, vbUnicode))
                Sum_Syuko_Qty = Sum_Syuko_Qty + wkZaiko_QTY
                    
                
                
                
                
                If Sum_Syuko_Qty = 0 Then
                    
                    
                    sts = BTRV(BtOpUpdate, wkP_STOCK_POS, wkP_STOCK_REC, Len(wkP_STOCK_REC), K1_wkP_STOCK, Len(K1_wkP_STOCK), 1)
            
                    Select Case sts
                        Case BtNoErr
            
                        Case Else
                            Call File_Error(sts, com, "�I���ް�")
                            Exit Function
            
                    End Select
                    
                    
                    Exit Do
                End If
                
                
                
                Call UniCode_Conv(wkP_STOCK_REC.SYUKO_QTY, Format(Val(StrConv(wkP_STOCK_REC.SYUKO_QTY, vbUnicode)) + Val(StrConv(wkP_STOCK_REC.NYUKO_QTY, vbUnicode)), "0000000"))
                
                Sum_Syuko_Qty = Sum_Syuko_Qty - Val(StrConv(wkP_STOCK_REC.NYUKO_QTY, vbUnicode))
                
                
                sts = BTRV(BtOpUpdate, wkP_STOCK_POS, wkP_STOCK_REC, Len(wkP_STOCK_REC), K1_wkP_STOCK, Len(K1_wkP_STOCK), 1)
        
                Select Case sts
                    Case BtNoErr
        
                    Case Else
                        Call File_Error(sts, com, "�I���ް�")
                        Exit Function
        
                End Select
                
                
                
                
                
                com = BtOpGetNext



            Loop

        
        
        
        
        End If
        
        com = BtOpGetNext
        
    Loop





End Function


Private Function ZAIKO_CHK_PROC(ZAIKO_CHK_F As Boolean) As Integer
'*************************************************************************
'   �I���J�n�O�`�F�b�N
'       2015.03.05

'*************************************************************************

Dim com         As Integer
Dim sts         As Integer


Dim DATA_CNT    As Long

Dim stream  As Integer

    On Error Resume Next
    Kill ZAIKO_FILE
    
    Open ZAIKO_FILE For Output As stream

    Close #stream

End Function

Private Function MULTI_TANKA_CHECK_PROC(yn As Integer) As Integer
'*************************************************************************
'   �}���`�P������
'       2016.01.07

'*************************************************************************
Dim sts         As Integer
Dim com         As Integer

Dim wKEIJYO_YM  As String * 6
Dim SKIP_FLG    As Boolean

Dim SYUSHI_ON   As Boolean

Dim i           As Long
Dim j           As Long

Dim Upd_Com     As Integer

Dim wk_Val      As Long

Dim FSW         As Boolean
Dim G_FSW       As Boolean  '2016.02.02


Dim mess        As String



Label3(0).Caption = "�}���`�P��������(�����d����)"
    
    
    MULTI_TANKA_CHECK_PROC = True
    
    G_FSW = False
    
    Erase MULTI_TANKA_TBL
    FSW = True
    
    wKEIJYO_YM = Left(Text1(ptxKEIJYO_YM).Text, 4) & Right(Text1(ptxKEIJYO_YM).Text, 2)

    Call UniCode_Conv(K2_P_SHUKEIRE.KEIJYO_YM, wKEIJYO_YM)
    Call UniCode_Conv(K2_P_SHUKEIRE.UKEIRE_DT, "")
    com = BtOpGetGreaterEqual

    Do

        DoEvents

        sts = BTRV(com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K2_P_SHUKEIRE, Len(K2_P_SHUKEIRE), 2)

        Select Case sts
            Case BtNoErr

                If StrConv(P_SHUKEIRE_REC.KEIJYO_YM, vbUnicode) <> wKEIJYO_YM Then
                    Exit Do
                End If

            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "���ގ���ް�")
                Exit Function
        End Select

        '�����ް��ǂݍ���
        Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
        SKIP_FLG = False
        Select Case sts
            Case BtNoErr
                '�i�ڂ̍݌Ɍv���׸ނ��`�F�b�N
                Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))

                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                SKIP_FLG = False
                Select Case sts
                    Case BtNoErr

                        If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_OFF Then
                            SKIP_FLG = True       '�l�����Ȃ��̂Ž����
                        End If

                        SYUSHI_ON = False               '2007.11.13
                        If GLB_SYUSHI_F = "" Then       '2007.11.13
                            SYUSHI_ON = True
                        Else
                            SYUSHI_ON = False

                            For i = 0 To UBound(G_SYUSHI_TBL)

                                If Trim(StrConv(ITEMREC.G_SYUSHI, vbUnicode)) = G_SYUSHI_TBL(i) Then
                                    SYUSHI_ON = True
                                    Exit For
                                End If


                            Next i
                        End If



                    Case BtErrKeyNotFound


                        SKIP_FLG = True       '�l�����Ȃ��̂Ž����


                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                        Exit Function
                End Select





            Case BtErrKeyNotFound


                SKIP_FLG = True       '�����Ȃ��͒ʏ��ް��ł͂Ȃ�


            Case Else
                Call File_Error(sts, BtOpGetEqual, "���ޒ����ް�")
                Exit Function
        End Select

Label3(1).Caption = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))


        If StrConv(P_SHORDER_REC.G_SHIIRE_KBN, vbUnicode) = ZEI_SHIIRE_KBN Then
            SKIP_FLG = True       '����Ŏd��
        End If

        If Not SKIP_FLG And SYUSHI_ON Then      '2007.11.13


            If FSW Then
                ReDim MULTI_TANKA_TBL(0 To 0)
                MULTI_TANKA_TBL(0).HINBAN = StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)
                
                ReDim Preserve MULTI_TANKA_TBL(0).TANKA(0 To 0)
                MULTI_TANKA_TBL(0).TANKA(0) = Val(StrConv(P_SHUKEIRE_REC.UKEIRE_TANKA, vbUnicode))
                FSW = False
                G_FSW = True        '2016.02.02
            Else
                For i = 0 To UBound(MULTI_TANKA_TBL)
                    If MULTI_TANKA_TBL(i).HINBAN = StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode) Then
                        Exit For
                    End If
                Next i
                    
                If i > UBound(MULTI_TANKA_TBL) Then
                    ReDim Preserve MULTI_TANKA_TBL(0 To i)
                    MULTI_TANKA_TBL(i).HINBAN = StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)
                    ReDim Preserve MULTI_TANKA_TBL(i).TANKA(0 To 0)
                    MULTI_TANKA_TBL(i).TANKA(0) = Val(StrConv(P_SHUKEIRE_REC.UKEIRE_TANKA, vbUnicode))
                Else
                    For j = 0 To UBound(MULTI_TANKA_TBL(i).TANKA)
                        If MULTI_TANKA_TBL(i).TANKA(j) = Val(StrConv(P_SHUKEIRE_REC.UKEIRE_TANKA, vbUnicode)) Then
                            Exit For
                        End If
                    Next j
                    If j > UBound(MULTI_TANKA_TBL(i).TANKA) Then
                        ReDim Preserve MULTI_TANKA_TBL(i).TANKA(0 To j)
                        MULTI_TANKA_TBL(i).TANKA(j) = Val(StrConv(P_SHUKEIRE_REC.UKEIRE_TANKA, vbUnicode))
                    Else
                        MULTI_TANKA_TBL(i).TANKA(j) = Val(StrConv(P_SHUKEIRE_REC.UKEIRE_TANKA, vbUnicode))
                    End If
                End If
            End If

        End If


        com = BtOpGetNext

    Loop
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
Label3(0).Caption = "�}���`�P��������(�݌�)"

    '-------------------------------------  ���݂�蓖���c�݌ɂ��W�v
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, SHIZAI)
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, "")
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, "")
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, "")
    Call UniCode_Conv(K1_ZAIKO.SOKO_NO, "")
    Call UniCode_Conv(K1_ZAIKO.Retu, "")
    Call UniCode_Conv(K1_ZAIKO.Ren, "")
    Call UniCode_Conv(K1_ZAIKO.Dan, "")

    com = BtOpGetGreaterEqual

    Do

        DoEvents

        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)

        Select Case sts
            Case BtNoErr

                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> SHIZAI Then
                    Exit Do
                End If

            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "�݌��ް�")
                Exit Function
        End Select

        SKIP_FLG = False
        
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(ZAIKOREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(ZAIKOREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ZAIKOREC.HIN_GAI, vbUnicode))


        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

        Select Case sts
            Case BtNoErr

                If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_OFF Or _
                    StrConv(ITEMREC.ZAIKO_CLR_F, vbUnicode) = "1" Then                          '2012.12.13
                    SKIP_FLG = True

                Else
                    If Not IsNumeric(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)) Then    '2008.02.13
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(0).TANKA, "00000000000")
                    End If

                    If Trim(StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode)) = "" Then
                        Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode))
                        '2008.11.24
                        Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, Format(CDbl(Trim(StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode))), "00000000.00"))
                    End If

                End If

Label3(1).Caption = Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode))

                SYUSHI_ON = False               '2007.11.13
                If GLB_SYUSHI_F = "" Then       '2007.11.13
                    SYUSHI_ON = True
                Else
                    SYUSHI_ON = False

                    For i = 0 To UBound(G_SYUSHI_TBL)

                        If Trim(StrConv(ITEMREC.G_SYUSHI, vbUnicode)) = G_SYUSHI_TBL(i) Then
                            SYUSHI_ON = True
                            Exit For
                        End If


                    Next i
                End If

            Case BtErrKeyNotFound
                SKIP_FLG = True

            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                Exit Function
        End Select


        If Not SKIP_FLG And SYUSHI_ON Then      '2007.11.13


            If FSW Then
                ReDim MULTI_TANKA_TBL(0 To 0)
                MULTI_TANKA_TBL(0).HINBAN = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
                ReDim Preserve MULTI_TANKA_TBL(0).TANKA(0 To 0)
                MULTI_TANKA_TBL(0).TANKA(0) = Val(StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode))
                FSW = False
                G_FSW = True        '2016.02.02
            
            Else
                For i = 0 To UBound(MULTI_TANKA_TBL)
                    If MULTI_TANKA_TBL(i).HINBAN = StrConv(ZAIKOREC.HIN_GAI, vbUnicode) Then
                        Exit For
                    End If
                Next i
                    
                If i > UBound(MULTI_TANKA_TBL) Then
                    ReDim Preserve MULTI_TANKA_TBL(0 To i)
                    MULTI_TANKA_TBL(i).HINBAN = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
                    ReDim Preserve MULTI_TANKA_TBL(i).TANKA(0 To 0)
                    MULTI_TANKA_TBL(i).TANKA(0) = Val(StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode))
                Else
                    For j = 0 To UBound(MULTI_TANKA_TBL(i).TANKA)
                        If MULTI_TANKA_TBL(i).TANKA(j) = Val(StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode)) Then
                            Exit For
                        End If
                    Next j
                    If j > UBound(MULTI_TANKA_TBL(i).TANKA) Then
                        ReDim Preserve MULTI_TANKA_TBL(i).TANKA(0 To j)
                        MULTI_TANKA_TBL(i).TANKA(j) = Val(StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode))
                    Else
                        MULTI_TANKA_TBL(i).TANKA(j) = Val(StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode))
                    End If
                End If
            End If



        End If

        com = BtOpGetNext

    Loop
    
    
Label3(0).Caption = "�}���`�P�������I��"
     yn = vbYes
    
    FSW = False
    
    If G_FSW Then
        For i = 0 To UBound(MULTI_TANKA_TBL)
        
            If UBound(MULTI_TANKA_TBL(i).TANKA) > 0 Then
                FSW = True
                Exit For
            End If
        Next i
    End If
    
    
    If FSW Then
        mess = ""
        mess = "�����P�������i�Ԃ��L��܂��B�ȉ��̕i�Ԃ��m�F���ĉ������B" & Chr(13) & Chr(10)
    
    
        For i = 0 To UBound(MULTI_TANKA_TBL)
                    
            If UBound(MULTI_TANKA_TBL(i).TANKA) > 0 Then
                mess = mess & MULTI_TANKA_TBL(i).HINBAN
                For j = 0 To UBound(MULTI_TANKA_TBL(i).TANKA)
                
                    mess = mess & " " & Format(MULTI_TANKA_TBL(i).TANKA(j), "0.00")
                
                Next j
                mess = mess & Chr(13) & Chr(10)
        
        
        
            End If
        
        Next i
        mess = mess & Chr(13) & Chr(10)
        mess = mess & "�I���J�n�������p�����܂����H"
    
    
        yn = MsgBox(mess, vbYesNo, "�m�F����")
    
    End If
    
    MULTI_TANKA_CHECK_PROC = False

End Function

Private Function SAKI_SHIIRE_Proc(ROW As Long, wkSAKI_SHIIRE As Long) As Integer
'*************************************************************************
'   ����t���d������
'       2017.04.22

'*************************************************************************
Dim wKEIJYO_YM              As String * 6

Dim com                     As Integer
Dim Upd_Com                 As Integer
Dim sts                     As Integer

Dim SKIP_FLG                As Boolean

Dim SYUSHI_ON               As Boolean

Dim i                       As Integer

Dim wk_Val                  As Double



    wkSAKI_SHIIRE = 0

    wKEIJYO_YM = Left(Text1(ptxKEIJYO_YM).Text, 4) & Right(Text1(ptxKEIJYO_YM).Text, 2)

    Call UniCode_Conv(K2_P_SHUKEIRE.KEIJYO_YM, wKEIJYO_YM)
    Call UniCode_Conv(K2_P_SHUKEIRE.UKEIRE_DT, "99999999")

    com = BtOpGetGreaterEqual


    Do

        DoEvents

        sts = BTRV(com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K2_P_SHUKEIRE, Len(K2_P_SHUKEIRE), 2)

        Select Case sts
            Case BtNoErr


            Case BtErrEOF

                Exit Do


            Case Else
                Call File_Error(sts, com, "���ގ���ް�")
                Exit Function
        End Select



        '�����ް��ǂݍ���
        Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
        SKIP_FLG = False
        Select Case sts
            Case BtNoErr
                
                
                If Trim(STOCK(ROW, colHIN_GAI)) = Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)) Then
                
                
                
                
                    '�i�ڂ̍݌Ɍv���׸ނ��`�F�b�N
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    SKIP_FLG = False
                    Select Case sts
                        Case BtNoErr
    
                            If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_OFF Then
                                SKIP_FLG = True       '�l�����Ȃ��̂Ž����
                            End If
    
                            SYUSHI_ON = False               '2007.11.13
                            If GLB_SYUSHI_F = "" Then       '2007.11.13
                                SYUSHI_ON = True
                            Else
                                SYUSHI_ON = False
    
                                For i = 0 To UBound(G_SYUSHI_TBL)
    
                                    If Trim(StrConv(ITEMREC.G_SYUSHI, vbUnicode)) = G_SYUSHI_TBL(i) Then
                                        SYUSHI_ON = True
                                        Exit For
                                    End If
    
    
                                Next i
                            End If
    
    
    
                        Case BtErrKeyNotFound
    
    
                            SKIP_FLG = True       '�l�����Ȃ��̂Ž����
    
    
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                            Exit Function
                    End Select
    

                End If


            Case BtErrKeyNotFound


                SKIP_FLG = True       '�����Ȃ��͒ʏ��ް��ł͂Ȃ�


            Case Else
                Call File_Error(sts, BtOpGetEqual, "���ޒ����ް�")
                Exit Function
        End Select

Label3(1).Caption = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))


        If StrConv(P_SHORDER_REC.G_SHIIRE_KBN, vbUnicode) = ZEI_SHIIRE_KBN Then
            SKIP_FLG = True       '����Ŏd��
        End If

        If Not SKIP_FLG And SYUSHI_ON Then


            If CDbl(STOCK(ROW, colSHI_TANKA)) = CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_TANKA, vbUnicode)) And _
                Trim(STOCK(ROW, colSHI_CODE)) = Trim(StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode)) Then

                wkSAKI_SHIIRE = wkSAKI_SHIIRE + CLng(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode))

            End If
        End If

        com = BtOpGetNext

    Loop


End Function
