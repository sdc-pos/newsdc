VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PR000601 
   Caption         =   "���Y���і��׏����s"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15150
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
   ScaleWidth      =   15150
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox txSEL_KEY 
      Height          =   375
      Left            =   10680
      TabIndex        =   26
      Text            =   "Text2"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "�o�͑Ώ�"
      Height          =   855
      Left            =   6960
      TabIndex        =   25
      Top             =   480
      Width           =   3015
      Begin VB.CheckBox Check1 
         Caption         =   "���ו\"
         Height          =   375
         Index           =   3
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�W�v�\"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   6
      Top             =   1200
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   1
      Left            =   2520
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   7
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Caption         =   "���E"
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�O��"
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   2520
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   4
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   7455
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   2040
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   13150
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "CODE"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "��z��"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
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
      Columns(12).Caption=   "���v"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "�����"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "�x�����v"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   15
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=15"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1561"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1455"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3201"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3096"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2381"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2275"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2381"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2275"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2381"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2275"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2381"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2275"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=2381"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2275"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=2381"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2275"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=2381"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2275"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(9).Width=2381"
      Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2275"
      Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=2"
      Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(51)=   "Column(10).Width=2381"
      Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=2275"
      Splits(0)._ColumnProps(54)=   "Column(10)._ColStyle=2"
      Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(56)=   "Column(11).Width=2381"
      Splits(0)._ColumnProps(57)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(11)._WidthInPix=2275"
      Splits(0)._ColumnProps(59)=   "Column(11)._ColStyle=2"
      Splits(0)._ColumnProps(60)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(61)=   "Column(12).Width=2381"
      Splits(0)._ColumnProps(62)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(12)._WidthInPix=2275"
      Splits(0)._ColumnProps(64)=   "Column(12)._ColStyle=2"
      Splits(0)._ColumnProps(65)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(66)=   "Column(13).Width=2381"
      Splits(0)._ColumnProps(67)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(13)._WidthInPix=2275"
      Splits(0)._ColumnProps(69)=   "Column(13)._ColStyle=2"
      Splits(0)._ColumnProps(70)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(71)=   "Column(14).Width=2381"
      Splits(0)._ColumnProps(72)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(14)._WidthInPix=2275"
      Splits(0)._ColumnProps(74)=   "Column(14)._ColStyle=2"
      Splits(0)._ColumnProps(75)=   "Column(14).Order=15"
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
      Caption         =   "���Y�W�v����"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFF&,.bold=0,.fontsize=975"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=�l�r �S�V�b�N"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=�l�r �S�V�b�N"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=975,.italic=0"
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
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=110,.parent=43,.alignment=0"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=107,.parent=44"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=108,.parent=45"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=109,.parent=47"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=16,.parent=43,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=13,.parent=44"
      _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=14,.parent=45"
      _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=15,.parent=47"
      _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=28,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(53)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(54)  =   ":id=28,.fontname=�l�r �S�V�b�N"
      _StyleDefs(55)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=44"
      _StyleDefs(56)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=45"
      _StyleDefs(57)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=47"
      _StyleDefs(58)  =   "Splits(0).Columns(4).Style:id=66,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(59)  =   ":id=66,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(60)  =   ":id=66,.fontname=�l�r �S�V�b�N"
      _StyleDefs(61)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=44"
      _StyleDefs(62)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=45"
      _StyleDefs(63)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=47"
      _StyleDefs(64)  =   "Splits(0).Columns(5).Style:id=32,.parent=43,.alignment=1,.bold=0,.fontsize=975"
      _StyleDefs(65)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(66)  =   ":id=32,.fontname=�l�r �S�V�b�N"
      _StyleDefs(67)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=44"
      _StyleDefs(68)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=45"
      _StyleDefs(69)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=47"
      _StyleDefs(70)  =   "Splits(0).Columns(6).Style:id=74,.parent=43,.alignment=1"
      _StyleDefs(71)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=44"
      _StyleDefs(72)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=45"
      _StyleDefs(73)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=47"
      _StyleDefs(74)  =   "Splits(0).Columns(7).Style:id=20,.parent=43,.alignment=1"
      _StyleDefs(75)  =   "Splits(0).Columns(7).HeadingStyle:id=17,.parent=44"
      _StyleDefs(76)  =   "Splits(0).Columns(7).FooterStyle:id=18,.parent=45"
      _StyleDefs(77)  =   "Splits(0).Columns(7).EditorStyle:id=19,.parent=47"
      _StyleDefs(78)  =   "Splits(0).Columns(8).Style:id=24,.parent=43,.alignment=1"
      _StyleDefs(79)  =   "Splits(0).Columns(8).HeadingStyle:id=21,.parent=44"
      _StyleDefs(80)  =   "Splits(0).Columns(8).FooterStyle:id=22,.parent=45"
      _StyleDefs(81)  =   "Splits(0).Columns(8).EditorStyle:id=23,.parent=47"
      _StyleDefs(82)  =   "Splits(0).Columns(9).Style:id=78,.parent=43,.alignment=1"
      _StyleDefs(83)  =   "Splits(0).Columns(9).HeadingStyle:id=75,.parent=44"
      _StyleDefs(84)  =   "Splits(0).Columns(9).FooterStyle:id=76,.parent=45"
      _StyleDefs(85)  =   "Splits(0).Columns(9).EditorStyle:id=77,.parent=47"
      _StyleDefs(86)  =   "Splits(0).Columns(10).Style:id=62,.parent=43,.alignment=1"
      _StyleDefs(87)  =   "Splits(0).Columns(10).HeadingStyle:id=59,.parent=44"
      _StyleDefs(88)  =   "Splits(0).Columns(10).FooterStyle:id=60,.parent=45"
      _StyleDefs(89)  =   "Splits(0).Columns(10).EditorStyle:id=61,.parent=47"
      _StyleDefs(90)  =   "Splits(0).Columns(11).Style:id=70,.parent=43,.alignment=1"
      _StyleDefs(91)  =   "Splits(0).Columns(11).HeadingStyle:id=67,.parent=44"
      _StyleDefs(92)  =   "Splits(0).Columns(11).FooterStyle:id=68,.parent=45"
      _StyleDefs(93)  =   "Splits(0).Columns(11).EditorStyle:id=69,.parent=47"
      _StyleDefs(94)  =   "Splits(0).Columns(12).Style:id=82,.parent=43,.alignment=1"
      _StyleDefs(95)  =   "Splits(0).Columns(12).HeadingStyle:id=79,.parent=44"
      _StyleDefs(96)  =   "Splits(0).Columns(12).FooterStyle:id=80,.parent=45"
      _StyleDefs(97)  =   "Splits(0).Columns(12).EditorStyle:id=81,.parent=47"
      _StyleDefs(98)  =   "Splits(0).Columns(13).Style:id=86,.parent=43,.alignment=1"
      _StyleDefs(99)  =   "Splits(0).Columns(13).HeadingStyle:id=83,.parent=44"
      _StyleDefs(100) =   "Splits(0).Columns(13).FooterStyle:id=84,.parent=45"
      _StyleDefs(101) =   "Splits(0).Columns(13).EditorStyle:id=85,.parent=47"
      _StyleDefs(102) =   "Splits(0).Columns(14).Style:id=90,.parent=43,.alignment=1"
      _StyleDefs(103) =   "Splits(0).Columns(14).HeadingStyle:id=87,.parent=44"
      _StyleDefs(104) =   "Splits(0).Columns(14).FooterStyle:id=88,.parent=45"
      _StyleDefs(105) =   "Splits(0).Columns(14).EditorStyle:id=89,.parent=47"
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      Index           =   1
      Left            =   1080
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
      Index           =   0
      Left            =   240
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�`"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   24
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�Ώ۔N����"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   23
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "PR000601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'�e�L�X�g�p�Y��
Private Const ptxS_YMD% = 0                 '�J�n�@�Ώ۔N����
Private Const ptxE_YMD% = 1                 '�I���@�Ώ۔N����

Private Const ptxGENERAL% = 2               '�O��
Private Const ptxNAISYOKU% = 3              '���E



'�R���{�p�Y��
Private Const pcmbGENERAL% = 0              '�O��
Private Const pcmbNAISYOKU% = 1             '���E


'�`�F�b�N�{�b�N�X�p�Y��
Private Const pchkGENERAL% = 0              '�O��
Private Const pchkNAISYOKU% = 1             '���E

Private Const pchkGK% = 2                   '�W�v�\
Private Const pchkDET% = 3                  '���ו\


'Glid�p��---------------------------------

'�d������
Private Const pGridDETAIL% = 0


Private SEISAN      As New XArrayDB


Private Const Min_Row% = 1                  '�ŏ��s��
Private Const Min_Col% = 0                  '�ŏ���
Private Const Max_Col% = 14                 '�ő��

Private Const colTORI_CODE% = 0             '����溰��
Private Const colTORI_NAME% = 1             '����於��
Private Const colSHUMUKE01_KIN% = 2         '�d������1
Private Const colSHUMUKE021_KIN% = 3        '�d������2
Private Const colSHUMUKE03_KIN% = 4         '�d������3
Private Const colSHUMUKE04_KIN% = 5         '�d������4
Private Const colSHUMUKE05_KIN% = 6         '�d������5
Private Const colSHUMUKE06_KIN% = 7         '�d������6
Private Const colSHUMUKE07_KIN% = 8         '�d������7
Private Const colSHUMUKE08_KIN% = 9         '�d������8
Private Const colSHUMUKE09_KIN% = 10        '�d������9
Private Const colSHUMUKE10_KIN% = 11        '�d������10
Private Const colTOTAL% = 12                '���v
Private Const colZEI% = 13                  '����Ŋz
Private Const colSHIHARAI% = 14             '�x�����z




Private Sort_Tbl(Min_Col To Max_Col) _
                As Integer                  '��Ă̐��� 0:���� 1:�~��
Private Tbl_Set_F   As Boolean




Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PR000601.MousePointer = vbHourglass

    Call Ctrl_Lock(PR000601)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PR000601)


    PR000601.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
    
Dim sts     As Integer
Dim com     As Integer
    
Dim i       As Integer
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        
        Case ptxS_YMD           '�Ώ۔N����
        
            
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "0000/01/01"
            End If
            
            If Not IsDate(Text1(Mode).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(Mode).SetFocus
                Exit Function
            Else
                
                Text1(Mode).Text = Format(CDate(Text1(Mode).Text), "YYYY/MM/DD")
            
            End If
        
        Case ptxE_YMD           '�Ώ۔N����
        
            
            If Text1(Mode).Text = "" Then
                Text1(Mode).Text = "9999/12/31"
            End If
            
            If Not IsDate(Text1(Mode).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(Mode).SetFocus
                Exit Function
            Else
                
                Text1(Mode).Text = Format(CDate(Text1(Mode).Text), "YYYY/MM/DD")
            
            End If
        
        
        
        
        Case ptxGENERAL     '�O������
           
           
            Combo1(pcmbGENERAL).ListIndex = -1
            For i = 0 To Combo1(pcmbGENERAL).ListCount - 1
                If Trim(Text1(ptxGENERAL).Text) = Trim(Right(Combo1(pcmbGENERAL).List(i), 5)) Then
                    Combo1(pcmbGENERAL).ListIndex = i
                    Exit For
                End If
            
            Next i
        
        Case ptxNAISYOKU    '���E����
           
           
            Combo1(pcmbNAISYOKU).ListIndex = -1
            For i = 0 To Combo1(pcmbNAISYOKU).ListCount - 1
                If Trim(Text1(ptxNAISYOKU).Text) = Trim(Right(Combo1(pcmbNAISYOKU).List(i), 5)) Then
                    Combo1(pcmbNAISYOKU).ListIndex = i
                    Exit For
                End If
            
            Next i
        
        
        
        
        
    End Select
        
        
    Error_Check_Proc = False
    

End Function


Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
    Select Case Index
        Case pcmbGENERAL        '�O��
        
            Text1(ptxGENERAL).Text = Trim(Right(Combo1(pcmbGENERAL).Text, 5))
        Case pcmbNAISYOKU       '���E
        
            Text1(ptxNAISYOKU).Text = Trim(Right(Combo1(pcmbNAISYOKU).Text, 5))
    End Select
    
    
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    Select Case Index
        Case pcmbGENERAL        '�O��
        
            Text1(ptxGENERAL).Text = Trim(Right(Combo1(pcmbGENERAL).Text, 5))
        Case pcmbNAISYOKU       '���E
        
            Text1(ptxNAISYOKU).Text = Trim(Right(Combo1(pcmbNAISYOKU).Text, 5))
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans         As Integer
Dim i           As Integer

Dim Data_Flg    As Boolean

Dim rpt             As New PR00060F1
Dim f               As New PR000603


Dim sts         As Integer

    Select Case Index
        Case P_CMD_Upd          '�X�V
        
        Case P_CMD_DEL          '�폜
        
        Case P_CMD_DSP                      '����/�\��
        
            For i = ptxS_YMD To ptxNAISYOKU
            
                If Error_Check_Proc(i) Then     '�G���[�`�F�b�N
                    Exit Sub
                End If
            
            Next i
            
            If SUM_Make_Proc(Data_Flg) Then
                Exit Sub
            End If
            
            
            If Not Data_Flg Then
                MsgBox "�Ώ��ް�������܂���"
                Exit Sub
            End If
        
        
            If List_Disp_Proc() Then
                Unload Me
            End If
        
            Text1(ptxS_YMD).SetFocus
        
        
        Case P_CMD_OUT                      '�ް��o��
        
        Case P_CMD_PRT                      '���
 
            For i = ptxS_YMD To ptxNAISYOKU
            
                If Error_Check_Proc(i) Then     '�G���[�`�F�b�N
                    Exit Sub
                End If
            
            Next i
            
            If SUM_Make_Proc(Data_Flg) Then
                Exit Sub
            End If
            
            
            If Not Data_Flg Then
                MsgBox "�Ώ��ް�������܂���"
                Exit Sub
            End If
                
            ans = MsgBox("������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                
            
                If Check1(pchkGK).Value = vbChecked Then
            
                    
                    Set rpt = New PR00060F1
                
                    '���|�[�g��������܂��B�itrue�F����_�C�A���O���� false�F�Ȃ��j
                    rpt.PrintReport False
                
                    Set rpt = Nothing
                    
'                    f.RunReport rpt
'                    f.Show
                
                End If
            
                If Check1(pchkDET).Value = vbChecked Then
                    '���ו\
                    If D_Print_Proc() Then
                        Unload Me
                    End If
                End If
            
            
            End If
            
            Text1(ptxS_YMD).SetFocus
            
            
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
                                '�N���X�}�X�^�n�o�d�m
    If P_Class_Open(BtOpenNomal) Then
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
                                '�Ǘ��}�X�^�n�o�d�m
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���Y���і����ް��n�o�d�m
    If P_SEISAN_DET_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���Y���і��׏W�v�ް��n�o�d�m
    If P_SEISAN_GK_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���i���w��(�e)�ް��n�o�d�m
    If P_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���i���w��(�q)�ް��n�o�d�m
    If P_SSHIJI_K_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���i���w����������ް��n�o�d�m
    If P_SUKEIRE_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    Load PR000602
    Load PR000603
    
    
    
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
    
    '�d������ݒ�
    If SHIMUKE_TBL_Proc(i, P_KBN04_CD) Then
        Unload Me
    End If
            
            
            
            
            
    If i = -1 Then
        MsgBox "�d�����悪�ݒ肳��Ă��܂���"
        Unload Me
    End If
    
    
            '�O���P���ύX�׸ސݒ�   2007.07.13
    If GetIni(App.EXEName, "GAICYU", "P_SYS", c) Then
        GAICYU_F = False
    Else
        If Trim(c) = "1" Then
            GAICYU_F = True
        Else
            GAICYU_F = False
        End If
    End If
    
    
    
    '�O����
    If UKEHARAI_Set_Proc(pcmbGENERAL, P_TORI_GENERAL) Then
        Unload Me
    End If
    '���E
    If UKEHARAI_Set_Proc(pcmbNAISYOKU, P_TORI_NAISYOKU) Then
        Unload Me
    End If
    
            
    '��ʏ����ݒ�
    If Init_Proc() Then
        Unload Me
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
                                            '�N���X�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�N���X�}�X�^")
        End If
    End If
                                            '�󕥐�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�󕥐�}�X�^")
        End If
    End If
    
                                            '�Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�N���X�}�X�^")
        End If
    End If
                                            '���Y���і��׏W�v�ް��b�k�n�r�d
    sts = BTRV(BtOpClose, P_SEISAN_GK_POS, P_SEISAN_GK_REC, Len(P_SEISAN_GK_REC), K0_P_SEISAN_GK, Len(K0_P_SEISAN_GK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���Y���і����ް�")
        End If
    End If
                                            '���Y���і��׃f�[�^CLOSE
    sts = BTRV(BtOpClose, P_SEISAN_DET_POS, P_SEISAN_DET_REC, Len(P_SEISAN_DET_REC), K0_P_SEISAN_DET, Len(K0_P_SEISAN_DET), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���Y���і����ް�")
        End If
    End If
                                            '���i���w���i�e�j�ް��b�k�n�r�d
    sts = BTRV(BtOpClose, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���w���i�e�j�ް�")
        End If
    End If
                                            '���i���w���i�q�j�ް��b�k�n�r�d
    sts = BTRV(BtOpClose, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���w���i�e�j�ް�")
        End If
    End If
                                            '���i����������ް��b�k�n�r�d
    sts = BTRV(BtOpClose, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i����������ް�")
        End If
    End If
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PR000601 = Nothing
    Set PR000602 = Nothing
    Set PR000603 = Nothing


    End
End Sub





Private Sub TDBGrid1_DblClick(Index As Integer)
    
    txSEL_KEY.Text = SEISAN(TDBGrid1(Index).Bookmark, colTORI_CODE)
    If Item_Input_Proc() Then           '���ד���
        Unload Me
    End If

End Sub

Private Sub TDBGrid1_HeadClick(Index As Integer, ByVal ColIndex As Integer)



    Select Case Index
        
        Case pGridDETAIL        '���Y���і���
    
    
            If Sort_Tbl(ColIndex) = 0 Then
                Sort_Tbl(ColIndex) = 1
            Else
                If Sort_Tbl(ColIndex) = 1 Then
                    Sort_Tbl(ColIndex) = 0
                End If
            
            End If
            
            If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then
                            
                SEISAN.QuickSort Min_Row, SEISAN.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
                
                Set TDBGrid1(Index).Array = SEISAN
                
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
    
    
    
    For i = ptxS_YMD To ptxNAISYOKU
        Text1(i).Text = ""
    Next i
    
    '�����N����������
    Text1(ptxS_YMD).Text = Format(Now, "YYYY/MM/DD")
    Text1(ptxE_YMD).Text = Format(Now, "YYYY/MM/DD")
    
    For i = pcmbGENERAL To pcmbNAISYOKU
        
        Combo1(i).ListIndex = -1
    
    Next i


    For i = pchkGENERAL To pchkDET
    
        Check1(i).Value = vbUnchecked
    Next i
    
    
    
    '��ď��̏�����
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0               '��̫�ď���
    Next i

    Init_Proc = False

End Function



Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'           ���ގ���f�[�^�̕\��
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim Row             As Long





    List_Disp_Proc = True
    PR000601.MousePointer = vbHourglass
        
    
    '-------------------------------------  '���і��ׂ̾��
    Set SEISAN = Nothing
    
    Row = Min_Row - 1
    
    
    
    com = BtOpGetFirst
    
    
    Do
    
        DoEvents
    
        sts = BTRV(com, P_SEISAN_GK_POS, P_SEISAN_GK_REC, Len(P_SEISAN_GK_REC), K0_P_SEISAN_GK, Len(K0_P_SEISAN_GK), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���Y���і��׏W�v�ް�")
                Exit Function
        End Select
    
    
        If Trim(StrConv(P_SEISAN_GK_REC.TORI_CODE, vbUnicode)) = "" Then
        Else
            Row = Row + 1
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
        End If
        
        com = BtOpGetNext
    
    
    
    Loop
    
    
    Set TDBGrid1(pGridDETAIL).Array = SEISAN
    TDBGrid1(pGridDETAIL).ReBind
    TDBGrid1(pGridDETAIL).Update
    TDBGrid1(pGridDETAIL).MoveFirst
    
    
    PR000601.MousePointer = vbDefault
    
    
    List_Disp_Proc = False
    


End Function


Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'           ���Y���т̓��e���د�ނɾ�Ă���
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim i           As Integer
Dim j           As Integer

Dim TOTAL       As Long

Dim ZEI         As Long


    Grid_Set_Proc = True
    
    
    SEISAN.ReDim Min_Row, Row, Min_Col, Max_Col


    '����溰��
    SEISAN(Row, colTORI_CODE) = StrConv(P_SEISAN_GK_REC.TORI_CODE, vbUnicode)
    '����於��
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SEISAN_GK_REC.TORI_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        
    Select Case sts
        Case BtNoErr
        
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
            Call UniCode_Conv(P_UKEHARAIREC.TORI_KBN, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�󕥐�Ͻ�")
            Exit Function
    End Select
    SEISAN(Row, colTORI_NAME) = StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode)
        
    j = colSHUMUKE01_KIN
    TOTAL = 0
    For i = 0 To UBound(SHIMUKE_TBL)
        
        '�d�������
        SEISAN(Row, j) = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(i).KIN, vbUnicode)), "#,##0")
        TOTAL = TOTAL + CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(i).KIN, vbUnicode))
        j = j + 1
    
    Next i
    SEISAN(Row, colTOTAL) = Format(TOTAL, "#,##0")
    
    Select Case StrConv(P_SEISAN_GK_REC.TORI_KBN, vbUnicode)
        Case P_TORI_GENERAL
            
            If GAICYU_F Then        '2007.07.17
            
                SEISAN(Row, colZEI) = ""
                
            Else
                
                ZEI = Int(Int(TOTAL * CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode) / 100)) + CInt(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode) / 10))
                SEISAN(Row, colZEI) = Format(ZEI, "#,##0")
            
            End If                  '2007.07.17
            
            
            SEISAN(Row, colSHIHARAI) = Format(TOTAL + ZEI, "#,##0")
        Case Else
            SEISAN(Row, colZEI) = ""
            SEISAN(Row, colSHIHARAI) = Format(TOTAL, "#,##0")
    End Select
    
    
    Grid_Set_Proc = False

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


Private Function UKEHARAI_Set_Proc(Index As Integer, KBN As String) As Integer
'----------------------------------------------------------------------------
'                   �󕥐�}�X�^���R���{�ɃZ�b�g����B
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
    
    UKEHARAI_Set_Proc = True
    
    Combo1(Index).Clear
    
    Combo1(Index).AddItem Space(5)
    
    Call UniCode_Conv(K1_P_UKEHARAI.TORI_KBN, KBN)
    Call UniCode_Conv(K1_P_UKEHARAI.UKEHARAI_CODE, "")

    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K1_P_UKEHARAI, Len(K1_P_UKEHARAI), 1)
            
        Select Case sts
            Case BtNoErr
            
                                
                If StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) <> KBN Then
                    
                    Exit Do
                
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�����}�X�^")
                Exit Function
        
        End Select

        
        
        
        Combo1(Index).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode) & "            " & _
                                StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        
        
        com = BtOpGetNext
    
    Loop

    UKEHARAI_Set_Proc = False
    



End Function


Private Function SUM_Make_Proc(Data_Flg As Boolean) As Integer
'----------------------------------------------------------------------------
'                   ���Y���яW�v�ް��쐬
'----------------------------------------------------------------------------
Dim sts                     As Integer
Dim com                     As Integer

Dim SKIP_Flg                As Boolean
    
    
Dim i                       As Integer
    
Dim SAVE_TORI_KBN           As String * 1
Dim SAVE_TORI_CODE          As String * 5


Dim ALL_KIN(0 To 9)         As Long
Dim ALL_CNT                 As Integer
Dim ALL_QTY                 As Double
Dim KAZEI                   As Long


Dim TOTAL_KIN(0 To 9)       As Long
Dim TOTAL_CNT               As Integer
Dim TOTAL_QTY               As Double
    
Dim wkTANKA                 As Double
    
    
    SUM_Make_Proc = True
    PR000601.MousePointer = vbHourglass

    '-----------------------------------------  �W�v�ް��S���폜


    com = BtOpGetFirst



    Do
    
    
        sts = BTRV(com, P_SEISAN_GK_POS, P_SEISAN_GK_REC, Len(P_SEISAN_GK_REC), K0_P_SEISAN_GK, Len(K0_P_SEISAN_GK), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
            
                Exit Do
            
            
            Case Else
                Call File_Error(sts, com, "���Y���і��׏W�v�ް�")
                Exit Function
        End Select

        sts = BTRV(BtOpDelete, P_SEISAN_GK_POS, P_SEISAN_GK_REC, Len(P_SEISAN_GK_REC), K0_P_SEISAN_GK, Len(K0_P_SEISAN_GK), 0)
            
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "���Y���і��׏W�v�ް�")
        End Select

    
        com = BtOpGetNext
    
    Loop
    
    com = BtOpGetFirst



    Do
    
    
        sts = BTRV(com, P_SEISAN_DET_POS, P_SEISAN_DET_REC, Len(P_SEISAN_DET_REC), K0_P_SEISAN_DET, Len(K0_P_SEISAN_DET), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
            
                Exit Do
            
            
            Case Else
                Call File_Error(sts, com, "���Y���і����ް�")
                Exit Function
        End Select

        sts = BTRV(BtOpDelete, P_SEISAN_DET_POS, P_SEISAN_DET_REC, Len(P_SEISAN_DET_REC), K0_P_SEISAN_DET, Len(K0_P_SEISAN_DET), 0)
            
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "���Y���і��׏W�v�ް�")
        End Select

    
        com = BtOpGetNext
    
    Loop
    
        
    '-----------------------------------------  �W�v�����J�n
    
    Data_Flg = False
        
    
    '----------------   �O��
    If Check1(pchkGENERAL).Value = vbChecked Then
    
        Call UniCode_Conv(K2_P_SUKEIRE.TORI_CODE, Text1(ptxGENERAL).Text)
        
        If Trim(Text1(ptxGENERAL).Text) = "" Then
            Call UniCode_Conv(K2_P_SUKEIRE.UKEIRE_DT, "")
        Else
            Call UniCode_Conv(K2_P_SUKEIRE.UKEIRE_DT, Format(Text1(ptxS_YMD).Text, "YYYYMMDD"))
        End If
    
        com = BtOpGetGreaterEqual
        
        Do
        
            DoEvents
        
            sts = BTRV(com, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K2_P_SUKEIRE, Len(K2_P_SUKEIRE), 2)
                
            Select Case sts
                Case BtNoErr
                    If Trim(Text1(ptxGENERAL).Text) = "" Then
                    Else
                        If Trim(Text1(ptxGENERAL).Text) <> Trim(StrConv(P_SUKEIRE_REC.TORI_CODE, vbUnicode)) Then
                            Exit Do
                        End If
                    End If
                
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "���i���w���������")
                    Exit Function
            End Select
    
    
    
            SKIP_Flg = False
            
        
        
                    
            '����N��������ڰ�
            If StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode) < Format(CDate(Text1(ptxS_YMD).Text), "YYYYMMDD") Or _
                StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode) > Format(CDate(Text1(ptxE_YMD).Text), "YYYYMMDD") Then
                SKIP_Flg = True
            End If
        
        
        
            '�w���ް��ǂݍ���
            Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_NO, StrConv(P_SUKEIRE_REC.SHIJI_NO, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            Select Case sts
                Case BtNoErr
                
                    If Trim(Text1(ptxGENERAL).Text) <> "" Then
                        If Trim(StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode)) <> Trim(Text1(ptxGENERAL).Text) Then
                            SKIP_Flg = True
                        End If
                    End If
                
                
                Case BtErrKeyNotFound
                    SKIP_Flg = True
                    Call UniCode_Conv(P_SSHIJI_O_REC.TORI_KBN, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "���i���w�}(�e)�ް�")
                    Exit Function
            End Select
            
            If StrConv(P_SSHIJI_O_REC.TORI_KBN, vbUnicode) <> P_TORI_GENERAL Then
                SKIP_Flg = True
            End If
    
If Trim(StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode)) = "27" Then
    Debug.Print
End If
    
            If Not SKIP_Flg Then
                Data_Flg = True
                                                '�����敪
                Call UniCode_Conv(P_SEISAN_DET_REC.TORI_KBN, StrConv(P_SSHIJI_O_REC.TORI_KBN, vbUnicode))
                                                '����溰��
                Call UniCode_Conv(P_SEISAN_DET_REC.TORI_CODE, StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode))
                                                '�����
                Call UniCode_Conv(P_SEISAN_DET_REC.UKEIRE_DT, StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode))
                                                '�w�}�[��
                Call UniCode_Conv(P_SEISAN_DET_REC.SHIJI_NO, StrConv(P_SUKEIRE_REC.SHIJI_NO, vbUnicode))
                                                '�d������
                Call UniCode_Conv(P_SEISAN_DET_REC.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                                                '�i��
                Call UniCode_Conv(P_SEISAN_DET_REC.HIN_GAI, StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))
                                                '����
                Call UniCode_Conv(P_SEISAN_DET_REC.UKEIRE_QTY, StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode))
                                                
                                                '���i���׽
                Call UniCode_Conv(P_SEISAN_DET_REC.S_CLASS_CODE, StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode))
                                                '�t���׽
                Call UniCode_Conv(P_SEISAN_DET_REC.F_CLASS_CODE, StrConv(P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode))
                                                '���E�׽
                Call UniCode_Conv(P_SEISAN_DET_REC.N_CLASS_CODE, StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode))
            
                wkTANKA = 0
                
                If Not GAICYU_F Then        ''2007.07.13
                
                    If Trim(StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode)) <> "" Then
                                                    '���i���P��
                        Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                        Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode))
                        sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                        Select Case sts
                            Case BtNoErr
                                wkTANKA = CDbl(StrConv(P_CLASSREC.KOURYOU, vbUnicode))
                            Case BtErrKeyNotFound
                                wkTANKA = 0
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "�׽Ͻ�")
                                Exit Function
                        End Select
                
                    End If
                
                    If Trim(StrConv(P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode)) <> "" Then
                                                    '�t���P��
                        Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                        Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, StrConv(P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode))
                        sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                        Select Case sts
                            Case BtNoErr
                                wkTANKA = wkTANKA + CDbl(StrConv(P_CLASSREC.KOURYOU, vbUnicode))
                            Case BtErrKeyNotFound
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "�׽Ͻ�")
                                Exit Function
                        End Select
                
                    End If
                End If
            
                If Trim(StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode)) <> "" Then
                                                '���E�P��
                    Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                    Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode))
                    sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                    Select Case sts
                        Case BtNoErr
                            wkTANKA = wkTANKA + CDbl(StrConv(P_CLASSREC.KOURYOU, vbUnicode))
                        Case BtErrKeyNotFound
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�׽Ͻ�")
                            Exit Function
                    End Select
            
                End If
            
            
            
            
                                                '�P��
                Call UniCode_Conv(P_SEISAN_DET_REC.KOURYOU, Format(wkTANKA, "00000000.00"))
                                                '���z
                Call UniCode_Conv(P_SEISAN_DET_REC.KIN, Format(wkTANKA * CDbl(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode)), "00000000000"))
                                                            
            
            
                sts = BTRV(BtOpInsert, P_SEISAN_DET_POS, P_SEISAN_DET_REC, Len(P_SEISAN_DET_REC), K0_P_SEISAN_DET, Len(K0_P_SEISAN_DET), 0)
                Select Case sts
                    Case BtNoErr
                    Case Else
                        Call File_Error(sts, BtOpInsert, "���Y���і����ް�")
                        Exit Function
                End Select
            
            
            End If
    
            com = BtOpGetNext
    
        Loop
    End If
    
    '----------------   ���E
    If Check1(pchkNAISYOKU).Value = vbChecked Then
    
        Call UniCode_Conv(K2_P_SUKEIRE.TORI_CODE, Text1(ptxNAISYOKU).Text)
        Call UniCode_Conv(K2_P_SUKEIRE.UKEIRE_DT, "")
    
    
        com = BtOpGetGreaterEqual
        
        Do
        
            DoEvents
        
            sts = BTRV(com, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K2_P_SUKEIRE, Len(K2_P_SUKEIRE), 2)
                
            Select Case sts
                Case BtNoErr
                    
                
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "���i���w���������")
                    Exit Function
            End Select
    
    
    
            SKIP_Flg = False
    
    
    
    
            '����N��������ڰ�
            If StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode) < Format(CDate(Text1(ptxS_YMD).Text), "YYYYMMDD") Or _
                StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode) > Format(CDate(Text1(ptxE_YMD).Text), "YYYYMMDD") Then
                SKIP_Flg = True
            End If
        
            '�w���ް��ǂݍ���
            Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_NO, StrConv(P_SUKEIRE_REC.SHIJI_NO, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
            Select Case sts
                Case BtNoErr
                
                
If Trim(StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode)) = "02" Then
Debug.Print
End If
                
                    If Trim(Text1(ptxNAISYOKU).Text) <> "" Then
                        If Trim(StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode)) <> Trim(Text1(ptxNAISYOKU).Text) Then
                            SKIP_Flg = True
                        End If
                    End If
                
                
                
                Case BtErrKeyNotFound
                    SKIP_Flg = True
                    Call UniCode_Conv(P_SSHIJI_O_REC.TORI_KBN, "")
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "���i���w�}(�e)�ް�")
                    Exit Function
            End Select
            
            If StrConv(P_SSHIJI_O_REC.TORI_KBN, vbUnicode) <> P_TORI_NAISYOKU Then
                SKIP_Flg = True
            End If
    
If Trim(StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode)) = "27" Then
    Debug.Print
End If
    
    
            If Not SKIP_Flg Then
                Data_Flg = True
                                                '�����敪
                Call UniCode_Conv(P_SEISAN_DET_REC.TORI_KBN, StrConv(P_SSHIJI_O_REC.TORI_KBN, vbUnicode))
                                                '����溰��
                Call UniCode_Conv(P_SEISAN_DET_REC.TORI_CODE, StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode))
                                                '�����
                Call UniCode_Conv(P_SEISAN_DET_REC.UKEIRE_DT, StrConv(P_SUKEIRE_REC.UKEIRE_DT, vbUnicode))
                                                '�w�}�[��
                Call UniCode_Conv(P_SEISAN_DET_REC.SHIJI_NO, StrConv(P_SUKEIRE_REC.SHIJI_NO, vbUnicode))
                                                '�d������
                Call UniCode_Conv(P_SEISAN_DET_REC.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                                                '�i��
                Call UniCode_Conv(P_SEISAN_DET_REC.HIN_GAI, StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))
                                                '����
                Call UniCode_Conv(P_SEISAN_DET_REC.UKEIRE_QTY, StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode))
                                                
                                                '���i���׽
                Call UniCode_Conv(P_SEISAN_DET_REC.S_CLASS_CODE, StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode))
                                                '�t���׽
                Call UniCode_Conv(P_SEISAN_DET_REC.F_CLASS_CODE, StrConv(P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode))
                                                '���E�׽
                Call UniCode_Conv(P_SEISAN_DET_REC.N_CLASS_CODE, StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode))
            
                wkTANKA = 0
            
'                If Trim(StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode)) <> "" Then
'                                                '���i���P��
'                    Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
'                    Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode))
'                    sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
'                    Select Case sts
'                        Case BtNoErr
'                            wkTANKA = CDbl(StrConv(P_CLASSREC.KOURYOU, vbUnicode))
'                        Case BtErrKeyNotFound
'                            wkTANKA = 0
'                        Case Else
'                            Call File_Error(sts, BtOpGetEqual, "�׽Ͻ�")
'                            Exit Function
'                    End Select
'
'                End If
            
'                If Trim(StrConv(P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode)) <> "" Then
'                                                '�t���P��
'                    Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
'                    Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, StrConv(P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode))
'                    sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
'                    Select Case sts
'                        Case BtNoErr
'                            wkTANKA = wkTANKA + CDbl(StrConv(P_CLASSREC.KOURYOU, vbUnicode))
'                        Case BtErrKeyNotFound
'                        Case Else
'                            Call File_Error(sts, BtOpGetEqual, "�׽Ͻ�")
'                            Exit Function
'                    End Select
'
'                End If
            
                
If StrConv(P_SSHIJI_O_REC.SHIJI_NO, vbUnicode) = "00233" Then
    Debug.Print
End If
                
                If Trim(StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode)) <> "" Then
                                                '���E�P��
                    Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                    Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode))
                    sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                    Select Case sts
                        Case BtNoErr
                            wkTANKA = wkTANKA + CDbl(StrConv(P_CLASSREC.KOURYOU, vbUnicode))
                        Case BtErrKeyNotFound
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�׽Ͻ�")
                            Exit Function
                    End Select
            
                End If
            
                                                '�P��
                Call UniCode_Conv(P_SEISAN_DET_REC.KOURYOU, Format(wkTANKA, "00000000.00"))
                                                '���z
                Call UniCode_Conv(P_SEISAN_DET_REC.KIN, Format(wkTANKA * CDbl(StrConv(P_SUKEIRE_REC.UKEIRE_QTY, vbUnicode)), "00000000000"))
                                                            
            
            
                sts = BTRV(BtOpInsert, P_SEISAN_DET_POS, P_SEISAN_DET_REC, Len(P_SEISAN_DET_REC), K0_P_SEISAN_DET, Len(K0_P_SEISAN_DET), 0)
                Select Case sts
                    Case BtNoErr
                    Case Else
                        Call File_Error(sts, BtOpInsert, "���Y���і����ް�")
                        Exit Function
                End Select
            
            
            End If
    
            com = BtOpGetNext
    
        Loop
    End If
    
    
    
    
    SAVE_TORI_CODE = ""
    
    ALL_CNT = 0
    ALL_QTY = 0
    For i = 0 To UBound(ALL_KIN)
        ALL_KIN(i) = 0
    Next i
    KAZEI = 0
    
    
    TOTAL_CNT = 0
    TOTAL_QTY = 0
    For i = 0 To UBound(TOTAL_KIN)
        TOTAL_KIN(i) = 0
    Next i
        
    
    com = BtOpGetFirst
    
    
    
    Do
    
    
        sts = BTRV(com, P_SEISAN_DET_POS, P_SEISAN_DET_REC, Len(P_SEISAN_DET_REC), K0_P_SEISAN_DET, Len(K0_P_SEISAN_DET), 0)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
            
                Exit Do
            
            
            Case Else
                Call File_Error(sts, com, "���Y���і����ް�")
                Exit Function
        End Select

    
        If com = BtOpGetFirst Then
            SAVE_TORI_KBN = StrConv(P_SEISAN_DET_REC.TORI_KBN, vbUnicode)
            SAVE_TORI_CODE = StrConv(P_SEISAN_DET_REC.TORI_CODE, vbUnicode)
        End If
    
    
        If SAVE_TORI_CODE <> StrConv(P_SEISAN_DET_REC.TORI_CODE, vbUnicode) Then
    
            If Sum_Total_Make_Proc(SAVE_TORI_KBN, SAVE_TORI_CODE, TOTAL_KIN(), TOTAL_CNT, TOTAL_QTY, 0) Then
                Exit Function
            End If
        
            ALL_CNT = ALL_CNT + TOTAL_CNT
            ALL_QTY = ALL_QTY + TOTAL_CNT
            
            For i = 0 To UBound(TOTAL_KIN)
                ALL_KIN(i) = ALL_KIN(i) + TOTAL_KIN(i)
            Next i
        
            If SAVE_TORI_KBN = P_TORI_NAISYOKU Then
            Else
                For i = 0 To UBound(TOTAL_KIN)
                    KAZEI = KAZEI + TOTAL_KIN(i)
                Next i
            End If
        
            TOTAL_CNT = 0
            TOTAL_QTY = 0
            For i = 0 To UBound(TOTAL_KIN)
                TOTAL_KIN(i) = 0
            Next i
        
            SAVE_TORI_KBN = StrConv(P_SEISAN_DET_REC.TORI_KBN, vbUnicode)
            SAVE_TORI_CODE = StrConv(P_SEISAN_DET_REC.TORI_CODE, vbUnicode)
        
        
        End If
        
        For i = 0 To UBound(SHIMUKE_TBL)
        
        
            If StrConv(P_SEISAN_DET_REC.SHIMUKE_CODE, vbUnicode) = SHIMUKE_TBL(i) Then
                TOTAL_KIN(i) = TOTAL_KIN(i) + CLng(StrConv(P_SEISAN_DET_REC.KIN, vbUnicode))
                Exit For
            End If
        
        Next i
        TOTAL_CNT = TOTAL_CNT + 1
        TOTAL_QTY = TOTAL_QTY + CDbl(StrConv(P_SEISAN_DET_REC.UKEIRE_QTY, vbUnicode))
        
        
        com = BtOpGetNext
    
    Loop
    
    If com <> BtOpGetFirst Then
        If Sum_Total_Make_Proc(SAVE_TORI_KBN, SAVE_TORI_CODE, TOTAL_KIN(), TOTAL_CNT, TOTAL_QTY, 0) Then
            Exit Function
        End If
    
        ALL_CNT = ALL_CNT + TOTAL_CNT
        ALL_QTY = ALL_QTY + TOTAL_CNT
        
        For i = 0 To UBound(TOTAL_KIN)
            ALL_KIN(i) = ALL_KIN(i) + TOTAL_KIN(i)
        Next i
    
        If SAVE_TORI_KBN = P_TORI_NAISYOKU Then
        Else
            For i = 0 To UBound(TOTAL_KIN)
                KAZEI = KAZEI + TOTAL_KIN(i)
            Next i
        End If
    
        If Sum_Total_Make_Proc("", "", ALL_KIN(), ALL_CNT, ALL_QTY, KAZEI) Then
            Exit Function
        End If
    
    
    
    
    End If
    
    
    
    

    PR000601.MousePointer = vbDefault

   SUM_Make_Proc = False

End Function






Private Function Sum_Total_Make_Proc(TORI_KBN As String, TORI_CODE As String, TOTAL_KIN() As Long, CNT As Integer, QTY As Double, KAZEI As Long) As Integer
'----------------------------------------------------------------------------
'           ���vں��ޏo��
'----------------------------------------------------------------------------
Dim i   As Integer
Dim sts As Integer
    
    Sum_Total_Make_Proc = True

    Call UniCode_Conv(P_SEISAN_GK_REC.TORI_KBN, TORI_KBN)       '�����敪
    Call UniCode_Conv(P_SEISAN_GK_REC.TORI_CODE, TORI_CODE)     '����溰��
                                                                
    For i = 0 To 9
        Call UniCode_Conv(P_SEISAN_GK_REC.UCHIWAKE_TBL(i).KIN, "00000000000")
    Next i
                                                                
                                                                
                                                                '�d������ʋ��z
    For i = 0 To UBound(TOTAL_KIN)
    
        Call UniCode_Conv(P_SEISAN_GK_REC.UCHIWAKE_TBL(i).KIN, Format(TOTAL_KIN(i), "00000000000"))
    
    Next i
                                                                '����
    Call UniCode_Conv(P_SEISAN_GK_REC.CNT, Format(CNT, "00000000000"))
                                                                '����
    Call UniCode_Conv(P_SEISAN_GK_REC.QTY, Format(QTY, "00000000.00"))
                                                                
                                                                
                                                                    '�ېőΏ�
    Call UniCode_Conv(P_SEISAN_GK_REC.KAZEI, Format(KAZEI, "00000000000"))


    sts = BTRV(BtOpInsert, P_SEISAN_GK_POS, P_SEISAN_GK_REC, Len(P_SEISAN_GK_REC), K0_P_SEISAN_GK, Len(K0_P_SEISAN_GK), 0)
    Select Case sts
        Case BtNoErr
        Case Else
            Call File_Error(sts, BtOpInsert, "���Y���і��׏W�v�ް�")
            Exit Function
    End Select

    Sum_Total_Make_Proc = False

End Function



Private Function SHIMUKE_TBL_Proc(i As Integer, KBN As String) As Integer
'----------------------------------------------------------------------------
'           �d�������ð���
'----------------------------------------------------------------------------

Dim com     As Integer
Dim sts     As Integer
Dim j       As Integer

    SHIMUKE_TBL_Proc = True

    ReDim Preserve SHIMUKE_TBL(0 To 9)

    For j = 0 To UBound(SHIMUKE_TBL)
        SHIMUKE_TBL(j) = ""
    Next j


    Call UniCode_Conv(K0_P_CODE.DATA_KBN, KBN)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")

    com = BtOpGetGreaterEqual
    i = -1

    Do
        DoEvents
    
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
                
                If Trim(StrConv(P_CODEREC.DATA_KBN, vbUnicode)) <> KBN Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "����Ͻ�")
                Exit Function
        End Select
    
        i = i + 1
        
        SHIMUKE_TBL(i) = Left(StrConv(P_CODEREC.C_Code, vbUnicode), 2)
            
    
    
        com = BtOpGetNext
    
    Loop

    If i = -1 Then
        SHIMUKE_TBL_Proc = False
        Exit Function
    End If
    
    
    j = colSHUMUKE01_KIN
    For i = 0 To UBound(SHIMUKE_TBL)
        
        If Trim(SHIMUKE_TBL(i)) = "" Then
            TDBGrid1(pGridDETAIL).Columns(j).Visible = False
        Else
        
            TDBGrid1(pGridDETAIL).Columns(j).Visible = True
            TDBGrid1(pGridDETAIL).Columns(j).Caption = SHIMUKE_TBL(i)
        End If
        j = j + 1
    Next i

    SHIMUKE_TBL_Proc = False

End Function
Private Function Item_Input_Proc() As Integer
'----------------------------------------------------------------------------
'                   ��ƊǗ����ד��͉�ʁ@�\��
'----------------------------------------------------------------------------
    Item_Input_Proc = True

    
    PR000602.Show vbModal                       '���ד��̓t�H�[���\��
    If G_SCREEN_FLG = SYS_ERR Then
        Exit Function
    End If

    

    Item_Input_Proc = False

End Function

Private Function D_Print_Proc() As Integer
'----------------------------------------------------------------------------
'           �������
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer




Dim rpt             As New PR00060F2
Dim f               As New PR000603
            
    
    D_Print_Proc = True
            
        
    com = BtOpGetFirst
    
    Do
        
        DoEvents
        
        sts = BTRV(com, P_SEISAN_GK_POS, P_SEISAN_GK_REC, Len(P_SEISAN_GK_REC), K0_P_SEISAN_GK, Len(K0_P_SEISAN_GK), 0)
            
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���Y���і��׏W�v�ް�")
                Exit Function
        End Select
    
        
        If Trim(StrConv(P_SEISAN_GK_REC.TORI_CODE, vbUnicode)) = "" Then
        Else
    
            Set rpt = New PR00060F2
        
            '���|�[�g��������܂��B�itrue�F����_�C�A���O���� false�F�Ȃ��j
            rpt.PrintReport False
        
            Set rpt = Nothing
            
            
'            f.RunReport rpt
'            f.Show
        End If
    
        com = BtOpGetNext
    
    Loop
        
        
 
 
 
 
 
    D_Print_Proc = False



End Function

