VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form SEI00701 
   Caption         =   "[�����V�X�e��]�~�j�}���������쐬����"
   ClientHeight    =   11145
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   17715
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
   ScaleWidth      =   17715
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1365
      TabIndex        =   0
      Top             =   1080
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��  ��"
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
      TabIndex        =   6
      Top             =   120
      Width           =   1380
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   2730
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   2
      Top             =   2040
      Width           =   4845
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   1365
      TabIndex        =   1
      Top             =   2040
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   10500
      TabIndex        =   4
      Top             =   2040
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   8820
      TabIndex        =   3
      Top             =   2040
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�I  ��"
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
      Left            =   3570
      TabIndex        =   8
      Top             =   120
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   9
      Top             =   10680
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXCEL"
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
      Left            =   1890
      TabIndex        =   7
      Top             =   120
      Width           =   1380
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   6975
      Left            =   -105
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2520
      Width           =   17025
      _ExtentX        =   30030
      _ExtentY        =   12303
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "������t"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "������"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "�v��N��"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "�����"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "�����敪"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "�o�c����"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "����"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "�������ځi��o�p�j"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "�������ځi�r�c�b�p�j"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "����"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "�P��"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "���z"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "�����"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "�E�v"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   14
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=14"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2355"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2223"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2196"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2064"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1931"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1799"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=4630"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=4498"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=2223"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2090"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=2090"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1958"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=1879"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1746"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=4128"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=3995"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=4842"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=4710"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=2249"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2117"
      Splits(0)._ColumnProps(40)=   "Column(9)._ColStyle=2"
      Splits(0)._ColumnProps(41)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(42)=   "Column(10).Width=2249"
      Splits(0)._ColumnProps(43)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(10)._WidthInPix=2117"
      Splits(0)._ColumnProps(45)=   "Column(10)._ColStyle=2"
      Splits(0)._ColumnProps(46)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(47)=   "Column(11).Width=2249"
      Splits(0)._ColumnProps(48)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(11)._WidthInPix=2117"
      Splits(0)._ColumnProps(50)=   "Column(11)._ColStyle=2"
      Splits(0)._ColumnProps(51)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(52)=   "Column(12).Width=2249"
      Splits(0)._ColumnProps(53)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(12)._WidthInPix=2117"
      Splits(0)._ColumnProps(55)=   "Column(12)._ColStyle=2"
      Splits(0)._ColumnProps(56)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(57)=   "Column(13).Width=4366"
      Splits(0)._ColumnProps(58)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(13)._WidthInPix=4233"
      Splits(0)._ColumnProps(60)=   "Column(13).Order=14"
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
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=87"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=110,.parent=87"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=107,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=108,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=109,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=114,.parent=87"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=111,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=112,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=113,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=118,.parent=87"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=115,.parent=88"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=116,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=117,.parent=91"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=16,.parent=87"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=13,.parent=88"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=14,.parent=89"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=15,.parent=91"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=20,.parent=87"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=17,.parent=88"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=18,.parent=89"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=19,.parent=91"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=24,.parent=87"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=21,.parent=88"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=22,.parent=89"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=23,.parent=91"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=28,.parent=87,.alignment=1"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=25,.parent=88"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=26,.parent=89"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=27,.parent=91"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=46,.parent=87,.alignment=1"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=43,.parent=88"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=44,.parent=89"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=45,.parent=91"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=50,.parent=87,.alignment=1"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=47,.parent=88"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=48,.parent=89"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=49,.parent=91"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=54,.parent=87,.alignment=1"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=51,.parent=88"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=52,.parent=89"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=53,.parent=91"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=58,.parent=87"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=55,.parent=88"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=56,.parent=89"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=57,.parent=91"
      _StyleDefs(92)  =   "Named:id=33:Normal"
      _StyleDefs(93)  =   ":id=33,.parent=0"
      _StyleDefs(94)  =   "Named:id=34:Heading"
      _StyleDefs(95)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(96)  =   ":id=34,.wraptext=-1"
      _StyleDefs(97)  =   "Named:id=35:Footing"
      _StyleDefs(98)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(99)  =   "Named:id=36:Selected"
      _StyleDefs(100) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(101) =   "Named:id=37:Caption"
      _StyleDefs(102) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(103) =   "Named:id=38:HighlightRow"
      _StyleDefs(104) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(105) =   "Named:id=39:EvenRow"
      _StyleDefs(106) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(107) =   "Named:id=40:OddRow"
      _StyleDefs(108) =   ":id=40,.parent=33"
      _StyleDefs(109) =   "Named:id=41:RecordSelector"
      _StyleDefs(110) =   ":id=41,.parent=34"
      _StyleDefs(111) =   "Named:id=42:FilterBar"
      _StyleDefs(112) =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '����
      Caption         =   "������"
      Height          =   375
      Index           =   3
      Left            =   210
      TabIndex        =   13
      Top             =   1080
      Width           =   1170
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '����
      Caption         =   "�����"
      Height          =   375
      Index           =   4
      Left            =   210
      TabIndex        =   12
      Top             =   2040
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "�`"
      Height          =   375
      Index           =   8
      Left            =   10185
      TabIndex        =   11
      Top             =   2160
      Width           =   330
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '����
      Caption         =   "���t�͈�"
      Height          =   375
      Index           =   7
      Left            =   7665
      TabIndex        =   10
      Top             =   2040
      Width           =   1170
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "�����I��"
      Begin VB.Menu SHORI 
         Caption         =   "�X�V"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "�I��"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "��ʈ��"
         Index           =   2
      End
   End
End
Attribute VB_Name = "SEI00701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxS_UKEHARAI_CODE% = 0   '�����@�I��


Private Const ptxDEN_NO% = 0            '�`�[��

Private Const ptxUKEHARAI_CODE% = 1     '�󕥐�
Private Const ptxS_JITU_DATE% = 2       '���t�͈́@�J�n
Private Const ptxE_JITU_DATE% = 3       '���t�͈́@�I��



Private Const pcmbUKEHARAI% = 0         '�����



Dim SE_MIN_URIAGE   As New XArrayDB

Private Const Min_Row% = 1              '�ŏ��s��

Dim Max_Row    As Integer               '�O���b�h�ő�\������


Private Const Min_Col% = 0              '�ŏ���
Private Const Max_Col% = 13             '�ő��

Private Const ColJITU_DATE% = 0         '������t
Private Const ColDEN_NO% = 1            '�`�[��

Private Const ColKEIJYO_YM% = 2         '�v��N��

Private Const ColUKEHARAI% = 3          '�����

Private Const ColSE_KBN% = 4            '�����敪
Private Const ColMANA_KBN% = 5          '�o�c����
Private Const ColPOST_CODE% = 6         '����
Private Const ColSUB_ITEM% = 7          '�������ځi��o�p�j
Private Const ColSDC_ITEM% = 8          '�������ځiSDC�p�j
Private Const ColSURYO% = 9             '����
Private Const ColTANKA% = 10            '�P��
Private Const ColURI_KIN% = 11          '���z
Private Const ColZEI_KIN% = 12          '�����
Private Const ColTEKIYO% = 13           '�E�v






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

'�����敪
Private Type SE_KBN_Tag
    No          As Integer
    SE_KBN      As String
End Type
Private SE_KBN()    As SE_KBN_Tag



Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)


    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    Call Tab_Ctrl(Shift)        '�ړ�


End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    
    Select Case Index
    
        Case pcmbUKEHARAI
            Text1(ptxS_UKEHARAI_CODE).Text = Trim(Right(Combo1(Index).Text, 5))
    
    End Select




End Sub

Private Sub Command1_Click(Index As Integer)

Dim yn      As Integer
Dim i       As Integer







    Select Case Index
    
        Case 0          '����
        
            If List_Disp_Proc() Then
                Unload Me
            End If
        
        
        
        
        Case 1          'EXCEL
        
        
        
            If Trim(Text1(ptxDEN_NO).Text) = "" Then
                MsgBox "����Ώۂ̐���������͂��ĉ������B"
                Exit Sub
            End If
        
            Beep
            yn = MsgBox("�������쐬���܂����H", vbYesNo + vbQuestion, "�m�F����")
            
            If yn = vbYes Then
                If COVER_Proc() Then
                    Unload Me
                End If
            End If
                
        
        Case 2          '�I��
            Unload Me
    
    
    
    
    End Select



End Sub




Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
Private Sub Form_Load()

Dim i           As Integer
Dim c           As String * 128
Dim sts         As Integer

Dim S_DATE      As String
Dim E_Date      As String
Dim S_YY        As String * 4
Dim S_MM        As String * 2
Dim S_DD        As String * 2


    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If


    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[�����V�X�e��]�~�j�}���������쐬����", Me.hwnd, 0)
    
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
                                
                                
                                
                                '�Ǘ��}�X�^�n�o�d�m
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�R�[�h�}�X�^�n�o�d�m
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�󕥐�}�X�^�n�o�d�m
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '������тn�o�d�m
    If SE_MIN_URIAGE_Open(BtOpenNomal) Then
        Unload Me
    End If







    '�󕥐�
    If Ukeharai_Set_Proc() Then
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



'�����敪��荞��
    i = 0
    Do
        i = i + 1
        If GetIni("SE_KBN", Format(i, "00"), "SEI_SYS", c) Then
            Exit Do
        End If
    
        ReDim Preserve SE_KBN(0 To i - 1)
    
        SE_KBN(i - 1).No = i
        SE_KBN(i - 1).SE_KBN = Trim(c)
    
    Loop







    '�����\��
    E_Date = Format(Now, "YYYY/MM/DD")
    S_DATE = DateAdd("m", -1, Left(E_Date, 8) & SHIMEBI)
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


    Text1(ptxS_JITU_DATE).Text = S_DATE
    Text1(ptxE_JITU_DATE).Text = E_Date
    
'    If List_Disp_Proc() Then
'        Unload Me
'    End If

    Text1(ptxS_UKEHARAI_CODE).SetFocus

End Sub



Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
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
            Call File_Error(sts, BtOpClose, "�󕥐�}�X�^�}�X�^")
        End If
    End If
                                            '������тb�k�n�r�d
    sts = BTRV(BtOpClose, SE_MIN_URIAGE_POS, SE_MIN_URIAGEREC, Len(SE_MIN_URIAGEREC), K0_SE_MIN_URIAGE, Len(K0_SE_MIN_URIAGE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "������у}�X�^�}�X�^")
        End If
    End If
    
    
    
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    End
End Sub




Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    SEI00701.MousePointer = vbHourglass

    Call Ctrl_Lock(SEI00701)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(SEI00701)


    SEI00701.MousePointer = vbDefault

End Sub
Private Function List_Disp_Proc(Optional Den_No As String = " ") As Integer
'----------------------------------------------------------------------------
'                   �w��͈͂̔���f�[�^��\������
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim Row         As Long
    
Dim E_Date      As String
    
Dim Skip_F      As Boolean
    
    List_Disp_Proc = True
                                    
    Call Input_Lock
                                    
                                    
                                    '�e�[�u�����Z�b�g
    Set SE_MIN_URIAGE = Nothing
                                    '������ѓǂݍ��݊J�n
    
    
    If IsDate(Text1(ptxS_JITU_DATE).Text) Then
        Call UniCode_Conv(K0_SE_MIN_URIAGE.JITU_DATE, Format(Text1(ptxS_JITU_DATE).Text, "YYYYMMDD"))
    Else
        Call UniCode_Conv(K0_SE_MIN_URIAGE.JITU_DATE, "")
    End If
    
    
    Call UniCode_Conv(K0_SE_MIN_URIAGE.Den_No, Den_No)
    Call UniCode_Conv(K0_SE_MIN_URIAGE.GYO_NO, "")
    
    
    
    
    If IsDate(Text1(ptxE_JITU_DATE).Text) Then
        E_Date = Format(Text1(ptxE_JITU_DATE).Text, "YYYYMMDD")
    Else
        E_Date = Text1(ptxS_JITU_DATE).Text
    End If
    
        
    
        
    
    
    
    
    Row = Min_Row - 1
        
    
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        sts = BTRV(com, SE_MIN_URIAGE_POS, SE_MIN_URIAGEREC, Len(SE_MIN_URIAGEREC), K0_SE_MIN_URIAGE, Len(K0_SE_MIN_URIAGE), 0)
        
    
    
        Select Case sts
            Case BtNoErr
        
                If StrConv(SE_MIN_URIAGEREC.JITU_DATE, vbUnicode) > E_Date Then
                    Exit Do
                End If
        
                If Trim(Den_No) <> "" Then
                    If Trim(StrConv(SE_MIN_URIAGEREC.Den_No, vbUnicode)) <> Trim(Den_No) Then
                        Exit Do
                    End If
                
                End If
        
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�������")
                Exit Function
        End Select
            
        
        Skip_F = False
        If Trim(Text1(ptxUKEHARAI_CODE).Text) <> "" Then
            If Trim(Text1(ptxUKEHARAI_CODE).Text) <> Trim(StrConv(SE_MIN_URIAGEREC.UKEHARAI_CODE, vbUnicode)) Then
                Skip_F = True
            End If
        End If
        
        
        
        
        If Not Skip_F Then
        
        
        
            Row = Row + 1
                        
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
        
        
        End If
        
        com = BtOpGetNext
        
    Loop
    
                                
                                'DB�e�[�u�������N
    Set TDBGrid1.Array = SE_MIN_URIAGE
    
    TDBGrid1.style.Locked = True
    
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    
    
    
    Call Input_UnLock
    
    
    List_Disp_Proc = False

    
End Function

Private Function Grid_Set_Proc(Row As Long) As Integer
'----------------------------------------------------------------------------
'                   ����f�[�^---��Grid
'----------------------------------------------------------------------------

Dim sts As Integer
Dim i   As Integer
    
    Grid_Set_Proc = True

    SE_MIN_URIAGE.ReDim Min_Row, Row, Min_Col, Max_Col


    SE_MIN_URIAGE(Row, ColJITU_DATE) = Mid(StrConv(SE_MIN_URIAGEREC.JITU_DATE, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(SE_MIN_URIAGEREC.JITU_DATE, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(SE_MIN_URIAGEREC.JITU_DATE, vbUnicode), 7, 2)
    SE_MIN_URIAGE(Row, ColDEN_NO) = StrConv(SE_MIN_URIAGEREC.Den_No, vbUnicode)
    
    
    SE_MIN_URIAGE(Row, ColKEIJYO_YM) = Mid(StrConv(SE_MIN_URIAGEREC.KEIJYO_YM, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(SE_MIN_URIAGEREC.KEIJYO_YM, vbUnicode), 5, 2)

    
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(SE_MIN_URIAGEREC.UKEHARAI_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    
    Select Case sts
        Case BtNoErr
    
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�󕥃}�X�^")
            Exit Function
    End Select
    SE_MIN_URIAGE(Row, ColUKEHARAI) = StrConv(SE_MIN_URIAGEREC.UKEHARAI_CODE, vbUnicode) & " " & Trim(StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode))
    
    SE_MIN_URIAGE(Row, ColSE_KBN) = ""
    For i = 0 To UBound(SE_KBN)
    
        If SE_KBN(i).No = StrConv(SE_MIN_URIAGEREC.SE_KBN, vbUnicode) Then
            SE_MIN_URIAGE(Row, ColSE_KBN) = SE_KBN(i).No & " " & SE_KBN(i).SE_KBN
            Exit For
        End If
    
    Next i
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN09_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(SE_MIN_URIAGEREC.MANA_KBN, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    
    Select Case sts
        Case BtNoErr
    
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�R�[�h�}�X�^")
            Exit Function
    End Select
    
    
    SE_MIN_URIAGE(Row, ColMANA_KBN) = StrConv(SE_MIN_URIAGEREC.MANA_KBN, vbUnicode) & " " & Trim(StrConv(P_CODEREC.C_NAME, vbUnicode))
    
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN10_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(SE_MIN_URIAGEREC.POST_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    
    Select Case sts
        Case BtNoErr
    
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�R�[�h�}�X�^")
            Exit Function
    End Select
    
    
    SE_MIN_URIAGE(Row, ColPOST_CODE) = StrConv(SE_MIN_URIAGEREC.POST_CODE, vbUnicode) & " " & Trim(StrConv(P_CODEREC.C_NAME, vbUnicode))
    
    SE_MIN_URIAGE(Row, ColSUB_ITEM) = Trim(StrConv(SE_MIN_URIAGEREC.SUB_ITEM, vbUnicode))
    SE_MIN_URIAGE(Row, ColSDC_ITEM) = Trim(StrConv(SE_MIN_URIAGEREC.SDC_ITEM, vbUnicode))
    
    SE_MIN_URIAGE(Row, ColSURYO) = Format(StrConv(SE_MIN_URIAGEREC.SURYO, vbUnicode), "#,##0.00")
    SE_MIN_URIAGE(Row, ColTANKA) = Format(StrConv(SE_MIN_URIAGEREC.TANKA, vbUnicode), "#,##0.00")
    SE_MIN_URIAGE(Row, ColURI_KIN) = Format(StrConv(SE_MIN_URIAGEREC.URI_KIN, vbUnicode), "#,##0")
    SE_MIN_URIAGE(Row, ColZEI_KIN) = Format(StrConv(SE_MIN_URIAGEREC.ZEI_KIN, vbUnicode), "#,##0")
    
    SE_MIN_URIAGE(Row, ColTEKIYO) = Trim(StrConv(SE_MIN_URIAGEREC.TEKIYO, vbUnicode))
    

    SE_MIN_URIAGE.ReDim Min_Row, Row, Min_Col, Max_Col
    
    
    Grid_Set_Proc = False
End Function


Private Function Ukeharai_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   �󕥐�}�X�^���R���{�ɃZ�b�g����B
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer




Dim i           As Integer
    
    Ukeharai_Set_Proc = True
    
    Combo1(pcmbUKEHARAI).Clear
    
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

        
        
        Combo1(pcmbUKEHARAI).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode) & "            " & _
                                StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        
        com = BtOpGetNext
    
    Loop

    Ukeharai_Set_Proc = False
    



End Function


Private Sub TDBGrid1_DblClick()

    
    Text1(ptxDEN_NO).Text = SE_MIN_URIAGE(TDBGrid1.Bookmark, ColDEN_NO)
    

    If List_Disp_Proc() Then
        Unload Me
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

Private Function Error_Check_Proc(mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
    
Dim sts As Integer
    
Dim i   As Integer
    
    
    Error_Check_Proc = True
    
    Select Case mode
    
    
    
        Case ptxS_UKEHARAI_CODE   '�����
            Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxS_UKEHARAI_CODE).Text)
            
            sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
            Select Case sts
                Case BtNoErr
                
                    For i = 0 To Combo1(pcmbUKEHARAI).ListCount - 1
                    
                        If Trim(Text1(ptxS_UKEHARAI_CODE).Text) = Trim(Right(Combo1(pcmbUKEHARAI).List(i), 5)) Then
                        
                            Combo1(pcmbUKEHARAI).ListIndex = i
                            Exit For
                        
                        End If
                    
                    Next i
                
                
                Case BtErrKeyNotFound
                    Combo1(pcmbUKEHARAI).ListIndex = -1
                    MsgBox "���͂������ڂ̓G���[�ł��B(�����)"
                    Text1(mode).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�󕥐�}�X�^")
                    Exit Function
            End Select
    
    
    
        Case ptxS_JITU_DATE '���t�͈́@�J�n
            
            If Trim(Text1(ptxS_JITU_DATE).Text) = "" Then
            Else
                If Not IsDate(Text1(ptxS_JITU_DATE).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(���t�͈́@�J�n)"
                    Text1(mode).SetFocus
                    Exit Function
                End If
            End If
    
        Case ptxS_JITU_DATE '���t�͈́@�I��
            
            If Trim(Text1(ptxE_JITU_DATE).Text) = "" Then
            Else
                If Not IsDate(Text1(ptxE_JITU_DATE).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(���t�͈́@�I��)"
                    Text1(mode).SetFocus
                    Exit Function
                End If
            End If
    
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function

Private Function COVER_Proc(Optional Den_No As String = " ") As Integer
'----------------------------------------------------------------------------
'                   �d�w�b�d�k�i�\���j�o��
'----------------------------------------------------------------------------

Dim excelApplication    As Excel.Application
Dim excelWorkBook       As Excel.Workbook
Dim excelSheet          As Excel.Worksheet

Dim i                   As Integer
Dim j                   As Integer
Dim Row                 As Integer
    
    
Dim sts                 As Integer
Dim com                 As Integer
Dim Skip_F              As Boolean
    
    
Dim E_Date              As String


Dim GK_KINGAKU          As Long
Dim GK_ZEIKIN           As Long


Dim c                   As String * 128
    
Dim Name1               As String
Dim Name2               As String
    
Dim ITEM                As String

Dim ADDR1               As String
Dim ADDR2               As String

Dim SYAMEI              As String

Dim BIKOU1              As String
Dim BIKOU2              As String
Dim BIKOU3              As String

Dim HIN_NAME            As String

Dim TEKIYO              As String

Dim SHIMEBI             As String
    

    COVER_Proc = True
    
    Call Input_Lock
    
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
    If GetIni(App.EXEName, "HIN_NAME", App.EXEName, c) Then
        HIN_NAME = ""
    Else
        HIN_NAME = Trim(c)
    End If
    If GetIni(App.EXEName, "TEKIYO", App.EXEName, c) Then
        TEKIYO = ""
    Else
        TEKIYO = Trim(c)
    End If
    If GetIni(App.EXEName, "SHIMEBI", App.EXEName, c) Then
        SHIMEBI = ""
    Else
        SHIMEBI = Trim(c)
    End If

    
    Set excelApplication = CreateObject("Excel.Application")
    excelApplication.Visible = True


    
    Set excelWorkBook = excelApplication.Workbooks.Add
'    Set excelSheet = excelWorkBook.Worksheets.Add
    Set excelSheet = excelWorkBook.Worksheets(1)
    

    
    excelApplication.StandardFontSize = 11
    
    excelApplication.StandardFont = "�l�r�@�S�V�b�N"

    
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
    
    
    
    
    
    
    
    
    If IsDate(Text1(ptxS_JITU_DATE).Text) Then
        Call UniCode_Conv(K0_SE_MIN_URIAGE.JITU_DATE, Format(Text1(ptxS_JITU_DATE).Text, "YYYYMMDD"))
    Else
        Call UniCode_Conv(K0_SE_MIN_URIAGE.JITU_DATE, "")
    End If
    
    
    Call UniCode_Conv(K0_SE_MIN_URIAGE.Den_No, Den_No)
    Call UniCode_Conv(K0_SE_MIN_URIAGE.GYO_NO, "")
    
    
    
    
    If IsDate(Text1(ptxE_JITU_DATE).Text) Then
        E_Date = Format(Text1(ptxE_JITU_DATE).Text, "YYYYMMDD")
    Else
        E_Date = Text1(ptxS_JITU_DATE).Text
    End If
    
        
    
        
    
    
    
    
    Row = Min_Row - 1
        
    
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        sts = BTRV(com, SE_MIN_URIAGE_POS, SE_MIN_URIAGEREC, Len(SE_MIN_URIAGEREC), K0_SE_MIN_URIAGE, Len(K0_SE_MIN_URIAGE), 0)
        
    
    
        Select Case sts
            Case BtNoErr
        
                If StrConv(SE_MIN_URIAGEREC.JITU_DATE, vbUnicode) > E_Date Then
                    Exit Do
                End If
        
                If Trim(Den_No) <> "" Then
                    If Trim(StrConv(SE_MIN_URIAGEREC.Den_No, vbUnicode)) <> Trim(Den_No) Then
                        Exit Do
                    End If
                
                End If
        
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�������")
                Exit Function
        End Select
            
        
        Skip_F = False
        If Trim(Text1(ptxUKEHARAI_CODE).Text) <> "" Then
            If Trim(Text1(ptxUKEHARAI_CODE).Text) <> Trim(StrConv(SE_MIN_URIAGEREC.UKEHARAI_CODE, vbUnicode)) Then
                Skip_F = True
            End If
        End If
        
        
        
        
        If Not Skip_F Then
        
        
        
            '���^��
            excelSheet.Application.Range(excelSheet.Application.Cells(15, 1), excelSheet.Application.Cells(15, 1)).HorizontalAlignment = xlCenter
            excelSheet.Application.Range(excelSheet.Application.Cells(15, 1), excelSheet.Application.Cells(15, 1)).Select
            excelSheet.Application.Selection.NumberFormatLocal = "@"
            excelSheet.Application.Selection.HorizontalAlignment = xlCenter
            excelSheet.Application.Cells(15, 1).Value = Format(CInt(Mid(Format(Now, "YYYYMMDD"), 5, 2)), "#") & "/" & SHIMEBI
            '�i��
            excelSheet.Application.Cells(15, 2).Value = Trim(StrConv(SE_MIN_URIAGEREC.SUB_ITEM, vbUnicode))
            '����
            excelSheet.Application.Range(excelSheet.Application.Cells(15, 3), excelSheet.Application.Cells(15, 3)).Select
            excelSheet.Application.Selection.NumberFormatLocal = "#,##0"
            excelSheet.Application.Cells(15, 3).Value = Format(CDbl(StrConv(SE_MIN_URIAGEREC.SURYO, vbUnicode)), "#,##0.00")
            '�P���`���z
            excelSheet.Application.Range(excelSheet.Application.Cells(15, 4), excelSheet.Application.Cells(15, 5)).Select
            excelSheet.Application.Selection.NumberFormatLocal = "#,##0"
            excelSheet.Application.Cells(15, 4).Value = Format(CDbl(StrConv(SE_MIN_URIAGEREC.TANKA, vbUnicode)), "#,##0.00")
            excelSheet.Application.Cells(15, 5).Value = Format(CLng(StrConv(SE_MIN_URIAGEREC.URI_KIN, vbUnicode)), "#,##0")
            '�E�v
            excelSheet.Application.Cells(15, 6).Value = Trim(StrConv(SE_MIN_URIAGEREC.TEKIYO, vbUnicode))
        
        
            GK_KINGAKU = GK_KINGAKU + CLng(StrConv(SE_MIN_URIAGEREC.URI_KIN, vbUnicode))
            GK_ZEIKIN = GK_ZEIKIN + CLng(StrConv(SE_MIN_URIAGEREC.ZEI_KIN, vbUnicode))
        
        
        
        
        
        
        End If
        
        com = BtOpGetNext
        
    Loop
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    '�Ŕ������z
    excelSheet.Application.Range(excelSheet.Application.Cells(29, 5), excelSheet.Application.Cells(31, 5)).Select
    excelSheet.Application.Selection.NumberFormatLocal = "#,##0;""�� ""#,##0"


    excelSheet.Application.Cells(29, 5).Value = GK_KINGAKU
    '�����
    excelSheet.Application.Cells(30, 5).Value = GK_ZEIKIN
    '�ō��݋��z
    excelSheet.Application.Cells(31, 5).Value = GK_KINGAKU + GK_ZEIKIN
    '���v���z
    excelSheet.Application.Cells(12, 4).Value = Format(GK_KINGAKU + GK_ZEIKIN, "\\#,##0")


    Set excelSheet = Nothing
    Set excelWorkBook = Nothing
    Set excelApplication = Nothing


    
    Call Input_UnLock
    COVER_Proc = False
    

End Function


