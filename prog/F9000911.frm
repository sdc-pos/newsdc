VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form F9000911 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�ڊǎ���m�F 2009.03.05"
   ClientHeight    =   12975
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   18255
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   12975
   ScaleWidth      =   18255
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Height          =   360
      Index           =   4
      Left            =   8190
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   24
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Height          =   360
      Index           =   3
      Left            =   7350
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   22
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   360
      Index           =   2
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   360
      Index           =   1
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   360
      Index           =   0
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command 
      Caption         =   "�I  ��"
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
      Left            =   10425
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   11760
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   9585
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   11760
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   8745
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   11760
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�f�[�^"
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
      Left            =   7905
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   11760
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�Ł@�V"
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
      Left            =   6585
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   11760
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   5745
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   11760
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   5
      Left            =   4905
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   11760
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   4
      Left            =   4065
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   11760
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   2745
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   11760
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   1905
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   11760
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   1065
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   11760
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   225
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   11760
      Width           =   855
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   10935
      Left            =   0
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   600
      Width           =   18165
      _ExtentX        =   32041
      _ExtentY        =   19288
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "�o�ד��t"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "ID��"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "�`�[��"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "���x"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "�i�ԁi�O���j"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "�i�@��"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "�������"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "�I�ԂP"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "�I�ԂQ"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "�I�ԂR"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "�������"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "����S����"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   13
      Splits(0)._UserFlags=   0
      Splits(0).Locked=   -1  'True
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=13"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2514"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2408"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2408"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2302"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1349"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1244"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=1455"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1349"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=2646"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2540"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2619"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=1773"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1667"
      Splits(0)._ColumnProps(28)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(30)=   "Column(7).Width=2037"
      Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=1931"
      Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(34)=   "Column(8).Width=2037"
      Splits(0)._ColumnProps(35)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(8)._WidthInPix=1931"
      Splits(0)._ColumnProps(37)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(38)=   "Column(9).Width=2037"
      Splits(0)._ColumnProps(39)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(9)._WidthInPix=1931"
      Splits(0)._ColumnProps(41)=   "Column(9)._ColStyle=0"
      Splits(0)._ColumnProps(42)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(43)=   "Column(10).Width=3281"
      Splits(0)._ColumnProps(44)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(10)._WidthInPix=3175"
      Splits(0)._ColumnProps(46)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(47)=   "Column(11).Width=3969"
      Splits(0)._ColumnProps(48)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(11)._WidthInPix=3863"
      Splits(0)._ColumnProps(50)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(51)=   "Column(12).Width=2328"
      Splits(0)._ColumnProps(52)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(12)._WidthInPix=2223"
      Splits(0)._ColumnProps(54)=   "Column(12).Order=13"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=10.5,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �S�V�b�N"
      PrintInfos(0).PageFooterFont=   "Size=10.5,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �S�V�b�N"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=�l�r �S�V�b�N"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bold=0,.fontsize=1050,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=�l�r �S�V�b�N"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.bold=0,.fontsize=1050,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=�l�r �S�V�b�N"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=31,.bold=0,.fontsize=1050,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=�l�r �S�V�b�N"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=87,.parent=2,.namedParent=89"
      _StyleDefs(23)  =   "FilterBarStyle:id=90,.parent=1,.namedParent=92"
      _StyleDefs(24)  =   "Splits(0).Style:id=53,.parent=1,.bgcolor=&HFFFF&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=62,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=54,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=55,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=56,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=58,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=57,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=59,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=60,.parent=9,.bgcolor=&HFF80&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=61,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=88,.parent=87"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=91,.parent=90"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=53"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=54"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=55"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=57"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=48,.parent=53"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=45,.parent=54"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=46,.parent=55"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=47,.parent=57"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=66,.parent=53"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=54"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=55"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=57"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=102,.parent=53"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=19,.parent=54"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=20,.parent=55"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=101,.parent=57"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=70,.parent=53"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=54"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=55"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=57"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=78,.parent=53"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=75,.parent=54"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=76,.parent=55"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=77,.parent=57"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=24,.parent=53,.alignment=1"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=21,.parent=54"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=22,.parent=55"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=23,.parent=57"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=18,.parent=53"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=15,.parent=54"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=16,.parent=55"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=17,.parent=57"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=14,.parent=53"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=11,.parent=54"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=12,.parent=55"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=13,.parent=57"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=82,.parent=53,.alignment=0,.locked=0"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=79,.parent=54,.alignment=3"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=80,.parent=55,.alignment=3"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=81,.parent=57"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=52,.parent=53"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=49,.parent=54"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=50,.parent=55"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=51,.parent=57"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=100,.parent=53"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=97,.parent=54"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=98,.parent=55"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=99,.parent=57"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=40,.parent=53"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=37,.parent=54"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=38,.parent=55"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=39,.parent=57"
      _StyleDefs(88)  =   "Named:id=29:Normal"
      _StyleDefs(89)  =   ":id=29,.parent=0"
      _StyleDefs(90)  =   "Named:id=30:Heading"
      _StyleDefs(91)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(92)  =   ":id=30,.wraptext=-1"
      _StyleDefs(93)  =   "Named:id=31:Footing"
      _StyleDefs(94)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(95)  =   "Named:id=32:Selected"
      _StyleDefs(96)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(97)  =   "Named:id=33:Caption"
      _StyleDefs(98)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(99)  =   "Named:id=34:HighlightRow"
      _StyleDefs(100) =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(101) =   "Named:id=35:EvenRow"
      _StyleDefs(102) =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(103) =   "Named:id=36:OddRow"
      _StyleDefs(104) =   ":id=36,.parent=29"
      _StyleDefs(105) =   "Named:id=89:RecordSelector"
      _StyleDefs(106) =   ":id=89,.parent=30"
      _StyleDefs(107) =   "Named:id=92:FilterBar"
      _StyleDefs(108) =   ":id=92,.parent=29"
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�^"
      Height          =   240
      Index           =   1
      Left            =   7980
      TabIndex        =   23
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "��������^�`�[����"
      Height          =   240
      Index           =   0
      Left            =   5145
      TabIndex        =   21
      Top             =   240
      Width           =   2160
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   240
      Index           =   7
      Left            =   4080
      TabIndex        =   19
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   240
      Index           =   6
      Left            =   3480
      TabIndex        =   18
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�N"
      Height          =   240
      Index           =   5
      Left            =   2880
      TabIndex        =   17
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�o�ד��t"
      Height          =   240
      Index           =   4
      Left            =   1080
      TabIndex        =   16
      Top             =   240
      Width           =   960
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
      Left            =   225
      TabIndex        =   15
      Top             =   12240
      Width           =   180
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Checked         =   -1  'True
         Index           =   0
      End
   End
End
Attribute VB_Name = "F9000911"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Dim OUT_FILE    As String               '�o�̓t�@�C��




Private Const ptxSyuka_YY% = 0          '�o�ד��@�N
Private Const ptxSyuka_MM% = 1          '�o�ד��@��
Private Const ptxSyuka_DD% = 2          '�o�ד��@��


Private Const ptxKAN_MAISU% = 3         '���i����
Private Const ptxDEN_MAISU% = 4         '�`�[����


Private Const Text_Max% = 4             '��ʍ��ڕʍő���ޯ��

Dim SYUKA As New XArrayDB

Private Const Min_Row% = 1              '�ŏ��s��
Dim Max_Row    As Integer               '�O���b�h�ő�\������

Dim SYUKA_DATA  As String               '�o�׃f�[�^�t���p�X


Private Const Min_Col% = 0              '�ŏ���
Private Const Max_Col% = 12             '�ő��

Private Const ColSYUKA_YMD% = 0         '�o�ח\���
Private Const ColID_NO% = 1             'ID��
Private Const ColDEN_NO% = 2            '�`�[��
Private Const ColSYUKO_SYUSI& = 3       '�o�Ɏ��x
Private Const ColHIN_GAI% = 4           '�i�ԁi�O���j
Private Const ColHIN_NAME% = 5          '�i��
Private Const ColYOTEI_QTY% = 6         '�o�ח\�萔

Private Const ColTANABAN1% = 7          '�I�ԂP
Private Const ColTANABAN2% = 8          '�I�ԂQ
Private Const ColTANABAN3% = 9          '�I�ԂR


Private Const ColKENPIN_Date% = 10      '���i��
Private Const ColKENPIN_TANTO% = 11     '���i�S����

Private Const ColMUKE_CODE% = 12        '������

Private Sort_Tbl(Min_Col To Max_Col) _
                As Integer                  '��Ă̐��� 0:���� 1:�~��
                
Private Const LAST_UPDATE_DAY$ = "2009.03.13 11:20"
                






Private Sub Command_Click(Index As Integer)

Dim ans As Integer

    Select Case Index
        Case 7                              '�ĕ\��
            If List_Disp_Proc Then
                Unload Me
            End If
        Case 8                              '�f�[�^�o��
        
            Beep
            ans = MsgBox("�u�ڊǎ���v�f�[�^�o�͂��܂����H", vbYesNo + vbQuestion, "�m�F����")
            
            If ans = vbYes Then
                If OUTPUT_Proc() Then
                    Unload Me
                End If
            
            
                If List_Disp_Proc Then
                    Unload Me
                End If
            
            
            End If
        
        
        
        
        Case 11                             '�I��
            Unload Me
        Case Else
            Beep
    End Select
    
    Text(ptxSyuka_YY).SetFocus
    
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
            Command(KeyCode - vbKeyF1).Value = True
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

Dim Start_Pos   As Integer


    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If

    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�ڊǎ������", Me.hwnd, 0)
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
                    
                    
                    
                                '�o�̓t�@�C������荞��
    If GetIni("FILE", "OUT_FILE", App.EXEName, c) Then
        Beep
        MsgBox "�o�̓t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    OUT_FILE = RTrim(c)
                    
    OUT_FILE = Replace(OUT_FILE, "mmdd", Mid(Format(Now, "YYMMDD"), 3, 4), , , 1)
                        
                    
                    
                    
                    
                    '�ő�\�������̊l��
    If GetIni(App.EXEName, "LISTMAX", "SYS", c) Then
        Max_Row = 9999
    Else
        Max_Row = CInt(RTrim(c))
    End If
                                '���ƕ���荞��
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

    ReDim Preserve JGYOBU_T(UBound(JGYOBU_T) + 1)
    JGYOBU_T(UBound(JGYOBU_T)).CODE = "*"
    JGYOBU_T(UBound(JGYOBU_T)).NAME = "�SBU"
    JGYOBU_T(UBound(JGYOBU_T)).COLOR = 12


    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F9000901.Caption = "����ڊǊm�F�i" + RTrim(JGYOBU_T(i).NAME) + ")" & " " & LAST_UPDATE_DAY
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)

                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�o�ח\��n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If

'�o�ד��t
    Text(ptxSyuka_YY).Text = Left(Format(Now, "YYYYMMDD"), 4)
    Text(ptxSyuka_MM).Text = Mid(Format(Now, "YYYYMMDD"), 5, 2)
    Text(ptxSyuka_DD).Text = Right(Format(Now, "YYYYMMDD"), 2)

    
        

    Text(ptxSyuka_YY).SetFocus
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
                                            '�S���҃}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�S���҃}�X�^")
        End If
    End If
                                            '�o�ח\��b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��")
        End If
    End If
    
    sts = BTRV(BtOpReset, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    '���j���[���I���v��
    If JGYOBU_T(Index).CODE = " " Then
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F9000901.Caption = "�ڊǎ���m�F�i" + RTrim(JGYOBU_T(Index).NAME) + ")" & " " & LAST_UPDATE_DAY

    SubMenu(Index).Checked = True
    If Last_JGYOBU <> JGYOBU_T(Index).CODE Then
        Last_JGYOBU = JGYOBU_T(Index).CODE
        LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
        LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

    End If

End Sub
Private Function List_Disp_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim Row         As Long
    
Dim DEN_MAISU   As Long
Dim KAN_MAISU   As Long
    
Dim Skip_Flg    As Boolean
    
    
    List_Disp_Proc = True
                                    
    
    
    Call Input_Lock
                                    
                                    
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�ڊǏ��@������", Me.hwnd, 0)
                                    
                                    
    Command(8).Enabled = False
                                    
                                    
                                    
                                    '�e�[�u�����Z�b�g
    Set SYUKA = Nothing
    '��ď��̏�����
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0             '��̫�ď���
    Next i
                                    '�o�ח\��ǂݍ��݊J�n
    
    If Last_JGYOBU = "*" Then
        Call UniCode_Conv(K2_Y_SYU.JGYOBU, "") '���ƕ�
    Else
        Call UniCode_Conv(K2_Y_SYU.JGYOBU, Last_JGYOBU) '���ƕ�
    End If
                                                    '�����敪
    Call UniCode_Conv(K2_Y_SYU.KEY_CYU_KBN, "")
                                                    '������
    Call UniCode_Conv(K2_Y_SYU.KEY_MUKE_CODE, "")
    Call UniCode_Conv(K2_Y_SYU.KEY_SS_CODE, "")
    
    
    Row = Min_Row - 1
        
    DEN_MAISU = 0
    KAN_MAISU = 0
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        
        
        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K2_Y_SYU, Len(K2_Y_SYU), 2)
    
    
    
        Select Case sts
            Case BtNoErr
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�o�ח\��")
                List_Disp_Proc = SYS_ERR
                Call Input_UnLock
                Exit Function
        End Select
                                '���ƕ� KEY��ڰ�
        
        If Last_JGYOBU = "*" Then
        Else
            If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                Exit Do
            End If
        End If
                                
        Skip_Flg = False
                                
        
        If Len(Trim((Text(ptxSyuka_YY).Text & Text(ptxSyuka_MM).Text & Text(ptxSyuka_DD).Text))) = 0 Then
        
        
        
        Else
            
            
            
            If (Text(ptxSyuka_YY).Text & Text(ptxSyuka_MM).Text & Text(ptxSyuka_DD).Text) <> StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) Then
                Skip_Flg = True
            End If
        End If
                
                
        If Not Skip_Flg Then
            
            
            DEN_MAISU = DEN_MAISU + 1
            
            
            If Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode)) = "" Then
            Else
                KAN_MAISU = KAN_MAISU + 1
            End If
            
            Row = Row + 1
            If Row > Max_Row Then
                Beep
                MsgBox "�ő�\���s���𒴂��܂����B"
                Exit Do
            End If
                    
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
        End If
        
        com = BtOpGetNext
        
    Loop
    
                                
                                'DB�e�[�u�������N
    
    Set TDBGrid1.Array = SYUKA
    
    TDBGrid1.style.Locked = True
    
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    
    
    Text(ptxKAN_MAISU).Text = KAN_MAISU
    
    Text(ptxDEN_MAISU).Text = DEN_MAISU
    
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�ڊǏ��@�����I��", Me.hwnd, 0)
    
    
    
    
    
    Call Input_UnLock
    
    If DEN_MAISU > 0 Then
        Command(8).Enabled = True
    End If
    
    List_Disp_Proc = False

    
End Function

Private Function OUTPUT_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer
Dim i           As Integer
Dim j           As Integer
    
    
Dim Ret         As Integer
    

Dim FileNo      As Integer
Dim fileName    As String
    
Dim Skip_Flg    As Boolean
    
    
Dim hinban      As String * 20
Dim tanaban     As String * 10
Dim KENPIN_DATE As String * 19
    
    
    
    OUTPUT_Proc = True
    
    Call Input_Lock
                                    
                                    

    
    
    
    Set TDBGrid1.Array = SYUKA
    TDBGrid1.Refresh
    
    TDBGrid1.Update
                                     
    If SYUKA.Count(1) < 1 Then
        OUTPUT_Proc = False
        Call Input_UnLock
        Exit Function
    End If
    
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�ڊǏ��@�o�͒�", Me.hwnd, 0)
                                    
                                    


    FileNo = FreeFile
    
    fileName = OUT_FILE
    
    On Error GoTo Error_Proc
    
    Open (OUT_FILE) For Output As FileNo
    
    
    
    
    
    For i = 1 To SYUKA.Count(1)
    
    
        If Trim(SYUKA(i, ColKENPIN_Date)) = "" Then
            Debug.Print
        Else
        
        
        
            If Trim(SYUKA(i, ColTANABAN3)) <> "" Then
                        
                hinban = SYUKA(i, ColHIN_GAI)
                Print #FileNo, hinban;          '�i��
                tanaban = SYUKA(i, ColTANABAN1)
                Print #FileNo, tanaban;         '�I�ԂP
                Print #FileNo, "I";             '�����敪
                                    
                                                '���i���t
                KENPIN_DATE = SYUKA(i, ColKENPIN_Date)
                Print #FileNo, Mid(KENPIN_DATE, 1, 4) & _
                                Mid(KENPIN_DATE, 6, 2) & _
                                Mid(KENPIN_DATE, 9, 2) & _
                                Mid(KENPIN_DATE, 12, 2) & _
                                Mid(KENPIN_DATE, 15, 2) & _
                                Mid(KENPIN_DATE, 18, 2)
        
        
                Print #FileNo, hinban;          '�i��
                tanaban = SYUKA(i, ColTANABAN2)
                Print #FileNo, tanaban;         '�I�ԂQ
                Print #FileNo, "I";             '�����敪
                                    
                                                '���i���t
                KENPIN_DATE = SYUKA(i, ColKENPIN_Date)
                Print #FileNo, Mid(KENPIN_DATE, 1, 4) & _
                                Mid(KENPIN_DATE, 6, 2) & _
                                Mid(KENPIN_DATE, 9, 2) & _
                                Mid(KENPIN_DATE, 12, 2) & _
                                Mid(KENPIN_DATE, 15, 2) & _
                                Mid(KENPIN_DATE, 18, 2)
        
        
                Print #FileNo, hinban;          '�i��
                tanaban = SYUKA(i, ColTANABAN3)
                Print #FileNo, tanaban;         '�I�ԂR
                Print #FileNo, "I";             '�����敪
                                    
                                                '���i���t
                KENPIN_DATE = SYUKA(i, ColKENPIN_Date)
                Print #FileNo, Mid(KENPIN_DATE, 1, 4) & _
                                Mid(KENPIN_DATE, 6, 2) & _
                                Mid(KENPIN_DATE, 9, 2) & _
                                Mid(KENPIN_DATE, 12, 2) & _
                                Mid(KENPIN_DATE, 15, 2) & _
                                Mid(KENPIN_DATE, 18, 2)
        
        
            Else
        
                If Trim(SYUKA(i, ColTANABAN2)) <> "" Then
                            
                    hinban = SYUKA(i, ColHIN_GAI)
                    Print #FileNo, hinban;          '�i��
                    tanaban = SYUKA(i, ColTANABAN1)
                    Print #FileNo, tanaban;         '�I�ԂP
                    Print #FileNo, "I";             '�����敪
                                        
                                                    '���i���t
                    KENPIN_DATE = SYUKA(i, ColKENPIN_Date)
                    Print #FileNo, Mid(KENPIN_DATE, 1, 4) & _
                                    Mid(KENPIN_DATE, 6, 2) & _
                                    Mid(KENPIN_DATE, 9, 2) & _
                                    Mid(KENPIN_DATE, 12, 2) & _
                                    Mid(KENPIN_DATE, 15, 2) & _
                                    Mid(KENPIN_DATE, 18, 2)
            
            
                    Print #FileNo, hinban;          '�i��
                    tanaban = SYUKA(i, ColTANABAN2)
                    Print #FileNo, tanaban;         '�I�ԂQ
                    Print #FileNo, "I";             '�����敪
                                        
                                                    '���i���t
                    KENPIN_DATE = SYUKA(i, ColKENPIN_Date)
                    Print #FileNo, Mid(KENPIN_DATE, 1, 4) & _
                                    Mid(KENPIN_DATE, 6, 2) & _
                                    Mid(KENPIN_DATE, 9, 2) & _
                                    Mid(KENPIN_DATE, 12, 2) & _
                                    Mid(KENPIN_DATE, 15, 2) & _
                                    Mid(KENPIN_DATE, 18, 2)
            
            
            
            
                Else
            
                    hinban = SYUKA(i, ColHIN_GAI)
                    Print #FileNo, hinban;          '�i��
                    tanaban = SYUKA(i, ColTANABAN1)
                    Print #FileNo, tanaban;         '�I�ԂP
                    Print #FileNo, "I";             '�����敪
                                        
                                                    '���i���t
                    KENPIN_DATE = SYUKA(i, ColKENPIN_Date)
                    Print #FileNo, Mid(KENPIN_DATE, 1, 4) & _
                                    Mid(KENPIN_DATE, 6, 2) & _
                                    Mid(KENPIN_DATE, 9, 2) & _
                                    Mid(KENPIN_DATE, 12, 2) & _
                                    Mid(KENPIN_DATE, 15, 2) & _
                                    Mid(KENPIN_DATE, 18, 2)
            
                End If
            End If
        End If
    
    
    
    
    
    Next i



    Close #FileNo
    
    Call Input_UnLock
    
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�ڊǏ��@�o�͏I��", Me.hwnd, 0)
    
    
    Beep
    MsgBox "�u" & fileName & "�v�͐���ɏo�͂���܂����B"

    
    OUTPUT_Proc = False
    
    Exit Function

Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "���g�p���ł��B"
        OUTPUT_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        OUTPUT_Proc = True
    End If

    Call Input_UnLock

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F9000901.MousePointer = vbHourglass

    Call Ctrl_Lock(F9000901)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F9000901)


    F9000901.MousePointer = vbDefault

End Sub

Private Function Grid_Set_Proc(Row As Long) As Integer

Dim sts As Integer

    
    Grid_Set_Proc = True

    

    SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
    
    
    
                                                                            '�o�ח\���
    SYUKA(Row, ColSYUKA_YMD) = Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 1, 4) & "/" _
                                & Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 5, 2) & "/" _
                                & Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 7, 4)
    
    
    SYUKA(Row, ColID_NO) = StrConv(Y_SYUREC.ID_NO, vbUnicode)               '�h�c��
    SYUKA(Row, ColDEN_NO) = StrConv(Y_SYUREC.DEN_NO, vbUnicode)             '�`�[��
    SYUKA(Row, ColSYUKO_SYUSI) = StrConv(Y_SYUREC.SYUKO_SYUSI, vbUnicode)   '�o�Ɏ��x
    SYUKA(Row, ColHIN_GAI) = StrConv(Y_SYUREC.HIN_NO, vbUnicode)            '�i�ԁi�O���j
                                                                    '�i�ڃ}�X�^�Ǎ���
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            SYUKA(Row, ColHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    End Select
                                                                    
                                                                    
                                                                    '�I�ԂP
    SYUKA(Row, ColTANABAN1) = Trim(StrConv(Y_SYUREC.TANABAN1, vbUnicode))
                                                                    '�I�ԂQ
    SYUKA(Row, ColTANABAN2) = Trim(StrConv(Y_SYUREC.TANABAN2, vbUnicode))
                                                                    '�I�ԂR
    SYUKA(Row, ColTANABAN3) = Trim(StrConv(Y_SYUREC.TANABAN3, vbUnicode))
                                                                    
                                                                    
                                                                    
                                                                    '�o�ח\�萔
    SYUKA(Row, ColYOTEI_QTY) = Format(CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)), "#,##0")
    
                                                                    '���i����
    If Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode)) <> "" Then
        SYUKA(Row, ColKENPIN_Date) = Mid(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode), 1, 4) & "/" _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode), 5, 2) & "/" _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode), 7, 2) & " " _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_HMS, vbUnicode), 1, 2) & ":" _
                                    & Mid(StrConv(Y_SYUREC.KENPIN_HMS, vbUnicode), 3, 2)

    Else
        SYUKA(Row, ColKENPIN_Date) = ""
    End If
    
    
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(Y_SYUREC.KENPIN_TANTO_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(TANTOREC.TANTO_NAME, "")
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
            Exit Function
    End Select
    
    
    SYUKA(Row, ColKENPIN_TANTO) = StrConv(Y_SYUREC.KENPIN_TANTO_CODE, vbUnicode) & " " & StrConv(TANTOREC.TANTO_NAME, vbUnicode)
    
    SYUKA(Row, ColMUKE_CODE) = StrConv(Y_SYUREC.SS_CODE, vbUnicode)
    
    
    
    Grid_Set_Proc = False
End Function


Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    If Sort_Tbl(ColIndex) = 0 Then
        Sort_Tbl(ColIndex) = 1
    Else
        If Sort_Tbl(ColIndex) = 1 Then
            Sort_Tbl(ColIndex) = 0
        End If
    
    End If
    
    If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then
                    
        SYUKA.QuickSort Min_Row, SYUKA.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
        
        Set TDBGrid1.Array = SYUKA
        
        TDBGrid1.ReBind
        TDBGrid1.Update
        TDBGrid1.MoveFirst


    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If


End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim sts As Integer
Dim i   As Integer

    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    Select Case Index
        
        Case ptxSyuka_YY
            If Len(Trim(Text(ptxSyuka_YY).Text)) = 0 Then
            Else
            
                If Not IsNumeric(Text(ptxSyuka_YY).Text) Then
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Exit Sub
                End If
            End If
        Case ptxSyuka_MM
            If Len(Trim(Text(ptxSyuka_MM).Text)) = 0 Then
            Else
                If Not IsNumeric(Text(ptxSyuka_MM).Text) Then
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Exit Sub
                End If
                Text(ptxSyuka_MM).Text = Format(CInt(Text(ptxSyuka_MM).Text), "00")
            End If
        Case ptxSyuka_DD
            If Len(Trim(Text(ptxSyuka_DD).Text)) = 0 Then
            Else
                If Not IsNumeric(Text(ptxSyuka_DD).Text) Then
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Exit Sub
                End If
                Text(ptxSyuka_DD).Text = Format(CInt(Text(ptxSyuka_DD).Text), "00")
            End If
        
        


    End Select
    
    For i = Index + 1 To Text_Max
        If Text(i).Visible And Text(i).Enabled And Text(i).TabStop And Not Text(i).Locked Then
            Text(i).SetFocus
            Exit For
        End If
    Next i

End Sub

