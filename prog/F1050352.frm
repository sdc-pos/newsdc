VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form F1050352 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�o�͗v���w��^�m�F"
   ClientHeight    =   4470
   ClientLeft      =   30
   ClientTop       =   3300
   ClientWidth     =   8385
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton Command1 
      Caption         =   "�I��/���I��"
      Height          =   495
      Left            =   6360
      TabIndex        =   30
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text 
      BackColor       =   &H8000000B&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   14
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox Text 
      BackColor       =   &H8000000B&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   4680
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   240
      Width           =   3135
   End
   Begin VB.ComboBox Combo 
      BackColor       =   &H8000000B&
      Height          =   336
      Index           =   0
      Left            =   1200
      Locked          =   -1  'True
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text 
      BackColor       =   &H8000000B&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text 
      BackColor       =   &H8000000B&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text 
      BackColor       =   &H8000000B&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text 
      BackColor       =   &H8000000B&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text 
      BackColor       =   &H8000000B&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   3720
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text 
      BackColor       =   &H8000000B&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   4200
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command 
      Caption         =   "��ݾ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   11
      Left            =   6240
      TabIndex        =   11
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   10
      Left            =   5760
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3960
      Width           =   492
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   9
      Left            =   5280
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3960
      Width           =   492
   End
   Begin VB.CommandButton Command 
      Caption         =   "�S��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   8
      Left            =   4440
      TabIndex        =   8
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   7
      Left            =   3840
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3960
      Width           =   492
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   6
      Left            =   3360
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3960
      Width           =   492
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   5
      Left            =   2880
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3960
      Width           =   492
   End
   Begin VB.CommandButton Command 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   4
      Left            =   2400
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3960
      Width           =   492
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   3
      Left            =   1800
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3960
      Width           =   492
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3960
      Width           =   492
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   840
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3960
      Width           =   492
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3960
      Width           =   492
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   1815
      Left            =   1680
      TabIndex        =   31
      Top             =   1320
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3201
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "�v��"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "����"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   4
      Columns(2)._MaxComboItems=   5
      Columns(2).ValueItems(0)._DefaultItem=   0
      Columns(2).ValueItems(0).Value=   ""
      Columns(2).ValueItems(0).Value.vt=   8
      Columns(2).ValueItems(0).DisplayValue=   "1"
      Columns(2).ValueItems(0).DisplayValue.vt=   8
      Columns(2).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems.Count=   1
      Columns(2).Caption=   "�I"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   714
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1085"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=953"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=4366"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4233"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=635"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=503"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �S�V�b�N"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �S�V�b�N"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=29,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=�l�r �S�V�b�N"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=33"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=30,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=�l�r �S�V�b�N"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=31,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=�l�r �S�V�b�N"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=32"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=34"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=35"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=36"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=41,.parent=2,.namedParent=43"
      _StyleDefs(23)  =   "FilterBarStyle:id=44,.parent=1,.namedParent=46"
      _StyleDefs(24)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=12,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=16,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=15,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=18,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=19,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=42,.parent=41"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=45,.parent=44"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=40,.parent=11,.alignment=3,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=37,.parent=12,.alignment=3"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=38,.parent=13,.alignment=3"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=39,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=24,.parent=11,.alignment=3,.locked=-1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=21,.parent=12,.alignment=3"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=22,.parent=13,.alignment=3"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=23,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=11,.alignment=2,.locked=0"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=12,.alignment=3"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=13,.alignment=3"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=15"
      _StyleDefs(48)  =   "Named:id=29:Normal"
      _StyleDefs(49)  =   ":id=29,.parent=0"
      _StyleDefs(50)  =   "Named:id=30:Heading"
      _StyleDefs(51)  =   ":id=30,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(52)  =   ":id=30,.wraptext=-1"
      _StyleDefs(53)  =   "Named:id=31:Footing"
      _StyleDefs(54)  =   ":id=31,.parent=29,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(55)  =   "Named:id=32:Selected"
      _StyleDefs(56)  =   ":id=32,.parent=29,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=33:Caption"
      _StyleDefs(58)  =   ":id=33,.parent=30,.alignment=2"
      _StyleDefs(59)  =   "Named:id=34:HighlightRow"
      _StyleDefs(60)  =   ":id=34,.parent=29,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(61)  =   "Named:id=35:EvenRow"
      _StyleDefs(62)  =   ":id=35,.parent=29,.bgcolor=&HFFFF00&"
      _StyleDefs(63)  =   "Named:id=36:OddRow"
      _StyleDefs(64)  =   ":id=36,.parent=29"
      _StyleDefs(65)  =   "Named:id=43:RecordSelector"
      _StyleDefs(66)  =   ":id=43,.parent=30"
      _StyleDefs(67)  =   "Named:id=46:FilterBar"
      _StyleDefs(68)  =   ":id=46,.parent=29"
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "�����яƉ��ʂɕ\�������v���̂ݏo�͑ΏۂƂȂ�܂��B     ����ȊO�̎w��͖����ł��B"
      Height          =   492
      Left            =   840
      TabIndex        =   29
      Top             =   3360
      Width           =   6732
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i��"
      Height          =   252
      Index           =   0
      Left            =   2160
      TabIndex        =   28
      Top             =   360
      Width           =   612
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����O"
      Height          =   252
      Index           =   33
      Left            =   240
      TabIndex        =   27
      Top             =   360
      Width           =   852
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���t"
      Height          =   252
      Index           =   1
      Left            =   480
      TabIndex        =   26
      Top             =   840
      Width           =   492
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   252
      Index           =   2
      Left            =   1800
      TabIndex        =   25
      Top             =   840
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   252
      Index           =   3
      Left            =   2280
      TabIndex        =   24
      Top             =   840
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`"
      Height          =   252
      Index           =   4
      Left            =   2760
      TabIndex        =   23
      Top             =   840
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   252
      Index           =   5
      Left            =   3600
      TabIndex        =   22
      Top             =   840
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   252
      Index           =   6
      Left            =   4080
      TabIndex        =   21
      Top             =   840
      Width           =   252
   End
End
Attribute VB_Name = "F1050352"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------
'                                   '�f�[�^�o�͋��ʒ�`
Private Const SYS_INI = "SYS"
Private Const P_ID = "F105035"      '�v���O�����h�c�D
Dim GW_Path As String

Private Const F105035_INI = "F105035"   '2016.01.15



Private Const pcmbNAIGAI% = 0           '�����O

Private Const ptxHin_Gai% = 0           '�i�ԁi�O���j
Private Const ptxHin_Name% = 1          '�i��
Private Const ptxST_DT_YY% = 2          '�J�n���t �N
Private Const ptxST_DT_MM% = 3          '�J�n���t ��
Private Const ptxST_DT_DD% = 4          '�J�n���t ��
Private Const ptxEN_DT_YY% = 5          '�I�����t �N
Private Const ptxEN_DT_MM% = 6          '�I�����t ��
Private Const ptxEN_DT_DD% = 7          '�I�����t ��


Dim YOIN_SEL     As New XArrayDB

Private Const Min_Row% = 1              '�ŏ��s��
'Private Const Max_Row& = 500            '�ő�s��

Private Const Min_Col% = 0              '�ŏ���
Private Const Max_Col% = 2              '�ő��

Private Const ColYOIN_CODE% = 0         '�v������
Private Const ColYOIN_Name% = 1         '�v������
Private Const ColSELECT% = 2            '�v������

Dim Chk_Flg As Integer



Private Sub Command_Click(Index As Integer)
            
Dim i As Integer
            

    Select Case Index
        Case 4                          '���̂�
        
            TDBGrid1.Update
            
'            If SDC_FLD_GET(SYS_INI, P_ID, GW_Path) Then
            If SDC_FLD_GET(F105035_INI, P_ID, GW_Path) Then
                F1050352.Hide
            End If
            
            If SDC_FLD_Return Then
            Else
                If Data_Out(1) Then
                    Unload Me
                End If
                F1050352.Hide
            End If
        
        Case 7                          '�ς݂̂�
        
            TDBGrid1.Update
            
'            If SDC_FLD_GET(SYS_INI, P_ID, GW_Path) Then
            If SDC_FLD_GET(F105035_INI, P_ID, GW_Path) Then
                F1050352.Hide
            End If
            
            If SDC_FLD_Return Then
            Else
                If Data_Out(2) Then
                    Unload Me
                End If
                F1050352.Hide
            End If
        
        
        
        Case 8                          '�S��
            TDBGrid1.Update
            
'            If SDC_FLD_GET(SYS_INI, P_ID, GW_Path) Then
            If SDC_FLD_GET(F105035_INI, P_ID, GW_Path) Then
                F1050352.Hide
            End If
            
            If SDC_FLD_Return Then
            Else
                If Data_Out(0) Then
                    Unload Me
                End If
                F1050352.Hide
            End If
        Case 11
            F1050352.Hide
    End Select

End Sub

Private Sub Command1_Click()
                                    
Dim Row As Integer
Dim com As Integer
Dim sts As Integer
                                    '�e�[�u�����Z�b�g
    Set YOIN_SEL = Nothing
    
    If Not Chk_Flg Then
        Chk_Flg = True
    Else
        Chk_Flg = False
    End If
    
    Row = Min_Row - 1
        
    com = BtOpGetFirst
    Do
        sts = BTRV(com, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    
        Select Case sts
            Case BtNoErr
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�v���}�X�^")
                Unload Me
        End Select
    
        Row = Row + 1
'�i�荞�݂悤�������̂őS���\��
'        If Row > Max_Row Then
'            Beep
'            MsgBox "�ő�\���s���𒴂��܂����B"
'            Exit Do
'        End If
    
    
        YOIN_SEL.ReDim Min_Row, Row, Min_Col, Max_Col
                                            '�v������
        YOIN_SEL(Row, ColYOIN_CODE) = StrConv(YOINREC.CODE_TYPE, vbUnicode) & StrConv(YOINREC.YOIN_CODE, vbUnicode)
                                            '�v������
        YOIN_SEL(Row, ColYOIN_Name) = StrConv(YOINREC.YOIN_DNAME, vbUnicode)
                                            
        
        
        YOIN_SEL(Row, ColSELECT) = Chk_Flg    '�I��
    
        com = BtOpGetNext
        DoEvents
    Loop
                                'DB�e�[�u�������N
    Set TDBGrid1.Array = YOIN_SEL
    
'    TDBGrid1.Style.Locked = True
    TDBGrid1.ReBind
    
    



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

Private Sub Form_Load()

    If Yoin_Set_Proc() Then     '�I����ʂɗv���ݒ�
        Unload Me
    End If

    Combo(pcmbNAIGAI).AddItem NAIGAI1
    Combo(pcmbNAIGAI).AddItem NAIGAI2
    Combo(pcmbNAIGAI).AddItem NAIGAI0


End Sub

Private Function Yoin_Set_Proc() As Integer
Dim sts         As Integer
Dim com         As Integer
Dim Row         As Long
    
    Yoin_Set_Proc = True
                                    '�e�[�u�����Z�b�g
    Set YOIN_SEL = Nothing
    
    
    Row = Min_Row - 1
        
    com = BtOpGetFirst
    Do
        sts = BTRV(com, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    
        Select Case sts
            Case BtNoErr
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�v���}�X�^")
                Exit Function
        End Select
    
        Row = Row + 1
'�i�荞�݂悤�������̂őS���\��
'        If Row > Max_Row Then
'            Beep
'            MsgBox "�ő�\���s���𒴂��܂����B"
'            Exit Do
'        End If
    
    
        YOIN_SEL.ReDim Min_Row, Row, Min_Col, Max_Col
                                            '�v������
        YOIN_SEL(Row, ColYOIN_CODE) = StrConv(YOINREC.CODE_TYPE, vbUnicode) & StrConv(YOINREC.YOIN_CODE, vbUnicode)
                                            '�v������
        YOIN_SEL(Row, ColYOIN_Name) = StrConv(YOINREC.YOIN_DNAME, vbUnicode)
                                            
        YOIN_SEL(Row, ColSELECT) = False    '�I��
    
        com = BtOpGetNext
        DoEvents
    Loop
                                'DB�e�[�u�������N
    Set TDBGrid1.Array = YOIN_SEL
    
'    TDBGrid1.Style.Locked = True
    TDBGrid1.ReBind
    
    Chk_Flg = False
    
    
    Yoin_Set_Proc = False


End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------
Dim i As Integer

    F1050352.MousePointer = vbHourglass

    Call Ctrl_Lock(F1050352)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1050352)


    F1050352.MousePointer = vbDefault

End Sub

Private Function Data_Out(Mode As Integer)

Dim sts         As Integer
Dim com         As Integer
Dim Row         As Integer
Dim i           As Integer
Dim NAIGAI      As String * 1

Dim Put_Flg     As Boolean

Dim Yoin_Tbl()  As String * 2

Dim FileNo      As Integer

Dim c           As String * 128
Dim Soko_No     As String * 2


Dim Key_Mode    As Integer  '2016.02.04
Dim SKIP_Flg    As Boolean  '2016.02.04

    Data_Out = True

    Call Input_Lock
                                '�Ώۗv���ݒ�
    i = 0
    For Row = Min_Row To YOIN_SEL.UpperBound(1)
        If YOIN_SEL(Row, ColSELECT) Then
            ReDim Preserve Yoin_Tbl(i)
            Yoin_Tbl(i) = YOIN_SEL(Row, ColYOIN_CODE)
            i = i + 1
        End If
    Next Row
                                    
    If i = 0 Then            '�I���Ȃ�
        Call Input_UnLock
        Data_Out = False
        Exit Function
    End If
    
    
    Select Case Combo(pcmbNAIGAI).Text
        Case NAIGAI0
            NAIGAI = NAIGAI_NON
        Case NAIGAI1
            NAIGAI = NAIGAI_NAI
        Case NAIGAI2
            NAIGAI = NAIGAI_GAI
    End Select
                                    
                                    
    '�e�L�X�g�t�@�C���n�o�d�m
    FileNo = FreeFile
    
    On Error GoTo Error_Proc
    
    Open Trim(GW_Path) For Output As FileNo
                                '�w�b�_�o��
    Write #FileNo, "���ƕ�", "�����O", "�i��(�O��)", "�i��", "�v��", "��", "���ѓ��t", "���ю���", "�`�[���t", "�`�[��", _
                    "���݌ɐ�", "���ɐ�(��)", "���ɐ�(��)", "�o�ɐ�(��)", "�o�ɐ�(��)", "�ݒ�(��)", "�ݒ�(��)", "�ړ�(��)", _
                        "�ړ�(��)", "���o�Ɍ�", "���o�ɐ�", "���ד�", "������", "ID", "�i��(����)", "����", "�`�[�h�c", _
                            "�S����CD", "�S���Җ�"
                                    
                                    
'>>>>>>>>>>>>>>>>>>>    KEY�Z�b�g�ύX   2016.02.04
                                    '�݌Ɉړ���ǂݍ��݊J�n
'    Call UniCode_Conv(K0_IDO.JGYOBU, Last_JGYOBU)
'    Call UniCode_Conv(K0_IDO.JITU_DT, Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text)
'    Call UniCode_Conv(K0_IDO.JITU_TM, "")
    
    
    
    Select Case Combo(pcmbNAIGAI).Text
        Case NAIGAI1                '����
            NAIGAI = NAIGAI_NAI
            If Len(Trim(Text(ptxHin_Gai).Text)) = 0 Then
                Key_Mode = 0
            Else
                Key_Mode = 1
            End If
        Case NAIGAI2                '�C�O
            NAIGAI = NAIGAI_GAI
            If Len(Trim(Text(ptxHin_Gai).Text)) = 0 Then
                Key_Mode = 0
            Else
                Key_Mode = 1
            End If
        Case NAIGAI0                '���O�w��Ȃ�
            Key_Mode = 0
    End Select
                                    
                                    
                                    '�݌Ɉړ���ǂݍ��݊J�n
    If Key_Mode = 0 Then
                                    '���n��œǂݍ���
        Call UniCode_Conv(K0_IDO.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_IDO.JITU_DT, Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text)
        If Mode = 0 Then
            Call UniCode_Conv(K0_IDO.JITU_DT, Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text)
            Call UniCode_Conv(K0_IDO.JITU_TM, "")           '����
        Else
            Call UniCode_Conv(K0_IDO.JITU_DT, Text(ptxEN_DT_YY).Text & Text(ptxEN_DT_MM).Text & Text(ptxEN_DT_DD).Text)
            Call UniCode_Conv(K0_IDO.JITU_TM, "zzzzzzzz")   '�~��
        End If
    
    Else
                                    '�i�ԁ����n��œǂݍ���
        Call UniCode_Conv(K1_IDO.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K1_IDO.NAIGAI, NAIGAI)
        Call UniCode_Conv(K1_IDO.HIN_GAI, Text(ptxHin_Gai).Text)
        If Mode = 0 Then
            Call UniCode_Conv(K1_IDO.JITU_DT, Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text)
            Call UniCode_Conv(K1_IDO.JITU_TM, "")           '����
        Else
            Call UniCode_Conv(K1_IDO.JITU_DT, Text(ptxEN_DT_YY).Text & Text(ptxEN_DT_MM).Text & Text(ptxEN_DT_DD).Text)
            Call UniCode_Conv(K1_IDO.JITU_TM, "zzzzzzzz")   '�~��
        End If
    End If
'>>>>>>>>>>>>>>>>>>>    KEY�Z�b�g�ύX   2016.02.04
    
    
    
    
    com = BtOpGetGreater
    
    Do
'>>>>>>>>>>>>>>>>>>>    KEY�Z�b�g�ύX   2016.02.04
'        sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
        
        If Key_Mode = 0 Then
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
        Else
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
        End If
'>>>>>>>>>>>>>>>>>>>    KEY�Z�b�g�ύX   2016.02.04
    
        Select Case sts
            Case BtNoErr
        
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�݌Ɉړ���")
                Data_Out = SYS_ERR
                Exit Function
        End Select
'>>>>>>>>>>>>>>>>>>>    KEY�Z�b�g�ύX   2016.02.04
        If Not SKIP_Flg Then
                                    
                                    '���ƕ� KEY��ڰ�
            If StrConv(IDOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                Exit Do
            End If
                                    '���t�͈͊O
            If StrConv(IDOREC.JITU_DT, vbUnicode) > (Text(ptxEN_DT_YY).Text & Text(ptxEN_DT_MM).Text & Text(ptxEN_DT_DD).Text) Then
                Exit Do
            End If
            
            If Key_Mode = 1 Then
                                    '�i�Ԏw�莞�A�i����ڰ�
                If StrConv(IDOREC.NAIGAI, vbUnicode) <> NAIGAI Or _
                    Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(Text(ptxHin_Gai).Text) Then
                    Exit Do
                End If
            End If
                                        
        End If
                                
        Put_Flg = True
                            
        If Key_Mode = 1 Then
                                '�i�Ԏw�莞�A�i����ڰ�
            If StrConv(IDOREC.NAIGAI, vbUnicode) <> NAIGAI Or _
                Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(Text(ptxHin_Gai).Text) Then
                Exit Do
            End If
        End If
        
        
        If Key_Mode = 0 Then
            If StrConv(IDOREC.NAIGAI, vbUnicode) <> NAIGAI Then
                Put_Flg = False
            Else
            End If
        End If
                                
'        If Put_Flg Then
'            If Len(Trim(Text(ptxHin_Gai).Text)) = 0 Then
'            Else
'                If RTrim(Text(ptxHin_Gai).Text) <> RTrim(StrConv(IDOREC.HIN_GAI, vbUnicode)) Then
'                                '�i�ԃu���[�N
'                    Put_Flg = False
'                Else
'                    Debug.Print
'                End If
'            End If
'        End If
'
'
'        If Put_Flg Then
'
'            Select Case Mode
'                Case 0          '�S���Ώ�
'                Case 1          '�����i�̂�
'
'                    If CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)) = 0 Then
'                        Put_Flg = False
'                    End If
'
'
'                Case 2          '���i���̂�
'
'
'                    If CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) = 0 Then
'                        Put_Flg = False
'                    End If
'
'
'            End Select
'
'        End If
'>>>>>>>>>>>>>>>>>>>    KEY�Z�b�g�ύX   2016.02.04
                                
        If Put_Flg Then
            If Len(Trim(Text(ptxHin_Gai).Text)) = 0 Then
            Else
                If Trim(Text(ptxHin_Gai).Text) <> Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) Then
                                '�i�ԃu���[�N)
                    Put_Flg = False
                End If
            End If
        End If
        
        
        If Put_Flg Then
        
            Select Case Mode
                Case 0          '�S���Ώ�
                Case 1          '�����i�̂�
                    
                    If CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)) = 0 Then
                        Put_Flg = False
                    End If
                
                
                Case 2          '���i���̂�
            
            
                    If CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) = 0 Then
                        Put_Flg = False
                    End If
            
            
            End Select
        
        End If
        
        
        If Put_Flg Then
            Put_Flg = False
            For i = 0 To UBound(Yoin_Tbl)
                If StrConv(IDOREC.RIRK_ID, vbUnicode) = Yoin_Tbl(i) Then
                    Put_Flg = True
                    Exit For
                End If
            Next i
        End If
                
        If Put_Flg Then
                                                                        
            Write #FileNo, StrConv(IDOREC.JGYOBU, vbUnicode),       '���ƕ�
            Select Case StrConv(IDOREC.NAIGAI, vbUnicode)           '�����O
                Case NAIGAI_NAI
                    Write #FileNo, NAIGAI1,
                Case NAIGAI_GAI
                    Write #FileNo, NAIGAI2,
            End Select
            Write #FileNo, Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)),         '�i�ԁi�O���j
            Write #FileNo, Trim(StrConv(IDOREC.HIN_NAME, vbUnicode)),        '�i��
            Write #FileNo, Trim(StrConv(IDOREC.RIRK_NAME, vbUnicode)),       '�v��
            Write #FileNo, Trim(StrConv(IDOREC.TOKU_MARK, vbUnicode)),       '������}�[�N
                                                                        '���ѓ��t
            Write #FileNo, Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 1, 4) & "/" & _
                      Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 5, 2) & "/" & _
                      Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 7, 2),
                                                                        '���ю���
            Write #FileNo, Mid(StrConv(IDOREC.JITU_TM, vbUnicode), 1, 2) & ":" & _
                      Mid(StrConv(IDOREC.JITU_TM, vbUnicode), 3, 2) & ":" & _
                      Mid(StrConv(IDOREC.JITU_TM, vbUnicode), 5, 2),
                                                                        
            If Len(Trim(StrConv(IDOREC.DEN_DT, vbUnicode))) = 0 Then '�`�[���t
                Write #FileNo, ,
            Else
                Write #FileNo, Mid(StrConv(IDOREC.DEN_DT, vbUnicode), 1, 4) & "/" & _
                          Mid(StrConv(IDOREC.DEN_DT, vbUnicode), 5, 2) & "/" & _
                          Mid(StrConv(IDOREC.DEN_DT, vbUnicode), 7, 2),
            End If
                                                            
            Write #FileNo, Trim(StrConv(IDOREC.DEN_NO, vbUnicode)),          '�`�[��
            Write #FileNo, CLng(StrConv(IDOREC.SUMI_HIN_Zaiko_Qty, vbUnicode)) + CLng(StrConv(IDOREC.MI_HIN_Zaiko_Qty, vbUnicode)), '���݌ɐ�
            
            Select Case StrConv(IDOREC.SUM_KBN, vbUnicode)
                Case SUM_KBN_IN                                        '���ɐ�
                    Write #FileNo, Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)), "#0"), Format(CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#0"), , , , , , ,
                                            
                Case SUM_KBN_OT                                        '�o�ɐ�
                    Write #FileNo, , , Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)), "#0"), Format(CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#0"), , , , ,
                
                Case SUM_KBN_ZT
                    If Mid(StrConv(IDOREC.RIRK_ID, vbUnicode), 1, 1) = ACT_ZAITEI_IN Then
                                                                        '�ݒ��i�{�j
                        Write #FileNo, , , , , Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)), "#0"), Format(CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#0"), , ,
                    Else
                                                                        '�ݒ��i�|�j
                        Write #FileNo, , , , , Format((CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) * -1), "#0"), Format((CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)) * -1), "#0"), , ,
                    End If
        
                Case SUM_KBN_MV                                         '�ړ���
                        Write #FileNo, , , , , , , Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)), "#0"), Format(CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#0"),
                Case Else
                        Write #FileNo, , , , , , , , ,
            End Select
                                                                        'FROM
            If Len(Trim(StrConv(IDOREC.FROM_SOKO, vbUnicode))) = 0 Then
                Write #FileNo, ,
            Else
                
                If GetIni("SOKO_NO", StrConv(IDOREC.FROM_SOKO, vbUnicode), "SYS", c) Then
                    Soko_No = StrConv(IDOREC.FROM_SOKO, vbUnicode)
                Else
                    Soko_No = Trim(c)
                End If
                
                
                Write #FileNo, Soko_No & "-" & _
                            StrConv(IDOREC.FROM_RETU, vbUnicode) & "-" & _
                            StrConv(IDOREC.FROM_REN, vbUnicode) & "-" & _
                            StrConv(IDOREC.FROM_DAN, vbUnicode),
            End If
                                                                        'TO
            If Len(Trim(StrConv(IDOREC.TO_SOKO, vbUnicode))) = 0 Then
                Write #FileNo, ,
            Else
                
                If GetIni("SOKO_NO", StrConv(IDOREC.TO_SOKO, vbUnicode), "SYS", c) Then
                    Soko_No = StrConv(IDOREC.TO_SOKO, vbUnicode)
                Else
                    Soko_No = Trim(c)
                End If
                
                
                Write #FileNo, Soko_No & "-" & _
                            StrConv(IDOREC.TO_RETU, vbUnicode) & "-" & _
                            StrConv(IDOREC.TO_REN, vbUnicode) & "-" & _
                            StrConv(IDOREC.TO_DAN, vbUnicode),
            End If
                                                                        '���ד�
            Write #FileNo, Mid(StrConv(IDOREC.NYUKA_DT, vbUnicode), 1, 4) & "/" & _
                      Mid(StrConv(IDOREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & _
                      Mid(StrConv(IDOREC.NYUKA_DT, vbUnicode), 7, 2),
            Write #FileNo, Trim(StrConv(IDOREC.MUKE_DNAME, vbUnicode)),      '������
            Write #FileNo, Trim(StrConv(IDOREC.WEL_ID, vbUnicode)),          'ID
            Write #FileNo, Trim(StrConv(IDOREC.HIN_NAI, vbUnicode)),         '�i�ԁi�����j
            Write #FileNo, Trim(StrConv(IDOREC.MEMO, vbUnicode)),            '����
            Write #FileNo, Trim(StrConv(IDOREC.ID_NO, vbUnicode)),           '�`�[�h�c
            
            Write #FileNo, Trim(StrConv(IDOREC.TANTO_CODE, vbUnicode)),     '�S���҂b�c     '2004.07.16
            Write #FileNo, Trim(StrConv(IDOREC.TANTO_NAME, vbUnicode)),     '�S���Җ���     '2004.07.16
                    
                        
                    
            Write #FileNo,
        
        
        End If
        
        If Mode = 0 Then
            com = BtOpGetNext   '����
        Else
            com = BtOpGetPrev   '�~��
        End If
        
        DoEvents
    Loop
    
    Close #FileNo
    
    Call Input_UnLock
    
    
    Data_Out = False

    Exit Function

Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox GW_Path & "���g�p���ł��B"
        Call Input_UnLock         '��ʍ��ڃ��b�N����
        Data_Out = False
    Else
        MsgBox "Err.Number" & Err.Number
        Data_Out = True
    End If
    
    
    
End Function

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If
                                            '�v���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�v���}�X�^")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1050351 = Nothing
    Set F1050352 = Nothing
    Set SDC_FLD_F = Nothing

    End
End Sub

