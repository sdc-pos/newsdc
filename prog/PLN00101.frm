VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PLN00101 
   Caption         =   "[���i���v��V�X�e��]���ח\��f�[�^�o�^"
   ClientHeight    =   9510
   ClientLeft      =   2025
   ClientTop       =   -5145
   ClientWidth     =   15210
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
   ScaleHeight     =   9510
   ScaleWidth      =   15210
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CheckBox Check1 
      Caption         =   "�i�ڃ}�X�^���o�^����荞��"
      Height          =   255
      Index           =   0
      Left            =   9960
      TabIndex        =   13
      Top             =   1320
      Width           =   3615
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
      Index           =   3
      Left            =   3720
      TabIndex        =   12
      ToolTipText     =   "�������I�����܂�"
      Top             =   0
      Width           =   1380
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   1
      ItemData        =   "PLN00101.frx":0000
      Left            =   5520
      List            =   "PLN00101.frx":000A
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   10
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   375
      Left            =   2280
      OLEDragMode     =   1  '����
      OLEDropMode     =   1  '�蓮
      TabIndex        =   9
      Top             =   1200
      Width           =   6975
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      ItemData        =   "PLN00101.frx":001A
      Left            =   2280
      List            =   "PLN00101.frx":0024
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   7
      Top             =   720
      Width           =   1815
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   7335
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   12938
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "��"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "�ΊO�i��"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "�Γ��i��"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "�i��"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "���ח\���"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "���ח\�萔"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "�d����"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   714
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1005"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=900"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=4974"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=4868"
      Splits(0)._ColumnProps(8)=   "Column(1)._ColStyle=0"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=5159"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=5054"
      Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=0"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=4339"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=4233"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=3281"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=3175"
      Splits(0)._ColumnProps(22)=   "Column(4)._ColStyle=1"
      Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(24)=   "Column(5).Width=2090"
      Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=1984"
      Splits(0)._ColumnProps(27)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(29)=   "Column(6).Width=3281"
      Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=3175"
      Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=67,.parent=1,.bgcolor=&HFFFF00&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=68,.parent=2,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(27)  =   ":id=68,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(28)  =   ":id=68,.fontname=�l�r �S�V�b�N"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=69,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=72,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=71,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=74,.parent=9,.bgcolor=&HFFFF00&"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=75,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=82,.parent=67,.alignment=3"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=68"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=69"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=71"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=94,.parent=67,.alignment=0"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=91,.parent=68"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=92,.parent=69"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=93,.parent=71"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=98,.parent=67,.alignment=0"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=95,.parent=68"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=96,.parent=69"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=97,.parent=71"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=20,.parent=67"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=17,.parent=68"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=18,.parent=69"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=19,.parent=71"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=102,.parent=67,.alignment=2"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=99,.parent=68"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=100,.parent=69"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=101,.parent=71"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=114,.parent=67,.alignment=1"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=111,.parent=68"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=112,.parent=69"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=113,.parent=71"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=16,.parent=67"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=13,.parent=68"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=14,.parent=69"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=15,.parent=71"
      _StyleDefs(66)  =   "Named:id=33:Normal"
      _StyleDefs(67)  =   ":id=33,.parent=0"
      _StyleDefs(68)  =   "Named:id=34:Heading"
      _StyleDefs(69)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(70)  =   ":id=34,.wraptext=-1"
      _StyleDefs(71)  =   "Named:id=35:Footing"
      _StyleDefs(72)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(73)  =   "Named:id=36:Selected"
      _StyleDefs(74)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(75)  =   "Named:id=37:Caption"
      _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(77)  =   "Named:id=38:HighlightRow"
      _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(79)  =   "Named:id=39:EvenRow"
      _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(81)  =   "Named:id=40:OddRow"
      _StyleDefs(82)  =   ":id=40,.parent=33"
      _StyleDefs(83)  =   "Named:id=41:RecordSelector"
      _StyleDefs(84)  =   ":id=41,.parent=34"
      _StyleDefs(85)  =   "Named:id=42:FilterBar"
      _StyleDefs(86)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�� ��"
      Enabled         =   0   'False
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
      Left            =   2040
      TabIndex        =   2
      ToolTipText     =   "�������I�����܂�"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�V �K"
      Enabled         =   0   'False
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
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "���ח\��f�[�^��o�^���܂�"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�� ��"
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
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "�f�[�^�敪"
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   11
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "�t�@�C����"
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "�a�t"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblDisp_Count 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BorderStyle     =   1  '����
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9960
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "�Ǎ�����"
      Height          =   255
      Index           =   1
      Left            =   8880
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "�����I��"
      Begin VB.Menu SHORI 
         Caption         =   "�Ǎ���"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu SHORI 
         Caption         =   "�V�K"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "�ǉ�"
         Index           =   2
      End
      Begin VB.Menu SHORI 
         Caption         =   "�I��"
         Index           =   3
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "PLN00101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim PLN_Y_NYUKA         As New XArrayDB

Private Const Min_Row% = 1              '�ŏ��s��

Dim Max_Row    As Integer               '�O���b�h�ő�\������


Private Const Min_Col% = 0              '�ŏ���
Private Const Max_Col% = 7              '�ő��

Private Const colNO% = 0                '��
Private Const colHIN_GAI% = 1           '�i��(�O��)
Private Const colHIN_NAI% = 2           '�i��(����)

Private Const colHIN_NAME% = 3          '�i��


Private Const colN_YOTEI_DT% = 4        '���ח\���
Private Const colN_YOTEI_QTY% = 5       '���ח\�萔
Private Const colSHIIRE% = 6            '�d���^�x����



Private Const pcmbBU% = 0
Private Const pcmbDATA_KB% = 1

Private DATA_KB     As Variant          '�f�[�^�敪

Private EXCEL_DATA  As Variant


Private SA_HIN_GAI      As Integer
Private SA_YOTEI_DT     As Integer
Private SA_SHIIRE       As Integer
Private SA_YOTEI_QTY    As Integer

Private SA_DAY          As Integer


Private CV_HIN_GAI      As Integer
Private CV_YOTEI_DT     As Integer
Private CV_YOTEI_QTY    As Integer

Private CV_DAY          As Integer

Private EP_HIN_GAI      As Integer
Private EP_YOTEI_DT     As Integer
Private EP_YOTEI_QTY    As Integer

Private EP_DAY          As Integer


Private PL_HIN_GAI      As Integer
Private PL_YOTEI_DT     As Integer
Private PL_YOTEI_QTY    As Integer

Private PL_DAY          As Integer


Private PP_HIN_GAI      As Integer
Private PP_YOTEI_DT     As Integer
Private PP_YOTEI_QTY    As Integer

Private PP_DAY          As Integer

Private PP_JYOGAI_TBL   As Variant




'2011.12.26
Private SN_JGYOBU       As Integer

Private SN_HIN_GAI      As Integer
Private SN_YOTEI_DT     As Integer
Private SN_YOTEI_QTY    As Integer

Private SN_DAY          As Integer
'2011.12.26


'2012.02.13
Private NA_JGYOBU       As Integer

Private NA_HIN_GAI      As Integer
Private NA_YOTEI_DT     As Integer
Private NA_YOTEI_QTY    As Integer

Private NA_DAY          As Integer

Private NA_JYOGAI_TBL   As Variant
'2012.02.13




Private HIN_NOT_INSERT  As Integer  '2012.04.13


Private Const LAST_UPDATE_DAY$ = "[PLN0010] 2012.04.13 10:30"

Private Sub Command1_Click(Index As Integer)

Dim sWk         As String
Dim i           As Long
Dim j           As Long


    Select Case Index



        Case 0          '�Ǎ���


            For i = 0 To UBound(EXCEL_DATA)

                If InStr(Trim(Text1.Text), EXCEL_DATA(i)) <> 0 Then
                    Exit For
                End If

            Next i


            If i > UBound(EXCEL_DATA) Then
                MsgBox "EXCEL�f�[�^�Ƃ��ĔF���o���܂���B�t�@�C�������ē��͂��Ă��������B"
                Text1.SetFocus
                Exit Sub
            End If

            '�捞���ް��\��
            
            Select Case Right(Combo1(pcmbDATA_KB).Text, 2)
                Case "SA"
                '�T�t�@�C��
                    If List_Disp_SA_Proc() Then
                        Unload Me
                    End If

                Case "CV"
                '�b�`�m�u�`�r
                    If List_Disp_CV_Proc() Then
                        Unload Me
                    End If

                Case "EP"
                '�d�o���[��
                    If List_Disp_EP_Proc() Then
                        Unload Me
                    End If

                Case "PL"
                '�o�k�t�r
                    If List_Disp_PL_Proc() Then
                        Unload Me
                    End If


                Case "PP"
                '�񓚔[��
                    If List_Disp_PP_Proc() Then
                        Unload Me
                    End If



                Case "SN"
                '���є[����       2011.12.26
                    If List_Disp_SN_Proc() Then
                        Unload Me
                    End If
                
                Case "NA"
                '�[���񓚏�(PPSC)   2012.02.13
                    If List_Disp_NA_Proc() Then
                        Unload Me
                    End If
                
                

            End Select



            If PLN_Y_NYUKA.Count(1) > 0 Then
                Command1(1).Enabled = True
                SHORI(1).Enabled = True
            
                Command1(2).Enabled = True
                SHORI(2).Enabled = True
            
            Else
                Command1(1).Enabled = False
                SHORI(1).Enabled = False
            
                Command1(2).Enabled = False
                SHORI(2).Enabled = False
            
            
            End If


        Case 1          '�o�^-->�V�K    2012.02.13


            If Update_Proc(1) Then
                Unload Me
            End If

        Case 2          '�ǉ�           2012.02.13


            If Update_Proc(2) Then
                Unload Me
            End If




        Case 3          '�I��   2-->3   2012.02.13

            Unload Me
    
    End Select



'    Command1(Index).SetFocus


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �j���� �c������ �O����
'----------------------------------------------------------------------------
    
    Select Case KeyCode
        Case vbKeyF12
            Unload Me
    End Select
    
    
    

End Sub

Private Sub Form_Load()


Dim c       As String * 128



    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[���i���v��V�X�e��]���i���p���ח\��f�[�^�Ǎ���", Me.hwnd, 0)
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

                                
                                '���ƕ���荞��
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
                                
                                '�f�[�^�敪
    If GetIni(App.EXEName, "DATA_KB", App.EXEName, c) Then
        Beep
        MsgBox "�f�[�^�敪�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    DATA_KB = Split(Trim(c), ",", -1)


                                'EXCEL�g���q
    If GetIni(App.EXEName, "EXCEL", App.EXEName, c) Then
        Beep
        MsgBox "EXCEL�g���q�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    EXCEL_DATA = Split(Trim(c), ",", -1)



'2012.04.13
    If GetIni(App.EXEName, "HIN_NOT_INSERT", App.EXEName, c) Then
        HIN_NOT_INSERT = vbUnchecked
    Else
        If Trim(c) = "1" Then
            HIN_NOT_INSERT = vbChecked
        Else
            HIN_NOT_INSERT = vbUnchecked
        End If
    End If

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �T�t�@�C��
                                'EXCEL ��
    If GetIni(App.EXEName, "SA_HIN_GAI", App.EXEName, c) Then
        Beep
        MsgBox "�T�t�@�C���i�ԗ�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "�T�t�@�C���i�ԗ�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            SA_HIN_GAI = Val(Trim(c))
        End If
    End If
    
    If GetIni(App.EXEName, "SA_YOTEI_DT", App.EXEName, c) Then
        Beep
        MsgBox "�T�t�@�C���\�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "�T�t�@�C���\�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            SA_YOTEI_DT = Val(Trim(c))
        End If
    End If
    
    If GetIni(App.EXEName, "SA_SHIIRE", App.EXEName, c) Then
        Beep
        MsgBox "�T�t�@�C���d�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "�T�t�@�C���d�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            SA_SHIIRE = Val(Trim(c))
        End If
    End If
    
    If GetIni(App.EXEName, "SA_YOTEI_QTY", App.EXEName, c) Then
        Beep
        MsgBox "�T�t�@�C���\�萔��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "�T�t�@�C���\�萔��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            SA_YOTEI_QTY = Val(Trim(c))
        End If
    End If

    If GetIni(App.EXEName, "SA_DAY", App.EXEName, c) Then
        SA_DAY = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            SA_DAY = 0
    
        Else
            SA_DAY = Val(Trim(c))
        End If
    End If

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �T�t�@�C��
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �b�`�m�u�`�r
                                'EXCEL ��
    If GetIni(App.EXEName, "CV_HIN_GAI", App.EXEName, c) Then
        Beep
        MsgBox "�b�`�m�u�`�r�i�ԗ�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "�b�`�m�u�`�r�i�ԗ�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            CV_HIN_GAI = Val(Trim(c))
        End If
    End If
    
    If GetIni(App.EXEName, "CV_YOTEI_DT", App.EXEName, c) Then
        Beep
        MsgBox "�b�`�m�u�`�r�\�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "�b�`�m�u�`�r�\�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            CV_YOTEI_DT = Val(Trim(c))
        End If
    End If
    
    If GetIni(App.EXEName, "CV_YOTEI_QTY", App.EXEName, c) Then
        Beep
        MsgBox "�b�`�m�u�`�r�\�萔��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "�b�`�m�u�`�r�\�萔��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            CV_YOTEI_QTY = Val(Trim(c))
        End If
    End If

    If GetIni(App.EXEName, "CV_DAY", App.EXEName, c) Then
        CV_DAY = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            CV_DAY = 0
    
        Else
            CV_DAY = Val(Trim(c))
        End If
    End If

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �b�`�m�u�`�r
    
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  EP���[��
                                'EXCEL ��
    If GetIni(App.EXEName, "EP_HIN_GAI", App.EXEName, c) Then
        Beep
        MsgBox "EP���[���i�ԗ�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "EP���[���i�ԗ�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            EP_HIN_GAI = Val(Trim(c))
        End If
    End If
    
    If GetIni(App.EXEName, "EP_YOTEI_DT", App.EXEName, c) Then
        Beep
        MsgBox "EP���[���\�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "EP���[���\�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            EP_YOTEI_DT = Val(Trim(c))
        End If
    End If
    
    If GetIni(App.EXEName, "EP_YOTEI_QTY", App.EXEName, c) Then
        Beep
        MsgBox "EP���[���\�萔��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "EP���[���\�萔��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            EP_YOTEI_QTY = Val(Trim(c))
        End If
    End If

    If GetIni(App.EXEName, "EP_DAY", App.EXEName, c) Then
        EP_DAY = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            EP_DAY = 0
    
        Else
            EP_DAY = Val(Trim(c))
        End If
    End If

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  EP���[��
    


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �o�k�t�r
                                'EXCEL ��
    If GetIni(App.EXEName, "PL_HIN_GAI", App.EXEName, c) Then
        Beep
        MsgBox "�o�k�t�r�i�ԗ�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "�o�k�t�r�i�ԗ�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            PL_HIN_GAI = Val(Trim(c))
        End If
    End If
    
    If GetIni(App.EXEName, "PL_YOTEI_DT", App.EXEName, c) Then
        Beep
        MsgBox "�o�k�t�r�\�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "�o�k�t�r�\�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            PL_YOTEI_DT = Val(Trim(c))
        End If
    End If
    
    If GetIni(App.EXEName, "PL_YOTEI_QTY", App.EXEName, c) Then
        Beep
        MsgBox "�o�k�t�r�\�萔��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "�o�k�t�r�\�萔��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            PL_YOTEI_QTY = Val(Trim(c))
        End If
    End If

    If GetIni(App.EXEName, "PL_DAY", App.EXEName, c) Then
        PL_DAY = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            PL_DAY = 0
    
        Else
            PL_DAY = Val(Trim(c))
        End If
    End If

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �o�k�t�r


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �񓚔[��
                                'EXCEL ��
    If GetIni(App.EXEName, "PP_HIN_GAI", App.EXEName, c) Then
        PP_HIN_GAI = 0
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "�񓚔[���i�ԗ�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            PP_HIN_GAI = Val(Trim(c))
        End If
    End If
    
    If GetIni(App.EXEName, "PP_YOTEI_DT", App.EXEName, c) Then
        PP_YOTEI_DT = 0
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "�񓚔[���\�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            PP_YOTEI_DT = Val(Trim(c))
        End If
    End If
    
    If GetIni(App.EXEName, "PP_YOTEI_QTY", App.EXEName, c) Then
        PP_YOTEI_QTY = 0
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "�񓚔[���\�萔��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            PP_YOTEI_QTY = Val(Trim(c))
        End If
    End If

    If GetIni(App.EXEName, "PP_DAY", App.EXEName, c) Then
        PP_DAY = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            PP_DAY = 0
    
        Else
            PP_DAY = Val(Trim(c))
        End If
    End If


    If GetIni(App.EXEName, "PP_JYOGAI", App.EXEName, c) Then
        c = "*"
    End If
    PP_JYOGAI_TBL = Split(Trim(c), ",", -1)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �񓚔[��



'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  ���є[����    2011.12.26
                                'EXCEL ��
    
    If GetIni(App.EXEName, "SN_JGYOBU", App.EXEName, c) Then
        SN_JGYOBU = 0
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "BU�敪��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            SN_JGYOBU = Val(Trim(c))
        End If
    End If
    
    
    If GetIni(App.EXEName, "SN_HIN_GAI", App.EXEName, c) Then
        SN_HIN_GAI = 0
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "���є[���񓚗�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            SN_HIN_GAI = Val(Trim(c))
        End If
    End If
    
    If GetIni(App.EXEName, "SN_YOTEI_DT", App.EXEName, c) Then
        SN_YOTEI_DT = 0
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "���є[���񓚗�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            SN_YOTEI_DT = Val(Trim(c))
        End If
    End If
    
    If GetIni(App.EXEName, "SN_YOTEI_QTY", App.EXEName, c) Then
        SN_YOTEI_QTY = 0
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "���є[���񓚗\�萔��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            SN_YOTEI_QTY = Val(Trim(c))
        End If
    End If

    If GetIni(App.EXEName, "SN_DAY", App.EXEName, c) Then
        SN_DAY = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            SN_DAY = 0
    
        Else
            SN_DAY = Val(Trim(c))
        End If
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  ���є[����    2011.12.26


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �[���񓚏�(PPSC)�@2012.02.13
                                'EXCEL ��
    If GetIni(App.EXEName, "NA_HIN_GAI", App.EXEName, c) Then
        NA_HIN_GAI = 0
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "�[���񓚏�(PPSC)�i�ԗ�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            NA_HIN_GAI = Val(Trim(c))
        End If
    End If
    
    If GetIni(App.EXEName, "NA_YOTEI_DT", App.EXEName, c) Then
        NA_YOTEI_DT = 0
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "�[���񓚏�(PPSC)�\�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            NA_YOTEI_DT = Val(Trim(c))
        End If
    End If
    
    If GetIni(App.EXEName, "PP_YOTEI_QTY", App.EXEName, c) Then
        NA_YOTEI_QTY = 0
    Else
        If Not IsNumeric(Trim(c)) Then
    
            Beep
            MsgBox "�[���񓚏�(PPSC)�\�萔��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
    
        Else
            NA_YOTEI_QTY = Val(Trim(c))
        End If
    End If

    If GetIni(App.EXEName, "NA_DAY", App.EXEName, c) Then
        NA_DAY = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            NA_DAY = 0
    
        Else
            NA_DAY = Val(Trim(c))
        End If
    End If


    If GetIni(App.EXEName, "NA_JYOGAI", App.EXEName, c) Then
        c = "*"
    End If
    NA_JYOGAI_TBL = Split(Trim(c), ",", -1)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �[���񓚏��@    2012.02.13





    PLN00101.Caption = PLN00101.Caption & " " & LAST_UPDATE_DAY


    Call Bu_Set_Proc
    Call Data_Kb_Set_Proc
                                '���i���p���ח\��t�@�C���n�o�d�m
    If PLN_Y_NYUKA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenRead) Then
        Unload Me
    End If



    Check1(0).Value = HIN_NOT_INSERT

End Sub


Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Text1.Text = Trim(Data.Files(1))

'    Text1.Text = Data.GetData(vbCFText)

    Command1(0).Value = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
Dim sts     As Integer
    
    sts = BTRV(BtOpClose, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K0_PLN_Y_NYUKA, Len(K0_PLN_Y_NYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���p���ח\��t�@�C��")
        End If
    End If
    
    
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K0_PLN_Y_NYUKA, Len(K0_PLN_Y_NYUKA), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If


    Set PLN00101 = Nothing



    End

End Sub

Private Sub SHORI_Click(Index As Integer)

    Select Case Index
    
        Case 0
            Command1(0).Value = True
        Case 1
            Command1(1).Value = True
        Case 2
            Command1(2).Value = True
    End Select



End Sub

Private Sub TDBGrid1_OLEDragDrop(ByVal Data As TrueDBGrid80.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Text1.Text = Trim(Data.Files(0))
'    Text1.Text = Data.GetData(0)


    Command1(0).Value = True


End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Text1.Text = Trim(Data.Files(1))
    
    Command1(0).Value = True


'    If Data.GetFormat(vbCFText) Then
'        Text1.Text = Data.GetData(vbCFText)
'        Command1(0).Value = True
'    End If

End Sub

Private Function Update_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   �u���i���p���ח\��t�@�C���v�o�^����
'----------------------------------------------------------------------------

Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer
    
Dim INS_NOW         As String * 14
    
Dim KEY_NO          As Long

Dim Row             As Long

'Dim c               As String * 128
'Dim FullPath        As String
    
Dim Ins_Flg         As Boolean      '2012.04.13
    
    
    If PLN_Y_NYUKA.Count(1) = 0 Then
        Exit Function
    End If



    Update_Proc = True
    
    Call Input_Lock

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   �������݂ɂ������f�[�^�̍폜�L���𔻒�
    If Mode = 1 Then

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "���i���p���ח\��f�[�^�폜�����@�����J�n�I�I", Me.hwnd, 0)
    
    '    sts = GetIni("FILE", PLN_Y_NYUKA_ID, "SYS", c)
    '    If sts <> False Then
    '        Call LOG_OUT(LOG_F, "SYSTEM.INI �ǂݍ��݃G���[<PLN_Y_NYUKA CREATE>")
    '        PLN_Y_NYUKA_Create = True
    '        Exit Function
    '    End If
    '
    '    FullPath = RTrim$(c)
    '
    '    On Error Resume Next
    '    Kill FullPath
    '    On Error GoTo 0
        
        
        
        
        
            
        Call UniCode_Conv(K3_PLN_Y_NYUKA.JGYOBU, Right(Combo1(pcmbBU).Text, 1))
        Call UniCode_Conv(K3_PLN_Y_NYUKA.DATA_KB, Right(Combo1(pcmbDATA_KB).Text, 2))
        
        Call UniCode_Conv(K3_PLN_Y_NYUKA.NAIGAI, "")
        Call UniCode_Conv(K3_PLN_Y_NYUKA.HIN_GAI, "")
        Call UniCode_Conv(K3_PLN_Y_NYUKA.N_YOTEI_DT, "")
        Call UniCode_Conv(K3_PLN_Y_NYUKA.SEQ_NO, "")
        
        
        com = BtOpGetGreaterEqual
    
        Do
            DoEvents
            Do
                sts = BTRV(com, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K3_PLN_Y_NYUKA, Len(K3_PLN_Y_NYUKA), 3)
            
                Select Case sts
                    Case BtNoErr
                        If StrConv(PLN_Y_NYUKA_R.JGYOBU, vbUnicode) <> Right(Combo1(pcmbBU).Text, 1) Or _
                            StrConv(PLN_Y_NYUKA_R.DATA_KB, vbUnicode) <> Right(Combo1(pcmbDATA_KB).Text, 2) Then
                            sts = BtErrEOF
                            Exit Do
                        
                        End If
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        
                        Beep
                        ans = MsgBox("�u���i���p���ח\��f�[�^�v���[���Ńf�[�^�g�p���ł��B<Y_SYUKA_TEI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Call Input_UnLock
                            Exit Function
                        End If
                    
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call Input_UnLock
                        Call File_Error(sts, com, "���i���p���ח\��f�[�^")
                        Exit Function
                End Select
            Loop
                
            If sts = BtErrEOF Then
                Exit Do
            End If
        
        
            Do
                sts = BTRV(BtOpDelete, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K3_PLN_Y_NYUKA, Len(K3_PLN_Y_NYUKA), 3)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        
                        Beep
                        ans = MsgBox("�u���i���p���ח\��f�[�^�v���[���Ńf�[�^�g�p���ł��B<Y_SYUKA_TEI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Call Input_UnLock
                            Exit Function
                        End If
                    
                    Case Else
                        Call Input_UnLock
                        Call File_Error(sts, com, "���i���p���ח\��f�[�^")
                        Exit Function
                End Select
            Loop
        
        
        
            com = BtOpGetNext
        
        
        Loop
    
    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   �������݂ɂ������f�[�^�̍폜�L���𔻒�


    sts = BTRV(BtOpGetLast, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K4_PLN_Y_NYUKA, Len(K4_PLN_Y_NYUKA), 4)
    Select Case sts
        Case BtNoErr
            
            KEY_NO = Val(StrConv(PLN_Y_NYUKA_R.KEY_NO, vbUnicode))
            If KEY_NO = 99999999 Then
                KEY_NO = 0
            End If
        
        Case BtErrEOF
            
            KEY_NO = 0
        
        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
            
            Beep
            ans = MsgBox("�u���i���p���ח\��f�[�^�v���[���Ńf�[�^�g�p���ł��B<PLN_Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
            If ans = vbCancel Then
                Call Input_UnLock
                Exit Function
            End If
        
        Case Else
            Call Input_UnLock
            Call File_Error(sts, com, "���i���p���ח\��f�[�^")
            Exit Function
    End Select





hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i���p���ח\��f�[�^�o�^�����@�����J�n�I�I", Me.hwnd, 0)
                                    
    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")
                                    
                                    '�e�[�u�����Z�b�g
    
    For Row = 1 To PLN_Y_NYUKA.UpperBound(1)
        
        
        DoEvents
        
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.04.13  �i�ڃ}�X�^�̗L���ɂ���荞�ݏ�����ύX
        Ins_Flg = True
        If Check1(0).Value = vbUnchecked Then
        
            Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo1(pcmbBU), 1))
            Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
            Call UniCode_Conv(K0_ITEM.HIN_GAI, PLN_Y_NYUKA(Row, colHIN_GAI))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                            
                    Call UniCode_Conv(K2_ITEM.JGYOBU, Right(Combo1(pcmbBU), 1))
                    Call UniCode_Conv(K2_ITEM.NAIGAI, "1")
                    Call UniCode_Conv(K2_ITEM.HIN_NAI, PLN_Y_NYUKA(Row, colHIN_NAI))
                        
                        
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                    Select Case sts
                        Case BtNoErr
                        
                        
                        Case BtErrKeyNotFound

                            Ins_Flg = False
                        
                        
                        Case Else
                            Call Input_UnLock
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            Exit Function
                    End Select
                            
                Case Else
                    Call Input_UnLock
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Function
            End Select
        End If
        If Ins_Flg Then
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.04.13
            Call PLN_Y_NYUKA_CLR
            
            Call UniCode_Conv(PLN_Y_NYUKA_R.TORIKOMI_DT, Format(Now, "YYYYMMDD"))
            Call UniCode_Conv(PLN_Y_NYUKA_R.JGYOBU, Right(Combo1(pcmbBU).Text, 1))
            Call UniCode_Conv(PLN_Y_NYUKA_R.DATA_KB, Right(Combo1(pcmbDATA_KB).Text, 2))
            Call UniCode_Conv(PLN_Y_NYUKA_R.HIN_GAI, PLN_Y_NYUKA(Row, colHIN_GAI))
            Call UniCode_Conv(PLN_Y_NYUKA_R.N_YOTEI_DT, Format(PLN_Y_NYUKA(Row, colN_YOTEI_DT), "YYYYMMDD"))
            
            '2012.04.02
            If IsNumeric(PLN_Y_NYUKA(Row, colN_YOTEI_QTY)) Then
                Call UniCode_Conv(PLN_Y_NYUKA_R.N_YOTEI_QTY, Format(CLng(PLN_Y_NYUKA(Row, colN_YOTEI_QTY)), "00000000"))
            Else
                Call UniCode_Conv(PLN_Y_NYUKA_R.N_YOTEI_QTY, PLN_Y_NYUKA(Row, colN_YOTEI_QTY))
            End If
            '2012.04.02
            Call UniCode_Conv(PLN_Y_NYUKA_R.HIN_NAI, PLN_Y_NYUKA(Row, colHIN_NAI))
            Call UniCode_Conv(PLN_Y_NYUKA_R.N_YOTEI_DT_MOTO, Format(PLN_Y_NYUKA(Row, colN_YOTEI_DT), "YYYYMMDD"))
            
            '2012.04.02
            If IsNumeric(PLN_Y_NYUKA(Row, colN_YOTEI_QTY)) Then
                Call UniCode_Conv(PLN_Y_NYUKA_R.N_YOTEI_QTY_MOTO, Format(CLng(PLN_Y_NYUKA(Row, colN_YOTEI_QTY)), "00000000"))
            Else
                Call UniCode_Conv(PLN_Y_NYUKA_R.N_YOTEI_QTY, PLN_Y_NYUKA(Row, colN_YOTEI_QTY))
            End If
            '2012.04.02
                        
                    
            Call UniCode_Conv(PLN_Y_NYUKA_R.SHIIRE, PLN_Y_NYUKA(Row, colSHIIRE))
                    
            KEY_NO = KEY_NO + 1
            Call UniCode_Conv(PLN_Y_NYUKA_R.KEY_NO, Format(KEY_NO, "00000000"))
    
            Call UniCode_Conv(PLN_Y_NYUKA_R.INS_TANTO, App.EXEName)
            Call UniCode_Conv(PLN_Y_NYUKA_R.Ins_DateTime, INS_NOW)
                        
            Do
                sts = BTRV(BtOpInsert, PLN_Y_NYUKA_POS, PLN_Y_NYUKA_R, Len(PLN_Y_NYUKA_R), K4_PLN_Y_NYUKA, Len(K4_PLN_Y_NYUKA), 4)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    
                    
                    
                    Case BtErrDuplicates
                    
                        KEY_NO = KEY_NO + 1
                        Call UniCode_Conv(PLN_Y_NYUKA_R.KEY_NO, Format(KEY_NO, "00000000"))
                    
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        
                        Beep
                        ans = MsgBox("�u���i���p���ח\��f�[�^�v���[���Ńf�[�^�g�p���ł��B<Y_SYUKA_TEI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Call Input_UnLock
                            Exit Function
                        End If
                    
                    Case Else
                        Call Input_UnLock
                        Call File_Error(sts, BtOpInsert, "���i���p���ח\��f�[�^")
                        Exit Function
                End Select
                    
            Loop
            
    
            Set TDBGrid1.Array = PLN_Y_NYUKA
            TDBGrid1.ReBind
            
            TDBGrid1.Update
            TDBGrid1.Bookmark = Row
        
        End If      '2012.04.13
    Next Row


    Set TDBGrid1.Array = PLN_Y_NYUKA
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst



hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i���p���ח\��f�[�^�@�����I���I�I", Me.hwnd, 0)





    Update_Proc = False
    Call Input_UnLock
    Exit Function

Error_Proc:
    
    MsgBox "Err.Number= " & Err.Number & " " & Err.Description
    Call Input_UnLock

End Function

Private Function List_Disp_SA_Proc() As Integer
'----------------------------------------------------------------------------
'                   �u���i���p���ח\��t�@�C���v�Ǎ��ݏ��� �T�t�@�C��
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
    
Dim INS_NOW         As String * 14
    
    
    

Dim xlApp           As Object
Dim xlBook          As Object
Dim xlSheet         As Object

Dim Row             As Long

Dim END_GYO         As Integer
Dim SKIP_F          As Boolean


Dim wkHIN_GAI       As String * 20
Dim wkYOTEI_DT      As String * 10
Dim wkSHIIRE        As String * 8
Dim wkYOTEI_Qty     As Long

Dim i               As Long

    List_Disp_SA_Proc = True

    Call Input_Lock







    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlApp.Worksheets(1)
    
    
    On Error GoTo Error_Proc

    '2011.12.03
    xlApp.Workbooks.Open (Text1.Text), ReadOnly:=True, UpdateLinks:=0

    On Error GoTo 0



hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i���p���ח\��t�@�C��[�T�t�@�C��]�@�\�������J�n�I�I", Me.hwnd, 0)


    Set PLN_Y_NYUKA = Nothing
    
    Row = Min_Row - 1
    lblDisp_Count.Caption = ""





    END_GYO = 0
    For i = 1 To 1048576
        
        SKIP_F = False
        
        
        On Error GoTo Error_Proc        '2011.12.03

            
        If Trim(xlSheet.Application.Cells(i, SA_HIN_GAI)) = "" And _
            Trim(xlSheet.Application.Cells(i, SA_YOTEI_DT)) = "" And _
            Trim(xlSheet.Application.Cells(i, SA_SHIIRE)) = "" And _
            Trim(xlSheet.Application.Cells(i, SA_YOTEI_QTY)) = "" Then
        
            SKIP_F = True
            END_GYO = END_GYO + 1
            
            If END_GYO > 5 Then
                Exit For
            End If
        Else
            
            
            
            END_GYO = 0
    
            If Trim(xlSheet.Application.Cells(i, SA_HIN_GAI)) = "" Or _
                Trim(xlSheet.Application.Cells(i, SA_YOTEI_DT)) = "" Or _
                Trim(xlSheet.Application.Cells(i, SA_SHIIRE)) = "" Or _
                Trim(xlSheet.Application.Cells(i, SA_YOTEI_QTY)) = "" Then
        
                SKIP_F = True
        
            Else
                '�i��
                wkHIN_GAI = Trim(xlSheet.Application.Cells(i, SA_HIN_GAI))
                '���t�G���[�`�F�b�N
                If Len(Trim(xlSheet.Application.Cells(i, SA_YOTEI_DT))) < 8 Then
                    SKIP_F = True
                Else
                    wkYOTEI_DT = xlSheet.Application.Cells(i, SA_YOTEI_DT)
                    If Len(Trim(wkYOTEI_DT)) = 8 Then
                     
                         
                        wkYOTEI_DT = Mid(wkYOTEI_DT, 1, 4) & "/" & Mid(wkYOTEI_DT, 5, 2) & "/" & Mid(wkYOTEI_DT, 7, 2)
                        
                    End If
                    
                    If Not IsDate(wkYOTEI_DT) Then
                        SKIP_F = True
                    End If
                End If
                '�d����
                wkSHIIRE = Trim(xlSheet.Application.Cells(i, SA_SHIIRE))
                '�\�萔�G���[�`�F�b�N
                If Not IsNumeric(Trim(xlSheet.Application.Cells(i, SA_YOTEI_QTY))) Then
                    SKIP_F = True
                Else
                    wkYOTEI_Qty = CLng(Trim(xlSheet.Application.Cells(i, SA_YOTEI_QTY)))
                End If
            
            
                If Not SKIP_F Then
            
                    Row = Row + 1
                    PLN_Y_NYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
                
                    PLN_Y_NYUKA(Row, colNO) = Row
            
            
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo1(pcmbBU), 1))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, wkHIN_GAI)
            
            
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        
                        Case BtErrKeyNotFound
                        
                            Call UniCode_Conv(K2_ITEM.JGYOBU, Right(Combo1(pcmbBU), 1))
                            Call UniCode_Conv(K2_ITEM.NAIGAI, "1")
                            Call UniCode_Conv(K2_ITEM.HIN_NAI, wkHIN_GAI)
                    
                    
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                            Select Case sts
                                Case BtNoErr
                                
                                
                                Case BtErrKeyNotFound
                                
                                    Call UniCode_Conv(ITEMREC.HIN_GAI, "")
                                    Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                                
                                
                                Case Else
                                    Call Input_UnLock
                                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                    Exit Function
                            End Select
                        
                        Case Else
                            Call Input_UnLock
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            Exit Function
                    End Select
            
            

                    PLN_Y_NYUKA(Row, colHIN_GAI) = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
                    PLN_Y_NYUKA(Row, colHIN_NAI) = Trim(wkHIN_GAI)
                    PLN_Y_NYUKA(Row, colHIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                    PLN_Y_NYUKA(Row, colN_YOTEI_DT) = DateAdd("d", SA_DAY, wkYOTEI_DT)
                    PLN_Y_NYUKA(Row, colN_YOTEI_QTY) = Format(wkYOTEI_Qty, "#,##0")
                    PLN_Y_NYUKA(Row, colSHIIRE) = wkSHIIRE


                End If
            End If
        
        End If
    
    
    
    Next i

    On Error GoTo 0                 '2011.12.03


    Set TDBGrid1.Array = PLN_Y_NYUKA
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst






    lblDisp_Count.Caption = Format(Row, "#0") & "��"








hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i���p���ח\��t�@�C��[�T�t�@�C��]�@�\�������I���I�I", Me.hwnd, 0)



    Call Input_UnLock

    xlApp.DisplayAlerts = False
    xlBook.Close False
    xlApp.Quit 'EXCEL�����
    Set xlApp = Nothing

    List_Disp_SA_Proc = False
    Exit Function

Error_Proc:
    

    Select Case Err.Number
        
        '52 �t�@�C�����܂��͔ԍ����s���ł��B
        '53 �t�@�C����������܂���B
        '54 �t�@�C�� ���[�h���s���ł��B
        '55 �t�@�C���͊��ɊJ����Ă��܂��B
        '57 �f�o�C�X I/O �G���[�ł��B
        '59 ���R�[�h������v���܂���B
        '61 �f�B�X�N�̋󂫗e�ʂ��s�����Ă��܂��B
        '62 �t�@�C���ɂ���ȏ�f�[�^������܂���B
        '63 ���R�[�h�ԍ����s���ł��B
        '68 �f�o�C�X����������Ă��܂���B
        '70 �������݂ł��܂���B
        '71 �f�B�X�N����������Ă��܂���B
        '75 �p�X���������ł��B
        '76 �p�X��������܂���B
        Case 52, 53, 54, 55, 57, 59, 61, 62, 63, 68, 70, 71, 75, 76
            
            
            MsgBox "�w��̃t�@�C����������܂���B" & Chr(13) & Chr(10) & "�������t�@�C��������͂��Ă��������B"
            
            
            
            List_Disp_SA_Proc = False      '


    '2011.12.03
        Case 13
        
            MsgBox "�Ǎ��ݑΏۂ�EXCEL�f�[�^�Ɉُ킪�L��܂��B���e���m�F��A�Ď��s���Ă��������B"
            
            xlApp.DisplayAlerts = False
            xlBook.Close False
            xlApp.Quit 'EXCEL�����
            Set xlApp = Nothing
            
            
            
            
            
            
            List_Disp_SA_Proc = False      '
            


        Case Else
            MsgBox Err.Description
    
    '2011.12.03
    End Select
    
    On Error GoTo 0
    
    
    Call Input_UnLock

End Function


Private Function List_Disp_CV_Proc() As Integer
'----------------------------------------------------------------------------
'                   �u���i���p���ח\��t�@�C���v�Ǎ��ݏ��� �b�`�m�u�`�r
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
    
Dim INS_NOW         As String * 14
    
    
    

Dim xlApp           As Object
Dim xlBook          As Object
Dim xlSheet         As Object

Dim Row             As Long

Dim END_GYO         As Integer
Dim SKIP_F          As Boolean


Dim wkHIN_GAI       As String * 20
Dim wkYOTEI_DT      As String * 10
Dim wkSHIIRE        As String * 8
Dim wkYOTEI_Qty     As Long

Dim i               As Long
Dim j               As Long

    List_Disp_CV_Proc = True

    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i���p���ח\��t�@�C��[�b�`�m�u�`�r]�@�\�������J�n�I�I", Me.hwnd, 0)





    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    
    
    On Error GoTo Error_Proc

    xlApp.Workbooks.Open (Text1.Text), ReadOnly:=True, UpdateLinks:=0

    On Error GoTo 0
    
    
    
    Set PLN_Y_NYUKA = Nothing
    
    Row = Min_Row - 1
    lblDisp_Count.Caption = ""
    
    
    
    
    
    
    For j = 1 To xlApp.Worksheets.Count
    
        Set xlSheet = xlApp.Worksheets(j)
        xlSheet.Activate
    
    
        END_GYO = 0
        For i = 1 To 1048576
            
            SKIP_F = False
            
            If Trim(xlSheet.Application.Cells(i, CV_HIN_GAI)) = "" And _
                Trim(xlSheet.Application.Cells(i, CV_YOTEI_DT)) = "" And _
                Trim(xlSheet.Application.Cells(i, CV_YOTEI_QTY)) = "" Then
            
                SKIP_F = True
                END_GYO = END_GYO + 1
                
                If END_GYO > 5 Then
                    Exit For
                End If
            Else
                
                
                
                END_GYO = 0
        
                If Trim(xlSheet.Application.Cells(i, CV_HIN_GAI)) = "" Or _
                    Trim(xlSheet.Application.Cells(i, CV_YOTEI_DT)) = "" Or _
                    Trim(xlSheet.Application.Cells(i, CV_YOTEI_QTY)) = "" Then
            
                    SKIP_F = True
            
                Else
                    '�i��
                    wkHIN_GAI = Trim(xlSheet.Application.Cells(i, CV_HIN_GAI))
                    '���t�G���[�`�F�b�N
                    If Len(Trim(xlSheet.Application.Cells(i, CV_YOTEI_DT))) < 8 Then
                        SKIP_F = True
                    Else
                        wkYOTEI_DT = xlSheet.Application.Cells(i, CV_YOTEI_DT)
                        
                        
                        If Len(Trim(wkYOTEI_DT)) = 8 Then
                         
                            wkYOTEI_DT = Mid(wkYOTEI_DT, 1, 4) & "/" & Mid(wkYOTEI_DT, 5, 2) & "/" & Mid(wkYOTEI_DT, 7, 2)
                        
                        End If
                        
                        
                        If Not IsDate(wkYOTEI_DT) Then
                            SKIP_F = True
                        End If
                    End If
                    '�\�萔�G���[�`�F�b�N
                    If Not IsNumeric(Trim(xlSheet.Application.Cells(i, CV_YOTEI_QTY))) Then
                        SKIP_F = True
                    Else
                        wkYOTEI_Qty = CLng(Trim(xlSheet.Application.Cells(i, CV_YOTEI_QTY)))
                    End If
                
                
                    If Not SKIP_F Then
                
                        Row = Row + 1
                        PLN_Y_NYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
                    
                        PLN_Y_NYUKA(Row, colNO) = Row
                
                
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo1(pcmbBU), 1))
                        Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, wkHIN_GAI)
                
                
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            
                            
                            Case BtErrKeyNotFound
                            
                                Call UniCode_Conv(K2_ITEM.JGYOBU, Right(Combo1(pcmbBU), 1))
                                Call UniCode_Conv(K2_ITEM.NAIGAI, "1")
                                Call UniCode_Conv(K2_ITEM.HIN_NAI, wkHIN_GAI)
                        
                        
                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                                Select Case sts
                                    Case BtNoErr
                                    
                                    
                                    Case BtErrKeyNotFound
                                    
                                        Call UniCode_Conv(ITEMREC.HIN_NAI, "")
                                        Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                                    
                                    
                                    Case Else
                                        Call Input_UnLock
                                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                        Exit Function
                                End Select
                            
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                Exit Function
                        End Select
                
                
    
                        PLN_Y_NYUKA(Row, colHIN_GAI) = wkHIN_GAI
                        PLN_Y_NYUKA(Row, colHIN_NAI) = Trim(StrConv(ITEMREC.HIN_NAI, vbUnicode))
                        PLN_Y_NYUKA(Row, colHIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                        PLN_Y_NYUKA(Row, colN_YOTEI_DT) = DateAdd("d", CV_DAY, wkYOTEI_DT)
                        PLN_Y_NYUKA(Row, colN_YOTEI_QTY) = Format(wkYOTEI_Qty, "#,##0")
                        PLN_Y_NYUKA(Row, colSHIIRE) = ""
    
    
                    End If
                End If
            
            End If
        
        
        
        Next i
    
    
    
    
    Next j

    Set TDBGrid1.Array = PLN_Y_NYUKA
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst






    lblDisp_Count.Caption = Format(Row, "#0") & "��"



    xlApp.DisplayAlerts = False

    xlBook.Close False
    xlApp.Quit 'EXCEL�����
    Set xlApp = Nothing




hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i���p���ח\��t�@�C��[�b�`�m�u�`�r]�@�\�������I���I�I", Me.hwnd, 0)



    Call Input_UnLock


    List_Disp_CV_Proc = False
    Exit Function

Error_Proc:
    

    Select Case Err.Number
        
        '52 �t�@�C�����܂��͔ԍ����s���ł��B
        '53 �t�@�C����������܂���B
        '54 �t�@�C�� ���[�h���s���ł��B
        '55 �t�@�C���͊��ɊJ����Ă��܂��B
        '57 �f�o�C�X I/O �G���[�ł��B
        '59 ���R�[�h������v���܂���B
        '61 �f�B�X�N�̋󂫗e�ʂ��s�����Ă��܂��B
        '62 �t�@�C���ɂ���ȏ�f�[�^������܂���B
        '63 ���R�[�h�ԍ����s���ł��B
        '68 �f�o�C�X����������Ă��܂���B
        '70 �������݂ł��܂���B
        '71 �f�B�X�N����������Ă��܂���B
        '75 �p�X���������ł��B
        '76 �p�X��������܂���B
        Case 52, 53, 54, 55, 57, 59, 61, 62, 63, 68, 70, 71, 75, 76
            
            
            MsgBox "�w��̃t�@�C����������܂���B" & Chr(13) & Chr(10) & "�������t�@�C��������͂��Ă��������B"
            
            
            
            List_Disp_CV_Proc = False      '


    '2011.12.03
        Case 13
        
            MsgBox "�Ǎ��ݑΏۂ�EXCEL�f�[�^�Ɉُ킪�L��܂��B���e���m�F��A�Ď��s���Ă��������B"
            
            xlApp.DisplayAlerts = False
            xlBook.Close False
            xlApp.Quit 'EXCEL�����
            Set xlApp = Nothing
            
            
            
            
            
            
            List_Disp_CV_Proc = False      '



        Case Else
    End Select
    Call Input_UnLock

End Function
Private Function List_Disp_EP_Proc() As Integer
'----------------------------------------------------------------------------
'                   �u���i���p���ח\��t�@�C���v�Ǎ��ݏ��� �d�o���[��
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
    
Dim INS_NOW         As String * 14
    
    
    

Dim xlApp           As Object
Dim xlBook          As Object
Dim xlSheet         As Object

Dim Row             As Long

Dim END_GYO         As Integer
Dim SKIP_F          As Boolean


Dim wkHIN_GAI       As String * 20
Dim wkYOTEI_DT      As String * 10
Dim wkSHIIRE        As String * 8
Dim wkYOTEI_Qty     As Long

Dim i               As Long

    List_Disp_EP_Proc = True

    Call Input_Lock







    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)

    On Error GoTo Error_Proc

    xlApp.Workbooks.Open (Text1.Text), ReadOnly:=True, UpdateLinks:=0


    On Error GoTo 0



hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i���p���ח\��t�@�C��[�d�o���[��]�@�\�������J�n�I�I", Me.hwnd, 0)


    Set PLN_Y_NYUKA = Nothing
    
    Row = Min_Row - 1
    lblDisp_Count.Caption = ""





    END_GYO = 0
    For i = 1 To 1048576
        
        SKIP_F = False
        
        If Trim(xlSheet.Application.Cells(i, EP_HIN_GAI)) = "" And _
            Trim(xlSheet.Application.Cells(i, EP_YOTEI_DT)) = "" And _
            Trim(xlSheet.Application.Cells(i, EP_YOTEI_QTY)) = "" Then
        
            SKIP_F = True
            END_GYO = END_GYO + 1
            
            If END_GYO > 5 Then
                Exit For
            End If
        Else
            
            
            
            END_GYO = 0
    
            If Trim(xlSheet.Application.Cells(i, EP_HIN_GAI)) = "" Or _
                Trim(xlSheet.Application.Cells(i, EP_YOTEI_DT)) = "" Or _
                Trim(xlSheet.Application.Cells(i, EP_YOTEI_QTY)) = "" Then
        
                SKIP_F = True
        
            Else
                '�i��
                wkHIN_GAI = Trim(xlSheet.Application.Cells(i, EP_HIN_GAI))
                '���t�G���[�`�F�b�N
                If Len(Trim(xlSheet.Application.Cells(i, EP_YOTEI_DT))) < 8 Then
                    SKIP_F = True
                Else
                    wkYOTEI_DT = xlSheet.Application.Cells(i, EP_YOTEI_DT)
                    
                    If Len(Trim(wkYOTEI_DT)) = 8 Then
                     
                        wkYOTEI_DT = Mid(wkYOTEI_DT, 1, 4) & "/" & Mid(wkYOTEI_DT, 5, 2) & "/" & Mid(wkYOTEI_DT, 7, 2)
                    
                    End If
                                
                    If Not IsDate(wkYOTEI_DT) Then
                        SKIP_F = True
                    End If
                End If
                '�d����
                wkSHIIRE = ""
                '�\�萔�G���[�`�F�b�N
                If Not IsNumeric(Trim(xlSheet.Application.Cells(i, EP_YOTEI_QTY))) Then
                    SKIP_F = True
                Else
                    wkYOTEI_Qty = CLng(Trim(xlSheet.Application.Cells(i, EP_YOTEI_QTY)))
                End If
            
            
                If Not SKIP_F Then
            
                    Row = Row + 1
                    PLN_Y_NYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
                
                    PLN_Y_NYUKA(Row, colNO) = Row
            
            
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo1(pcmbBU), 1))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, wkHIN_GAI)
            
            
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        
                        Case BtErrKeyNotFound
                        
                            Call UniCode_Conv(K2_ITEM.JGYOBU, Right(Combo1(pcmbBU), 1))
                            Call UniCode_Conv(K2_ITEM.NAIGAI, "1")
                            Call UniCode_Conv(K2_ITEM.HIN_NAI, wkHIN_GAI)
                    
                    
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                            Select Case sts
                                Case BtNoErr
                                
                                
                                Case BtErrKeyNotFound
                                
                                    Call UniCode_Conv(ITEMREC.HIN_NAI, "")
                                    Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                                
                                
                                Case Else
                                    Call Input_UnLock
                                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                    Exit Function
                            End Select
                        
                        Case Else
                            Call Input_UnLock
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            Exit Function
                    End Select
            
            

                    PLN_Y_NYUKA(Row, colHIN_GAI) = Trim(wkHIN_GAI)
                    PLN_Y_NYUKA(Row, colHIN_NAI) = Trim(StrConv(ITEMREC.HIN_NAI, vbUnicode))
                    PLN_Y_NYUKA(Row, colHIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                    PLN_Y_NYUKA(Row, colN_YOTEI_DT) = DateAdd("d", SA_DAY, wkYOTEI_DT)
                    PLN_Y_NYUKA(Row, colN_YOTEI_QTY) = Format(wkYOTEI_Qty, "#,##0")
                    PLN_Y_NYUKA(Row, colSHIIRE) = wkSHIIRE


                End If
            End If
        
        End If
    
    
    
    Next i


    Set TDBGrid1.Array = PLN_Y_NYUKA
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst






    lblDisp_Count.Caption = Format(Row, "#0") & "��"








hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i���p���ח\��t�@�C��[�d�o���[��]�@�\�������I���I�I", Me.hwnd, 0)



    Call Input_UnLock

    
    xlApp.DisplayAlerts = False
    xlBook.Close
    
    xlApp.Quit 'EXCEL�����
'    xlApp.DisplayAlerts = True
    
    
    Set xlApp = Nothing

    List_Disp_EP_Proc = False
    Exit Function

Error_Proc:
    

    Select Case Err.Number
        
        '52 �t�@�C�����܂��͔ԍ����s���ł��B
        '53 �t�@�C����������܂���B
        '54 �t�@�C�� ���[�h���s���ł��B
        '55 �t�@�C���͊��ɊJ����Ă��܂��B
        '57 �f�o�C�X I/O �G���[�ł��B
        '59 ���R�[�h������v���܂���B
        '61 �f�B�X�N�̋󂫗e�ʂ��s�����Ă��܂��B
        '62 �t�@�C���ɂ���ȏ�f�[�^������܂���B
        '63 ���R�[�h�ԍ����s���ł��B
        '68 �f�o�C�X����������Ă��܂���B
        '70 �������݂ł��܂���B
        '71 �f�B�X�N����������Ă��܂���B
        '75 �p�X���������ł��B
        '76 �p�X��������܂���B
        Case 52, 53, 54, 55, 57, 59, 61, 62, 63, 68, 70, 71, 75, 76
            
            
            MsgBox "�w��̃t�@�C����������܂���B" & Chr(13) & Chr(10) & "�������t�@�C��������͂��Ă��������B"
            
            
            
            List_Disp_EP_Proc = False      '


    '2011.12.03
        Case 13
        
            MsgBox "�Ǎ��ݑΏۂ�EXCEL�f�[�^�Ɉُ킪�L��܂��B���e���m�F��A�Ď��s���Ă��������B"
            
            xlApp.DisplayAlerts = False
            xlBook.Close False
            xlApp.Quit 'EXCEL�����
            Set xlApp = Nothing
            
            
            
            
            
            
            List_Disp_EP_Proc = False      '



        Case Else
    End Select
    Call Input_UnLock

End Function

Private Function List_Disp_PL_Proc() As Integer
'----------------------------------------------------------------------------
'                   �u���i���p���ח\��t�@�C���v�Ǎ��ݏ��� �o�t�k�r
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
    
Dim INS_NOW         As String * 14
    
    
    

Dim xlApp           As Object
Dim xlBook          As Object
Dim xlSheet         As Object

Dim Row             As Long

Dim END_GYO         As Integer
Dim SKIP_F          As Boolean


Dim wkHIN_GAI       As String * 20
Dim wkYOTEI_DT      As String * 10
Dim wkSHIIRE        As String * 8
Dim wkYOTEI_Qty     As Long

Dim i               As Long
Dim j               As Long

    List_Disp_PL_Proc = True

    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i���p���ח\��t�@�C��[�o�k�t�r]�@�\�������J�n�I�I", Me.hwnd, 0)





    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlApp.Worksheets(1)
    
    
    On Error GoTo Error_Proc

    xlApp.Workbooks.Open (Text1.Text), ReadOnly:=True, UpdateLinks:=0

    On Error GoTo 0
    
    
    
    Set PLN_Y_NYUKA = Nothing
    
    Row = Min_Row - 1
    lblDisp_Count.Caption = ""
    
    
    
    
    
    
    


    END_GYO = 0
    For i = 1 To 1048576
        
        SKIP_F = False
        
If i = 48 Then
Debug.Print
End If
Debug.Print xlSheet.Application.Cells(i, 1)
        
        
        If Trim(xlSheet.Application.Cells(i, PL_HIN_GAI)) = "" And _
            Trim(xlSheet.Application.Cells(i, PL_YOTEI_DT)) = "" And _
            Trim(xlSheet.Application.Cells(i, PL_YOTEI_QTY)) = "" Then
        
            SKIP_F = True
            END_GYO = END_GYO + 1
            
            If END_GYO > 100 Then
                Exit For
            End If
        Else
            
            
            
            END_GYO = 0
    
            If Trim(xlSheet.Application.Cells(i, PL_HIN_GAI)) = "" Or _
                Trim(xlSheet.Application.Cells(i, PL_YOTEI_DT)) = "" Or _
                Trim(xlSheet.Application.Cells(i, PL_YOTEI_QTY)) = "" Then
        
                SKIP_F = True
        
            Else
                '�i��
                wkHIN_GAI = Trim(xlSheet.Application.Cells(i, PL_HIN_GAI))
                '���t�G���[�`�F�b�N
                If Len(Trim(xlSheet.Application.Cells(i, PL_YOTEI_DT))) < 8 Then
                    SKIP_F = True
                Else
                    wkYOTEI_DT = xlSheet.Application.Cells(i, PL_YOTEI_DT)
                    
                    
                    If Len(Trim(wkYOTEI_DT)) = 8 Then
                     
                        wkYOTEI_DT = Mid(wkYOTEI_DT, 1, 4) & "/" & Mid(wkYOTEI_DT, 5, 2) & "/" & Mid(wkYOTEI_DT, 7, 2)
                    
                    End If
                    
                    
                    If Not IsDate(wkYOTEI_DT) Then
                        SKIP_F = True
                    End If
                End If
                '�\�萔�G���[�`�F�b�N
                If Not IsNumeric(Trim(xlSheet.Application.Cells(i, PL_YOTEI_QTY))) Then
                    SKIP_F = True
                Else
                    wkYOTEI_Qty = CLng(Trim(xlSheet.Application.Cells(i, PL_YOTEI_QTY)))
                End If
            
            
                If Not SKIP_F Then
            
                    Row = Row + 1
                    PLN_Y_NYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
                
                    PLN_Y_NYUKA(Row, colNO) = Row
            
            
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo1(pcmbBU), 1))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, wkHIN_GAI)
            
            
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        
                        Case BtErrKeyNotFound
                        
                            Call UniCode_Conv(K2_ITEM.JGYOBU, Right(Combo1(pcmbBU), 1))
                            Call UniCode_Conv(K2_ITEM.NAIGAI, "1")
                            Call UniCode_Conv(K2_ITEM.HIN_NAI, wkHIN_GAI)
                    
                    
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                            Select Case sts
                                Case BtNoErr
                                
                                
                                Case BtErrKeyNotFound
                                
                                    Call UniCode_Conv(ITEMREC.HIN_GAI, "")
                                    Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                                
                                
                                Case Else
                                    Call Input_UnLock
                                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                    Exit Function
                            End Select
                        
                        Case Else
                            Call Input_UnLock
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            Exit Function
                    End Select
            
            

                    PLN_Y_NYUKA(Row, colHIN_GAI) = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
                    PLN_Y_NYUKA(Row, colHIN_NAI) = wkHIN_GAI
                    PLN_Y_NYUKA(Row, colHIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                    PLN_Y_NYUKA(Row, colN_YOTEI_DT) = DateAdd("d", CV_DAY, wkYOTEI_DT)
                    PLN_Y_NYUKA(Row, colN_YOTEI_QTY) = Format(wkYOTEI_Qty, "#,##0")
                    PLN_Y_NYUKA(Row, colSHIIRE) = ""


                End If
            End If
        
        End If
    
    
    
    Next i
    

    Set TDBGrid1.Array = PLN_Y_NYUKA
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst






    lblDisp_Count.Caption = Format(Row, "#0") & "��"



    xlApp.DisplayAlerts = False

    xlBook.Close False
    xlApp.Quit 'EXCEL�����
    Set xlApp = Nothing




hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i���p���ח\��t�@�C��[�o�k�t�r]�@�\�������I���I�I", Me.hwnd, 0)



    Call Input_UnLock


    List_Disp_PL_Proc = False
    Exit Function

Error_Proc:
    

    Select Case Err.Number
        
        '52 �t�@�C�����܂��͔ԍ����s���ł��B
        '53 �t�@�C����������܂���B
        '54 �t�@�C�� ���[�h���s���ł��B
        '55 �t�@�C���͊��ɊJ����Ă��܂��B
        '57 �f�o�C�X I/O �G���[�ł��B
        '59 ���R�[�h������v���܂���B
        '61 �f�B�X�N�̋󂫗e�ʂ��s�����Ă��܂��B
        '62 �t�@�C���ɂ���ȏ�f�[�^������܂���B
        '63 ���R�[�h�ԍ����s���ł��B
        '68 �f�o�C�X����������Ă��܂���B
        '70 �������݂ł��܂���B
        '71 �f�B�X�N����������Ă��܂���B
        '75 �p�X���������ł��B
        '76 �p�X��������܂���B
        Case 52, 53, 54, 55, 57, 59, 61, 62, 63, 68, 70, 71, 75, 76
            
            
            MsgBox "�w��̃t�@�C����������܂���B" & Chr(13) & Chr(10) & "�������t�@�C��������͂��Ă��������B"
            
            
            
            List_Disp_PL_Proc = False      '


    '2011.12.03
        Case 13
        
            MsgBox "�Ǎ��ݑΏۂ�EXCEL�f�[�^�Ɉُ킪�L��܂��B���e���m�F��A�Ď��s���Ă��������B"
            
            xlApp.DisplayAlerts = False
            xlBook.Close False
            xlApp.Quit 'EXCEL�����
            Set xlApp = Nothing
            
            
            
            
            
            
            List_Disp_PL_Proc = False      '



        Case Else
    End Select
    Call Input_UnLock

End Function
Private Function List_Disp_PP_Proc() As Integer
'----------------------------------------------------------------------------
'                   �u���i���p���ח\��t�@�C���v�Ǎ��ݏ��� �[����
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
    
Dim INS_NOW         As String * 14
    
    
    

Dim xlApp           As Object
Dim xlBook          As Object
Dim xlSheet         As Object

Dim Row             As Long

Dim END_GYO         As Integer
Dim SKIP_F          As Boolean


Dim wkHIN_GAI       As String * 20
Dim wkYOTEI_DT      As String
Dim wkYOTEI_Qty     As Long

Dim i               As Long
Dim j               As Long

    List_Disp_PP_Proc = True

    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i���p���ח\��t�@�C��[�[����]�@�\�������J�n�I�I", Me.hwnd, 0)





    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlApp.Worksheets(1)
    
    
    On Error GoTo Error_Proc

    xlApp.Workbooks.Open (Text1.Text), ReadOnly:=True, UpdateLinks:=0

    On Error GoTo 0
    
    
    
    Set PLN_Y_NYUKA = Nothing
    
    Row = Min_Row - 1
    lblDisp_Count.Caption = ""
    
    END_GYO = 0
    
    On Error GoTo Error_Proc
    
    For i = 1 To 1048576
        SKIP_F = False
        If Trim(xlSheet.Application.Cells(i, PP_HIN_GAI)) = "" And _
            Trim(xlSheet.Application.Cells(i, PP_YOTEI_DT)) = "" And _
            Trim(xlSheet.Application.Cells(i, PP_YOTEI_QTY)) = "" Then
            SKIP_F = True
            END_GYO = END_GYO + 1
            If END_GYO > 3 Then
                Exit For
            End If
        Else
            If Trim(xlSheet.Application.Cells(i, PP_HIN_GAI)) = "" Or _
                Trim(xlSheet.Application.Cells(i, PP_YOTEI_DT)) = "" Or _
                (Trim(xlSheet.Application.Cells(i, PP_YOTEI_QTY)) <> "" And Not IsNumeric(Trim(xlSheet.Application.Cells(i, PP_YOTEI_QTY)))) Then
            Else
                END_GYO = 0
                '�i��
                wkHIN_GAI = Trim(xlSheet.Application.Cells(i, PP_HIN_GAI))
                '���t�G���[�`�F�b�N
                wkYOTEI_DT = Trim(xlSheet.Application.Cells(i, PP_YOTEI_DT))
                For j = 0 To UBound(PP_JYOGAI_TBL)
                    If wkYOTEI_DT = PP_JYOGAI_TBL(j) Then
                        SKIP_F = True
                        Exit For
                    End If
                Next j
                If Not SKIP_F Then
                    If Len(Trim(xlSheet.Application.Cells(i, PP_YOTEI_DT))) < 8 Then
                    Else
                        If Len(Trim(wkYOTEI_DT)) = 8 Then
                            wkYOTEI_DT = Mid(wkYOTEI_DT, 1, 4) & "/" & Mid(wkYOTEI_DT, 5, 2) & "/" & Mid(wkYOTEI_DT, 7, 2)
                        End If
                        If Not IsDate(wkYOTEI_DT) Then
                            wkYOTEI_DT = ""
                        Else
                            wkYOTEI_DT = wkYOTEI_DT
                        End If
                    End If
                    '�\�萔�G���[�`�F�b�N
                    If Trim(xlSheet.Application.Cells(i, PP_YOTEI_QTY)) = "" Then
                        '>>>>>>>>>>>>>>>>>>>>   2012.02.13  ���z�󔒂͏��O
                        SKIP_F = True
                        '>>>>>>>>>>>>>>>>>>>>   2012.02.13
                    Else
                        If Not IsNumeric(Trim(xlSheet.Application.Cells(i, PP_YOTEI_QTY))) Then
                            SKIP_F = True
                        Else
                            wkYOTEI_Qty = CLng(Trim(xlSheet.Application.Cells(i, PP_YOTEI_QTY)))
                        End If
                    End If
                    If Not SKIP_F Then
                        Row = Row + 1
                        PLN_Y_NYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
                        PLN_Y_NYUKA(Row, colNO) = Row
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo1(pcmbBU), 1))
                        Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, wkHIN_GAI)
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                Call UniCode_Conv(K2_ITEM.JGYOBU, Right(Combo1(pcmbBU), 1))
                                Call UniCode_Conv(K2_ITEM.NAIGAI, "1")
                                Call UniCode_Conv(K2_ITEM.HIN_NAI, wkHIN_GAI)
                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                                Select Case sts
                                    Case BtNoErr
                                    Case BtErrKeyNotFound
                                        Call UniCode_Conv(ITEMREC.HIN_GAI, "")
                                        Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                                    Case Else
                                        Call Input_UnLock
                                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                        Exit Function
                                End Select
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                Exit Function
                        End Select
                        PLN_Y_NYUKA(Row, colHIN_GAI) = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
                        PLN_Y_NYUKA(Row, colHIN_NAI) = wkHIN_GAI
                        PLN_Y_NYUKA(Row, colHIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                        If IsDate(wkYOTEI_DT) Then
                            PLN_Y_NYUKA(Row, colN_YOTEI_DT) = DateAdd("d", PP_DAY, wkYOTEI_DT)
                        Else
                            PLN_Y_NYUKA(Row, colN_YOTEI_DT) = ""
                        End If
                        PLN_Y_NYUKA(Row, colN_YOTEI_QTY) = Format(wkYOTEI_Qty, "#,##0")
                        PLN_Y_NYUKA(Row, colSHIIRE) = ""
                    End If
                End If
            End If
        End If
    Next i
    
    On Error GoTo 0


    Set TDBGrid1.Array = PLN_Y_NYUKA
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst






    lblDisp_Count.Caption = Format(Row, "#0") & "��"



    xlApp.DisplayAlerts = False

    xlBook.Close False
    xlApp.Quit 'EXCEL�����
    Set xlApp = Nothing




hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i���p���ח\��t�@�C��[�[����]�@�\�������I���I�I", Me.hwnd, 0)



    Call Input_UnLock


    List_Disp_PP_Proc = False
    Exit Function

Error_Proc:
    

    Select Case Err.Number
        
        '52 �t�@�C�����܂��͔ԍ����s���ł��B
        '53 �t�@�C����������܂���B
        '54 �t�@�C�� ���[�h���s���ł��B
        '55 �t�@�C���͊��ɊJ����Ă��܂��B
        '57 �f�o�C�X I/O �G���[�ł��B
        '59 ���R�[�h������v���܂���B
        '61 �f�B�X�N�̋󂫗e�ʂ��s�����Ă��܂��B
        '62 �t�@�C���ɂ���ȏ�f�[�^������܂���B
        '63 ���R�[�h�ԍ����s���ł��B
        '68 �f�o�C�X����������Ă��܂���B
        '70 �������݂ł��܂���B
        '71 �f�B�X�N����������Ă��܂���B
        '75 �p�X���������ł��B
        '76 �p�X��������܂���B
        Case 52, 53, 54, 55, 57, 59, 61, 62, 63, 68, 70, 71, 75, 76
            
            
            MsgBox "�w��̃t�@�C����������܂���B" & Chr(13) & Chr(10) & "�������t�@�C��������͂��Ă��������B"
            
            
            
            List_Disp_PP_Proc = False      '





    '2011.12.03
        Case 13
        
            MsgBox "�Ǎ��ݑΏۂ�EXCEL�f�[�^�Ɉُ킪�L��܂��B���e���m�F��A�Ď��s���Ă��������B"
            
            xlApp.DisplayAlerts = False
            xlBook.Close False
            xlApp.Quit 'EXCEL�����
            Set xlApp = Nothing
            
            
            
            
            
            
            List_Disp_PP_Proc = False      '
            


        Case Else
            MsgBox Err.Description
    
    '2011.12.03
    End Select
    
    On Error GoTo 0
    
    Call Input_UnLock

End Function



Private Function List_Disp_SN_Proc() As Integer
'----------------------------------------------------------------------------
'                   �u���i���p���ח\��t�@�C���v�Ǎ��ݏ��� ���є[����
'
'                   2011.12.26
'
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
    
Dim INS_NOW         As String * 14
    
    
    

Dim xlApp           As Object
Dim xlBook          As Object
Dim xlSheet         As Object

Dim Row             As Long

Dim END_GYO         As Integer
Dim SKIP_F          As Boolean


Dim wkHIN_GAI       As String * 20
Dim wkYOTEI_DT      As String
Dim wkYOTEI_Qty     As Long

Dim i               As Long
Dim j               As Long

    List_Disp_SN_Proc = True

    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i���p���ח\��t�@�C��[���є[����]�@�\�������J�n�I�I", Me.hwnd, 0)





    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlApp.Worksheets(1)
    
    
    On Error GoTo Error_Proc

    xlApp.Workbooks.Open (Text1.Text), ReadOnly:=True, UpdateLinks:=0

    On Error GoTo 0
    
    
    
    Set PLN_Y_NYUKA = Nothing
    
    Row = Min_Row - 1
    lblDisp_Count.Caption = ""
    
    END_GYO = 0
    
    On Error GoTo Error_Proc
    
    For i = 1 To 1048576
        SKIP_F = False
        If Trim(xlSheet.Application.Cells(i, SN_HIN_GAI)) = "" And _
            Trim(xlSheet.Application.Cells(i, SN_YOTEI_DT)) = "" And _
            Trim(xlSheet.Application.Cells(i, SN_YOTEI_QTY)) = "" Then
            SKIP_F = True
            END_GYO = END_GYO + 1
            If END_GYO > 3 Then
                Exit For
            End If
        Else
            If Trim(xlSheet.Application.Cells(i, SN_HIN_GAI)) = "" Or _
                Trim(xlSheet.Application.Cells(i, SN_YOTEI_DT)) = "" Or _
                (Trim(xlSheet.Application.Cells(i, SN_YOTEI_QTY)) <> "" And Not IsNumeric(Trim(xlSheet.Application.Cells(i, SN_YOTEI_QTY)))) Then
            Else
                END_GYO = 0
                
                '���ƕ�
                If Trim(xlSheet.Application.Cells(i, SN_JGYOBU)) <> Right(Combo1(pcmbBU).Text, 1) Then
                    SKIP_F = True
                End If
                '�i��
                wkHIN_GAI = Trim(xlSheet.Application.Cells(i, SN_HIN_GAI))
                '���t�G���[�`�F�b�N
                wkYOTEI_DT = Trim(xlSheet.Application.Cells(i, SN_YOTEI_DT))
                If Not SKIP_F Then
                    If Len(Trim(xlSheet.Application.Cells(i, SN_YOTEI_DT))) < 8 Then
                    Else
                        If Len(Trim(wkYOTEI_DT)) = 8 Then
                            wkYOTEI_DT = Mid(wkYOTEI_DT, 1, 4) & "/" & Mid(wkYOTEI_DT, 5, 2) & "/" & Mid(wkYOTEI_DT, 7, 2)
                        End If
                        If Not IsDate(wkYOTEI_DT) Then
                            wkYOTEI_DT = ""
                        Else
                            wkYOTEI_DT = wkYOTEI_DT
                        End If
                    End If
                    '�\�萔�G���[�`�F�b�N
                    If Trim(xlSheet.Application.Cells(i, SN_YOTEI_QTY)) = "" Then
                        wkYOTEI_Qty = 0
                    Else
                        If Not IsNumeric(Trim(xlSheet.Application.Cells(i, SN_YOTEI_QTY))) Then
                            SKIP_F = True
                        Else
                            wkYOTEI_Qty = CLng(Trim(xlSheet.Application.Cells(i, SN_YOTEI_QTY)))
                        End If
                    End If
                    If Not SKIP_F Then
                        Row = Row + 1
                        PLN_Y_NYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
                        PLN_Y_NYUKA(Row, colNO) = Row
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo1(pcmbBU), 1))
                        Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, wkHIN_GAI)
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                Call UniCode_Conv(K2_ITEM.JGYOBU, Right(Combo1(pcmbBU), 1))
                                Call UniCode_Conv(K2_ITEM.NAIGAI, "1")
                                Call UniCode_Conv(K2_ITEM.HIN_NAI, wkHIN_GAI)
                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                                Select Case sts
                                    Case BtNoErr
                                    Case BtErrKeyNotFound
                                        Call UniCode_Conv(ITEMREC.HIN_GAI, "")
                                        Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                                    Case Else
                                        Call Input_UnLock
                                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                        Exit Function
                                End Select
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                Exit Function
                        End Select
                        PLN_Y_NYUKA(Row, colHIN_GAI) = wkHIN_GAI
                        PLN_Y_NYUKA(Row, colHIN_NAI) = Trim(StrConv(ITEMREC.HIN_NAI, vbUnicode))
                        PLN_Y_NYUKA(Row, colHIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                        If IsDate(wkYOTEI_DT) Then
                            PLN_Y_NYUKA(Row, colN_YOTEI_DT) = DateAdd("d", SN_DAY, wkYOTEI_DT)
                        Else
                            PLN_Y_NYUKA(Row, colN_YOTEI_DT) = ""
                        End If
                        PLN_Y_NYUKA(Row, colN_YOTEI_QTY) = Format(wkYOTEI_Qty, "#,###")
                        PLN_Y_NYUKA(Row, colSHIIRE) = ""
                    End If
                End If
            End If
        End If
    Next i
    
    On Error GoTo 0


    Set TDBGrid1.Array = PLN_Y_NYUKA
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst






    lblDisp_Count.Caption = Format(Row, "#0") & "��"



    xlApp.DisplayAlerts = False

    xlBook.Close False
    xlApp.Quit 'EXCEL�����
    Set xlApp = Nothing




hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i���p���ח\��t�@�C��[���є[����]�@�\�������I���I�I", Me.hwnd, 0)



    Call Input_UnLock


    List_Disp_SN_Proc = False
    Exit Function

Error_Proc:
    

    Select Case Err.Number
        
        '52 �t�@�C�����܂��͔ԍ����s���ł��B
        '53 �t�@�C����������܂���B
        '54 �t�@�C�� ���[�h���s���ł��B
        '55 �t�@�C���͊��ɊJ����Ă��܂��B
        '57 �f�o�C�X I/O �G���[�ł��B
        '59 ���R�[�h������v���܂���B
        '61 �f�B�X�N�̋󂫗e�ʂ��s�����Ă��܂��B
        '62 �t�@�C���ɂ���ȏ�f�[�^������܂���B
        '63 ���R�[�h�ԍ����s���ł��B
        '68 �f�o�C�X����������Ă��܂���B
        '70 �������݂ł��܂���B
        '71 �f�B�X�N����������Ă��܂���B
        '75 �p�X���������ł��B
        '76 �p�X��������܂���B
        Case 52, 53, 54, 55, 57, 59, 61, 62, 63, 68, 70, 71, 75, 76
            
            
            MsgBox "�w��̃t�@�C����������܂���B" & Chr(13) & Chr(10) & "�������t�@�C��������͂��Ă��������B"
            
            
            
            List_Disp_SN_Proc = False      '





    '2011.12.03
        Case 13
        
            MsgBox "�Ǎ��ݑΏۂ�EXCEL�f�[�^�Ɉُ킪�L��܂��B���e���m�F��A�Ď��s���Ă��������B"
            
            xlApp.DisplayAlerts = False
            xlBook.Close False
            xlApp.Quit 'EXCEL�����
            Set xlApp = Nothing
            
            
            
            
            
            
            List_Disp_SN_Proc = False      '
            


        Case Else
            MsgBox Err.Description
    
    '2011.12.03
    End Select
    
    On Error GoTo 0
    
    Call Input_UnLock

End Function

Private Function List_Disp_NA_Proc() As Integer
'----------------------------------------------------------------------------
'                   �u���i���p���ח\��t�@�C���v�Ǎ��ݏ��� �[���񓚏�(PPSC) 2012.02.13
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
    
Dim INS_NOW         As String * 14
    
    
    

Dim xlApp           As Object
Dim xlBook          As Object
Dim xlSheet         As Object

Dim Row             As Long

Dim END_GYO         As Integer
Dim SKIP_F          As Boolean


Dim wkHIN_GAI       As String * 20
Dim wkYOTEI_DT      As String
Dim wkYOTEI_Qty     As Long

Dim i               As Long
Dim j               As Long

    List_Disp_NA_Proc = True

    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i���p���ח\��t�@�C��[�[���񓚏�(PPSC)]�@�\�������J�n�I�I", Me.hwnd, 0)





    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlApp.Worksheets(1)
    
    
    On Error GoTo Error_Proc

    xlApp.Workbooks.Open (Text1.Text), ReadOnly:=True, UpdateLinks:=0

    On Error GoTo 0
    
    
    
    Set PLN_Y_NYUKA = Nothing
    
    Row = Min_Row - 1
    lblDisp_Count.Caption = ""
    
    END_GYO = 0
    
    On Error GoTo Error_Proc
    
    For i = 1 To 1048576
        SKIP_F = False
        If Trim(xlSheet.Application.Cells(i, NA_HIN_GAI)) = "" And _
            Trim(xlSheet.Application.Cells(i, NA_YOTEI_DT)) = "" And _
            Trim(xlSheet.Application.Cells(i, NA_YOTEI_QTY)) = "" Then
            SKIP_F = True
            END_GYO = END_GYO + 1
            If END_GYO > 3 Then
                Exit For
            End If
        Else
            If Trim(xlSheet.Application.Cells(i, NA_HIN_GAI)) = "" Or _
                Trim(xlSheet.Application.Cells(i, NA_YOTEI_DT)) = "" Or _
                (Trim(xlSheet.Application.Cells(i, NA_YOTEI_QTY)) <> "" And Not IsNumeric(Trim(xlSheet.Application.Cells(i, NA_YOTEI_QTY)))) Then
            Else
                END_GYO = 0
                '�i��
                wkHIN_GAI = Trim(xlSheet.Application.Cells(i, NA_HIN_GAI))
                '���t�G���[�`�F�b�N
                wkYOTEI_DT = Trim(xlSheet.Application.Cells(i, NA_YOTEI_DT))
                For j = 0 To UBound(NA_JYOGAI_TBL)
                    If wkYOTEI_DT = NA_JYOGAI_TBL(j) Then
                        SKIP_F = True
                        Exit For
                    End If
                Next j
                If Not SKIP_F Then
                    If Len(Trim(xlSheet.Application.Cells(i, NA_YOTEI_DT))) < 8 Then
                    Else
                        If Len(Trim(wkYOTEI_DT)) = 8 Then
                            wkYOTEI_DT = Mid(wkYOTEI_DT, 1, 4) & "/" & Mid(wkYOTEI_DT, 5, 2) & "/" & Mid(wkYOTEI_DT, 7, 2)
                        End If
                        If Not IsDate(wkYOTEI_DT) Then
                            wkYOTEI_DT = ""
                        Else
                            wkYOTEI_DT = wkYOTEI_DT
                        End If
                    End If
                    '�\�萔�G���[�`�F�b�N
                    If Trim(xlSheet.Application.Cells(i, NA_YOTEI_QTY)) = "" Then
                        SKIP_F = True
                    Else
                        If Not IsNumeric(Trim(xlSheet.Application.Cells(i, NA_YOTEI_QTY))) Then
                            SKIP_F = True
                        Else
                            wkYOTEI_Qty = CLng(Trim(xlSheet.Application.Cells(i, NA_YOTEI_QTY)))
                        End If
                    End If
                    If Not SKIP_F Then
                        Row = Row + 1
                        PLN_Y_NYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
                        PLN_Y_NYUKA(Row, colNO) = Row
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Right(Combo1(pcmbBU), 1))
                        Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, wkHIN_GAI)
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                Call UniCode_Conv(K2_ITEM.JGYOBU, Right(Combo1(pcmbBU), 1))
                                Call UniCode_Conv(K2_ITEM.NAIGAI, "1")
                                Call UniCode_Conv(K2_ITEM.HIN_NAI, wkHIN_GAI)
                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                                Select Case sts
                                    Case BtNoErr
                                    Case BtErrKeyNotFound
                                        Call UniCode_Conv(ITEMREC.HIN_GAI, "")
                                        Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                                    Case Else
                                        Call Input_UnLock
                                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                        Exit Function
                                End Select
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                Exit Function
                        End Select
                        PLN_Y_NYUKA(Row, colHIN_GAI) = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
                        PLN_Y_NYUKA(Row, colHIN_NAI) = wkHIN_GAI
                        PLN_Y_NYUKA(Row, colHIN_NAME) = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                        If IsDate(wkYOTEI_DT) Then
                            PLN_Y_NYUKA(Row, colN_YOTEI_DT) = DateAdd("d", NA_DAY, wkYOTEI_DT)
                        Else
                            PLN_Y_NYUKA(Row, colN_YOTEI_DT) = ""
                        End If
                        PLN_Y_NYUKA(Row, colN_YOTEI_QTY) = Format(wkYOTEI_Qty, "#,##0")
                        PLN_Y_NYUKA(Row, colSHIIRE) = ""
                    End If
                End If
            End If
        End If
    Next i
    
    On Error GoTo 0


    Set TDBGrid1.Array = PLN_Y_NYUKA
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst






    lblDisp_Count.Caption = Format(Row, "#0") & "��"



    xlApp.DisplayAlerts = False

    xlBook.Close False
    xlApp.Quit 'EXCEL�����
    Set xlApp = Nothing




hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i���p���ח\��t�@�C��[�[���񓚏�(PPSC)]�@�\�������I���I�I", Me.hwnd, 0)



    Call Input_UnLock


    List_Disp_NA_Proc = False
    Exit Function

Error_Proc:
    

    Select Case Err.Number
        
        '52 �t�@�C�����܂��͔ԍ����s���ł��B
        '53 �t�@�C����������܂���B
        '54 �t�@�C�� ���[�h���s���ł��B
        '55 �t�@�C���͊��ɊJ����Ă��܂��B
        '57 �f�o�C�X I/O �G���[�ł��B
        '59 ���R�[�h������v���܂���B
        '61 �f�B�X�N�̋󂫗e�ʂ��s�����Ă��܂��B
        '62 �t�@�C���ɂ���ȏ�f�[�^������܂���B
        '63 ���R�[�h�ԍ����s���ł��B
        '68 �f�o�C�X����������Ă��܂���B
        '70 �������݂ł��܂���B
        '71 �f�B�X�N����������Ă��܂���B
        '75 �p�X���������ł��B
        '76 �p�X��������܂���B
        Case 52, 53, 54, 55, 57, 59, 61, 62, 63, 68, 70, 71, 75, 76
            
            
            MsgBox "�w��̃t�@�C����������܂���B" & Chr(13) & Chr(10) & "�������t�@�C��������͂��Ă��������B"
            
            
            
            List_Disp_NA_Proc = False      '





    '2011.12.03
        Case 13
        
            MsgBox "�Ǎ��ݑΏۂ�EXCEL�f�[�^�Ɉُ킪�L��܂��B���e���m�F��A�Ď��s���Ă��������B"
            
            xlApp.DisplayAlerts = False
            xlBook.Close False
            xlApp.Quit 'EXCEL�����
            Set xlApp = Nothing
            
            
            
            
            
            
            List_Disp_NA_Proc = False      '
            


        Case Else
            MsgBox Err.Description
    
    '2011.12.03
    End Select
    
    On Error GoTo 0
    
    Call Input_UnLock

End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------
Dim i   As Integer


    PLN00101.MousePointer = vbHourglass

    Call Ctrl_Lock(PLN00101)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(PLN00101)


    PLN00101.MousePointer = vbDefault

End Sub



Private Sub Bu_Set_Proc()
'----------------------------------------------------------------------------
'                   ��ʍ��ځi�a�t�j�̃Z�b�g
'----------------------------------------------------------------------------
Dim i   As Integer




    Combo1(pcmbBU).Clear


    



    For i = 0 To UBound(JGYOBU_T)
            
        Combo1(pcmbBU).AddItem JGYOBU_T(i).NAME & "          " & JGYOBU_T(i).CODE
            
            
    Next i

    Combo1(pcmbBU).ListIndex = 0
End Sub


Private Sub Data_Kb_Set_Proc()
'----------------------------------------------------------------------------
'                   ��ʍ��ځi�f�[�^�敪�j�̃Z�b�g
'----------------------------------------------------------------------------
Dim i   As Integer




    Combo1(pcmbDATA_KB).Clear


    



    For i = 0 To UBound(DATA_KB) - 1 Step 2
            
        Combo1(pcmbDATA_KB).AddItem DATA_KB(i + 1) & "                    " & DATA_KB(i)
            
            
    Next i

    Combo1(pcmbDATA_KB).ListIndex = 0

End Sub

