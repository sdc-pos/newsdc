VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PLN00701 
   Caption         =   "[商品化計画システム]資材所要量確認"
   ClientHeight    =   9510
   ClientLeft      =   2025
   ClientTop       =   -4500
   ClientWidth     =   15150
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  '手動
   ScaleHeight     =   9510
   ScaleWidth      =   15150
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame1 
      Caption         =   "表示対象"
      Height          =   735
      Left            =   5400
      TabIndex        =   7
      Top             =   120
      Width           =   5295
      Begin VB.CheckBox Check1 
         Caption         =   "外装"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "個装"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "構成"
         Height          =   255
         Index           =   4
         Left            =   4200
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "同梱"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "資材"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "展　開"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   6
      ToolTipText     =   "商品化予定をもとに展開処理を行います"
      Top             =   0
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   372
      Index           =   0
      Left            =   2400
      TabIndex        =   3
      Top             =   480
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "表　示"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "展開済みの情報を再表示します"
      Top             =   0
      Width           =   1380
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   8175
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   14420
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "種別ｺｰﾄﾞ"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "種別"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "事業部"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "品番"
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
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   714
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1561"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1455"
      Splits(0)._ColumnProps(4)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(5)=   "Column(0).FetchStyle=1"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=714"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=609"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=8192"
      Splits(0)._ColumnProps(11)=   "Column(1).FetchStyle=1"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1217"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1111"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=8196"
      Splits(0)._ColumnProps(17)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(18)=   "Column(2).FetchStyle=1"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=1376"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1270"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=8192"
      Splits(0)._ColumnProps(24)=   "Column(3).FetchStyle=1"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=1111"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1005"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(30)=   "Column(4).FetchStyle=1"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=1482"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1376"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=8194"
      Splits(0)._ColumnProps(36)=   "Column(5).FetchStyle=1"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
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
      _StyleDefs(5)   =   ":id=0,.fontname=ＭＳ ゴシック"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ＭＳ ゴシック"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=67,.parent=1,.bgcolor=&HFFFF00&,.bold=0,.fontsize=975"
      _StyleDefs(25)  =   ":id=67,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(26)  =   ":id=67,.fontname=ＭＳ ゴシック"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=68,.parent=2,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(29)  =   ":id=68,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(30)  =   ":id=68,.fontname=ＭＳ ゴシック"
      _StyleDefs(31)  =   "Splits(0).FooterStyle:id=69,.parent=3"
      _StyleDefs(32)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
      _StyleDefs(33)  =   "Splits(0).SelectedStyle:id=72,.parent=6"
      _StyleDefs(34)  =   "Splits(0).EditorStyle:id=71,.parent=7"
      _StyleDefs(35)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
      _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=74,.parent=9,.bgcolor=&HFFFFFF&"
      _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=75,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
      _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=24,.parent=67"
      _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=68"
      _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=69"
      _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=71"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=94,.parent=67,.alignment=0,.locked=-1,.bold=0"
      _StyleDefs(45)  =   ":id=94,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(46)  =   ":id=94,.fontname=ＭＳ ゴシック"
      _StyleDefs(47)  =   "Splits(0).Columns(1).HeadingStyle:id=91,.parent=68,.bold=0,.fontsize=825"
      _StyleDefs(48)  =   ":id=91,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(49)  =   ":id=91,.fontname=ＭＳ ゴシック"
      _StyleDefs(50)  =   "Splits(0).Columns(1).FooterStyle:id=92,.parent=69"
      _StyleDefs(51)  =   "Splits(0).Columns(1).EditorStyle:id=93,.parent=71"
      _StyleDefs(52)  =   "Splits(0).Columns(2).Style:id=20,.parent=67,.locked=-1"
      _StyleDefs(53)  =   "Splits(0).Columns(2).HeadingStyle:id=17,.parent=68"
      _StyleDefs(54)  =   "Splits(0).Columns(2).FooterStyle:id=18,.parent=69"
      _StyleDefs(55)  =   "Splits(0).Columns(2).EditorStyle:id=19,.parent=71"
      _StyleDefs(56)  =   "Splits(0).Columns(3).Style:id=102,.parent=67,.alignment=0,.locked=-1,.bold=0"
      _StyleDefs(57)  =   ":id=102,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(58)  =   ":id=102,.fontname=ＭＳ ゴシック"
      _StyleDefs(59)  =   "Splits(0).Columns(3).HeadingStyle:id=99,.parent=68"
      _StyleDefs(60)  =   "Splits(0).Columns(3).FooterStyle:id=100,.parent=69"
      _StyleDefs(61)  =   "Splits(0).Columns(3).EditorStyle:id=101,.parent=71"
      _StyleDefs(62)  =   "Splits(0).Columns(4).Style:id=98,.parent=67,.alignment=1,.bold=0,.fontsize=825"
      _StyleDefs(63)  =   ":id=98,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(64)  =   ":id=98,.fontname=ＭＳ ゴシック"
      _StyleDefs(65)  =   "Splits(0).Columns(4).HeadingStyle:id=95,.parent=68"
      _StyleDefs(66)  =   "Splits(0).Columns(4).FooterStyle:id=96,.parent=69"
      _StyleDefs(67)  =   "Splits(0).Columns(4).EditorStyle:id=97,.parent=71"
      _StyleDefs(68)  =   "Splits(0).Columns(5).Style:id=16,.parent=67,.alignment=1,.locked=-1,.bold=0"
      _StyleDefs(69)  =   ":id=16,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(70)  =   ":id=16,.fontname=ＭＳ ゴシック"
      _StyleDefs(71)  =   "Splits(0).Columns(5).HeadingStyle:id=13,.parent=68"
      _StyleDefs(72)  =   "Splits(0).Columns(5).FooterStyle:id=14,.parent=69"
      _StyleDefs(73)  =   "Splits(0).Columns(5).EditorStyle:id=15,.parent=71"
      _StyleDefs(74)  =   "Named:id=33:Normal"
      _StyleDefs(75)  =   ":id=33,.parent=0"
      _StyleDefs(76)  =   "Named:id=34:Heading"
      _StyleDefs(77)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(78)  =   ":id=34,.wraptext=-1"
      _StyleDefs(79)  =   "Named:id=35:Footing"
      _StyleDefs(80)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(81)  =   "Named:id=36:Selected"
      _StyleDefs(82)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(83)  =   "Named:id=37:Caption"
      _StyleDefs(84)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(85)  =   "Named:id=38:HighlightRow"
      _StyleDefs(86)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(87)  =   "Named:id=39:EvenRow"
      _StyleDefs(88)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(89)  =   "Named:id=40:OddRow"
      _StyleDefs(90)  =   ":id=40,.parent=33"
      _StyleDefs(91)  =   "Named:id=41:RecordSelector"
      _StyleDefs(92)  =   ":id=41,.parent=34"
      _StyleDefs(93)  =   "Named:id=42:FilterBar"
      _StyleDefs(94)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終 了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   0
      ToolTipText     =   "資材所要量確認画面を終了します"
      Top             =   0
      Width           =   1380
   End
   Begin VB.Label lblLast_DATETIME 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   10920
      TabIndex        =   14
      Top             =   600
      Width           =   90
   End
   Begin VB.Label lblLAST_DATE 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   10920
      TabIndex        =   13
      Top             =   240
      Width           =   90
   End
   Begin VB.Label Label1 
      Caption         =   "から2週間"
      Height          =   255
      Index           =   2
      Left            =   3840
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "展開対象開始日"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "表示"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "展開"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   2
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "PLN00701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Const ptxStart_Date% = 0

Private Const chkSHIZAI% = 0
Private Const chkKOSOU% = 1
Private Const chkGAISOU% = 2
Private Const chkDOUKON% = 3
Private Const chkKOUSEI% = 4


Private ZAIKO_RIREKI    As New XArrayDB
' 実行時に新しいスタイルを定義するためのオブジェクトです。
Private Rstyle_Red          As New style
Private Rstyle_Black        As New style


Private Rstyle_Blue         As New style
Private Rstyle_White        As New style
    
                                


Private Const Min_Row% = 1              '最小行数
Private Max_Row             As Integer      'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Max_Col As Long                 '最大列数

Private Const colSYUBETSU_CODE% = 0     '種別

Private Const colSYUBETSU% = 1          '種別
Private Const colJGYOBU% = 2            '事業部


Private Const colHIN_GAI% = 3           '品番
Private Const colTITLE% = 4             '項目
Private Const colDAY% = 5               '日別



Private List_Week   As Long             '表示するn週間

Private SHIZAI_TBL  As Variant          '資材分

Private KOSOU_TBL   As Variant          '個装分

Private GAISOU_TBL  As Variant          '外装分

Private DOUKON_TBL  As Variant          '同梱分

Private KOUSEI_TBL  As Variant          '構成分


Private LAST_TENKAI_DateTime _
                    As String


Private DATE_TBL()  As String * 10


Private Type Z_RIREKI_Tbl_tag
    SYUBETSU        As String * 2
    JGYOBU          As String * 1
    NAIGAI          As String * 1
    HIN_GAI         As String * 20
    
    Start_Zaiko_QTY As Long
    
    SYOUHI_QTY()    As Long
    NYUKA_QTY()     As Long
    ZAIKO_QTY()     As Long

    DATA_KBN        As String * 1
End Type


Private Z_RIREKI_Tbl()  As Z_RIREKI_Tbl_tag

Private LAST_DATETIME   As String

Private LAST_Date       As String


Private wkBackColor     As Boolean
Private svBookMark      As Variant
Private List_Disp_F     As Boolean


Private KOSOU_CODE      As String * 2
Private GAISOU_CODE     As String * 2


Private Const LAST_UPDATE_DAY$ = "[PLN0070] 2012.01.07 14:00"

Private Sub Command1_Click(Index As Integer)

Dim sWk         As String
Dim i           As Long
Dim j           As Long


    Select Case Index



        Case 0          '読込み
            
            '取込みﾃﾞｰﾀ表示
            If List_Disp_Proc(Left(LAST_DATETIME, 10)) Then
                Unload Me
            End If

        Case 1          '展開


            If Not IsDate(Text1(ptxStart_Date).Text) Then
                MsgBox "入力した項目はエラーです。（展開対象開始日）"
                Text1(ptxStart_Date).SetFocus
                Exit Sub
            End If

            If Text1(ptxStart_Date).Text < Format(Now, "YYYY/MM/DD") Then
                MsgBox "入力した項目はエラーです。（展開対象開始日 < 本日）"
                Text1(ptxStart_Date).SetFocus
                Exit Sub
            End If
            If Update_Proc() Then
                Unload Me
            End If
            '取込みﾃﾞｰﾀ表示
            If List_Disp_Proc(Text1(ptxStart_Date).Text) Then
                Unload Me
            End If

        Case 2          '終了

            Unload Me
    End Select



    Command1(Index).SetFocus


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    Select Case KeyCode
        
        Case vbKeyF12
            Unload Me
    End Select
    
    
    

End Sub

Private Sub Form_Load()


Dim c       As String * 128



    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[商品化計画システム]資材所要量確認画面", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)



    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                '表示期間取り込み
    If GetIni(App.EXEName, "WEEK", App.EXEName, c) Then
        List_Week = 2
    Else
        If Not IsNumeric(Trim(c)) Then
            List_Week = 2
        Else
            If Val(Trim(c)) < 1 Then
                List_Week = 1
            Else
                If Val(Trim(c)) > 8 Then
                    List_Week = 8
                Else
                    List_Week = Val(Trim(c))
                End If
            End If
        End If
    End If
    Call List_Make_Proc
                                '前回展開年月日時分秒
    If GetIni(App.EXEName, "LAST_DATETIME", App.EXEName, c) Then
        lblLast_DATETIME.Caption = ""
        LAST_DATETIME = ""
    Else
        lblLast_DATETIME.Caption = "　前回展開日時：" & Trim(c) & "現在"
        LAST_DATETIME = Trim(c)
    End If
                                
                                
                                '前回展開開始年月日
    If GetIni(App.EXEName, "LAST_DATE", App.EXEName, c) Then
        lblLAST_DATE.Caption = ""
        lblLAST_DATE.Caption = ""
    Else
        lblLAST_DATE.Caption = "前回展開開始日：" & Trim(c)
        LAST_Date = Trim(c)
    End If
                                
                                
                                '資材として表示する(*=終端)
    If GetIni(App.EXEName, "SHIZAI", App.EXEName, c) Then
        c = "*"
    End If
    SHIZAI_TBL = Split(Trim(c), ",", -1)
                                '個装として表示する(個装欄登録分含む)
    If GetIni(App.EXEName, "KOSOU", App.EXEName, c) Then
        c = "*"
    End If
    KOSOU_TBL = Split(Trim(c), ",", -1)
                                '外装として表示する(外装欄登録分含む)
    If GetIni(App.EXEName, "GAISOU", App.EXEName, c) Then
        c = "*"
    End If
    GAISOU_TBL = Split(Trim(c), ",", -1)
                                '同梱として表示する
    If GetIni(App.EXEName, "DOUKON", App.EXEName, c) Then
        c = "*"
    End If
    DOUKON_TBL = Split(Trim(c), ",", -1)
                                '構成として表示する
    If GetIni(App.EXEName, "KOUSEI", App.EXEName, c) Then
        c = "*"
    End If
    KOUSEI_TBL = Split(Trim(c), ",", -1)


                                '個装種別ｺｰﾄﾞとして表示する
    If GetIni(App.EXEName, "KOSOU_CODE", App.EXEName, c) Then
        KOSOU_CODE = ""
    Else
        KOSOU_CODE = Trim(c)
    End If
                                '外装種別ｺｰﾄﾞとして表示する
    If GetIni(App.EXEName, "GAISOU_CODE", App.EXEName, c) Then
        GAISOU_CODE = ""
    Else
        GAISOU_CODE = Trim(c)
    End If




    PLN00701.Caption = PLN00701.Caption & " " & LAST_UPDATE_DAY
    
''    If Trim(Text1(ptxLAST_DateTime).Text) = "" Then
        Text1(ptxStart_Date).Text = Format(Now, "YYYY/MM/DD")
''    Else
''        Text1(ptxStart_Date).Text = Format(Left(Text1(ptxLAST_DateTime).Text, 10), "YYYY/MM/DD")
''    End If
    Label1(2).Caption = "から" & List_Week & "週間"
    
    Check1(chkSHIZAI).Value = vbChecked
    Check1(chkKOSOU).Value = vbChecked
    Check1(chkGAISOU).Value = vbChecked


    Check1(chkDOUKON).Value = vbUnchecked
    Check1(chkKOUSEI).Value = vbUnchecked


                                '商品化予定ファイル
    If PLN_S_YOTEI_Open(BtOpenRead) Then
        Unload Me
    End If
                                '品目マスタ
    If ITEM_Open(BtOpenRead) Then
        Unload Me
    End If
                                
                                
                                '構成マスタ
    If P_COMPO_Open(BtOpenRead) Then
        Unload Me
    End If
                                'コードマスタ
    If P_CODE_Open(BtOpenRead) Then
        Unload Me
    End If
                                '資材注文ﾃﾞｰﾀ
    If P_SHORDER_Open(BtOpenRead) Then
        Unload Me
    End If
                                '在庫データ
    If ZAIKO_Open(BtOpenRead) Then
        Unload Me
    End If
                                '資材所要量確認画面中間ファイル
    If PLN_tmpZaiko_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '資材所要量ファイル
    If PLN_tmpP_COMP_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc
                                
                                
' 新しいスタイルを定義します。
                                    '黒
    Set Rstyle_Black = TDBGrid1.Styles.Add("Rstyle_Black")
    Rstyle_Black.ForeColor = &H0                '文字色＝黒
                                    '赤
    Set Rstyle_Red = TDBGrid1.Styles.Add("Rstyle_Red")
    Rstyle_Red.ForeColor = &HFF                 '文字色＝赤
                                
                                
    Set Rstyle_Blue = TDBGrid1.Styles.Add("Rstyle_Blue")
    Rstyle_Blue.BackColor = &HFFFFC0            '背景色＝水色
    
                                
    Set Rstyle_White = TDBGrid1.Styles.Add("Rstyle_White")
    Rstyle_White.BackColor = &H80000005         '背景色＝白
                                
    wkBackColor = False
    svBookMark = -1
    List_Disp_F = False
    
    Load PLN00702
                                
    Show


    If Trim(LAST_Date) <> "" Then
        If List_Disp_Proc(LAST_Date) Then
            Unload Me
        End If
    End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
    
    
Dim sts     As Integer
    
    sts = BTRV(BtOpClose, PLN_tmpZaiko_POS, PLN_tmpZaikoREC, Len(PLN_tmpZaikoREC), K0_PLN_tmpZaiko, Len(K0_PLN_tmpZaiko), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材所要量確認画面中間ファイル")
        End If
    End If

    
    
    sts = BTRV(BtOpReset, PLN_tmpZaiko_POS, PLN_tmpZaikoREC, Len(PLN_tmpZaikoREC), K0_PLN_tmpZaiko, Len(K0_PLN_tmpZaiko), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If


    
    Set PLN00701 = Nothing



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



Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   「資材所要量確認画面中間ファイル」作成処理
'----------------------------------------------------------------------------

Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer
    
Dim c               As String * 128
    
Dim Upd_Com         As Integer
Dim SKIP_FLG        As Integer
    
Dim List_Day        As Long
Dim End_Date        As String

Dim Yobi            As Integer
Dim Yobi_NAME       As String


    
Dim INS_NOW         As String * 14

Dim Row             As Long
Dim i               As Long
Dim j               As Long

Dim Ing_Date        As String
Dim Start_Date      As String

Dim SHIMUKE_CODE    As String * 2

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long


Dim svSYUBETSU      As String * 2       '種別
Dim svJGYOBU        As String * 1       '事業部区分
Dim svNAIGAI        As String * 1       '国内外
Dim svHIN_GAI       As String * 20      '品番（外部）
Dim svDATA_KBN      As String * 1       'ﾃﾞｰﾀ区分

Dim Fast_Flg        As Boolean

Dim ST_ZAIKO_QTY    As Long             '開始時在庫数



    Update_Proc = True
    
    Call Input_Lock


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "資材所要量中間ファイル作成！！[全件削除開始]", Me.hwnd, 0)

    sts = BTRV(BtOpClose, PLN_tmpZaiko_POS, PLN_tmpZaikoREC, Len(PLN_tmpZaikoREC), K0_PLN_tmpZaiko, Len(K0_PLN_tmpZaiko), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材所要量確認画面中間ファイル")
        End If
    End If

    sts = BTRV(BtOpClose, PLN_tmpP_COMP_POS, PLN_tmpP_COMP_REC, Len(PLN_tmpP_COMP_REC), K0_PLN_tmpP_COMP, Len(K0_PLN_tmpP_COMP), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材所要量中間ファイル")
        End If
    End If
    On Error Resume Next
    sts = GetIni("FILE", PLN_tmpZaiko_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [PLN_tmpZaiko]読み込みエラー")
        Exit Function
    End If
    Kill RTrim(c)
    sts = GetIni("FILE", PLN_tmpP_COMP_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [PLN_PLN_tmpP_COMP]読み込みエラー")
        Exit Function
    End If
    Kill RTrim(c)
    On Error GoTo 0

                                '資材所要量確認画面中間ファイル
    If PLN_tmpZaiko_Open(BtOpenNomal) Then
        Exit Function
    End If
                                '資材所要量ファイル
    If PLN_tmpP_COMP_Open(BtOpenNomal) Then
        Unload Me
    End If


                                    
                                    
    com = BtOpGetFirst
                                    
                                    
    Do
        DoEvents
    
    
        sts = BTRV(com, PLN_tmpZaiko_POS, PLN_tmpZaikoREC, Len(PLN_tmpZaikoREC), K0_PLN_tmpZaiko, Len(K0_PLN_tmpZaiko), 0)
        Select Case sts
            Case BtNoErr
                
            
            Case BtErrEOF
                Exit Do
            Case Else
                
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "資材所要量確認画面中間ファイル作成　異常停止！！[全件削除処理]", Me.hwnd, 0)
                
                Call Input_UnLock
                Call File_Error(sts, com, "資材所要量確認画面中間ファイル")
                Exit Function
        
        End Select
    
    
        sts = BTRV(BtOpDelete, PLN_tmpZaiko_POS, PLN_tmpZaikoREC, Len(PLN_tmpZaikoREC), K0_PLN_tmpZaiko, Len(K0_PLN_tmpZaiko), 0)
        Select Case sts
            Case BtNoErr
                
            
            Case Else
                
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "資材所要量確認画面中間ファイル作成　異常停止！！[全件削除処理]", Me.hwnd, 0)
                Call Input_UnLock
                Call File_Error(sts, BtOpDelete, "資材所要量確認画面中間ファイル")
                Exit Function
        
        End Select
    
        com = BtOpGetNext
    
    Loop
                                    
                                    
    com = BtOpGetFirst
                                    
                                    
    Do
        DoEvents
    
    
        sts = BTRV(com, PLN_tmpP_COMP_POS, PLN_tmpP_COMP_REC, Len(PLN_tmpP_COMP_REC), K0_PLN_tmpP_COMP, Len(K0_PLN_tmpP_COMP), 0)
        Select Case sts
            Case BtNoErr
                
            
            Case BtErrEOF
                Exit Do
            Case Else
                
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "資材所要量中間ファイル作成　異常停止！！[全件削除処理]", Me.hwnd, 0)
                
                Call Input_UnLock
                Call File_Error(sts, com, "資材所要量中間ファイル")
                Exit Function
        
        End Select
    
    
        sts = BTRV(BtOpDelete, PLN_tmpP_COMP_POS, PLN_tmpP_COMP_REC, Len(PLN_tmpP_COMP_REC), K0_PLN_tmpP_COMP, Len(K0_PLN_tmpP_COMP), 0)
        Select Case sts
            Case BtNoErr
                
            
            Case Else
                
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "資材所要量中間ファイル作成　異常停止！！[全件削除処理]", Me.hwnd, 0)
                Call Input_UnLock
                Call File_Error(sts, BtOpDelete, "資材所要量中間ファイル")
                Exit Function
        
        End Select
    
        com = BtOpGetNext
    
    Loop
                                    
                                    
                                    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "資材所要量確認画面中間ファイル作成！！[展開処理開始]", Me.hwnd, 0)
                                    
    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")
    List_Day = List_Week * 7
    
    
    Start_Date = Text1(ptxStart_Date).Text
    If Format(Text1(ptxStart_Date).Text, "YYYY/MM/DD") > Format(Now, "YYYY/MM/DD") Then
        List_Day = List_Day + DateDiff("d", Format(Now, "YYYY/MM/DD"), Text1(ptxStart_Date).Text)
        Start_Date = Format(Now, "YYYY/MM/DD")
    
    End If
    
    End_Date = DateAdd("d", List_Day, Start_Date)
                                    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<子部品展開>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                                    
    Call UniCode_Conv(K2_PLN_S_YOTEI.YOTEI_DT, Format(Now, "YYYYMMDD"))                     '商品化予定日付
    Call UniCode_Conv(K2_PLN_S_YOTEI.JGYOBU, "")                                            '事業部区分
    Call UniCode_Conv(K2_PLN_S_YOTEI.NAIGAI, "")                                            '国内外
    Call UniCode_Conv(K2_PLN_S_YOTEI.HIN_GAI, "")                                           '品番（外部）
    
    com = BtOpGetGreater
                                    
    Do
        DoEvents
    
        sts = BTRV(com, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K2_PLN_S_YOTEI, Len(K2_PLN_S_YOTEI), 2)
        Select Case sts
            Case BtNoErr
                
                If StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode) > Format(End_Date, "YYYYMMDD") Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "資材所要量確認画面中間ファイル作成　異常停止！！[展開処理]", Me.hwnd, 0)
                
                Call Input_UnLock
                Call File_Error(sts, com, "資材所要量確認画面中間ファイル")
                Exit Function
        
        End Select
    
        SKIP_FLG = False
        If Trim(StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode)) = "" Or Val(StrConv(PLN_S_YOTEI_R.YOTEI_QTY, vbUnicode)) = 0 Then
            SKIP_FLG = True
        End If
        If GetIni(App.EXEName, StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode), App.EXEName, c) Then
            SKIP_FLG = True
        Else
            SHIMUKE_CODE = Trim(c)
        End If
    
        If Not SKIP_FLG Then
            
            
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "資材所要量確認画面中間ファイル作成！！[展開処理開始][" & StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode) & "]", Me.hwnd, 0)
            
            
            
            Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, SHIMUKE_CODE)
            Call UniCode_Conv(K0_P_COMPO.JGYOBU, StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_P_COMPO.NAIGAI, StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_P_COMPO.HIN_GAI, StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode))
            Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "")
            Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
    
            com = BtOpGetGreaterEqual
            Do
                
                SKIP_FLG = False
                sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                Select Case sts
                    Case BtNoErr
                    
                        If SHIMUKE_CODE <> Trim(StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode)) Or _
                            Trim(StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode)) <> StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) Or _
                            Trim(StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode)) <> StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) Or _
                            Trim(StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode)) <> Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) Then
                            Exit Do
                        End If
                    
                    
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        
                        hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                            "資材所要量確認画面中間ファイル作成　異常停止！！[展開処理]", Me.hwnd, 0)
                        
                        Call Input_UnLock
                        Call File_Error(sts, com, "構成マスタ")
                        Exit Function
                
                End Select
            
        
                If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_HEAD Then
                    SKIP_FLG = True
                End If
        
                Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode))
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI Then
                            If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_OFF Then
                                SKIP_FLG = True
                            End If
                        End If
                    Case BtErrKeyNotFound
                        SKIP_FLG = True
                    Case Else
                        
                        hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                            "資材所要量確認画面中間ファイル作成　異常停止！！[展開処理]", Me.hwnd, 0)
                        
                        Call Input_UnLock
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                
                End Select
        
                If Not SKIP_FLG Then
'---------------------------------------------------------->>>>>    消費集計   <<<<<
                    Call UniCode_Conv(K0_PLN_tmpZaiko.SYUBETSU, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
                    Call UniCode_Conv(K0_PLN_tmpZaiko.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_PLN_tmpZaiko.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_PLN_tmpZaiko.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    Call UniCode_Conv(K0_PLN_tmpZaiko.RIREKI_DT, StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode))
                    
                    sts = BTRV(BtOpGetEqual, PLN_tmpZaiko_POS, PLN_tmpZaikoREC, Len(PLN_tmpZaikoREC), K0_PLN_tmpZaiko, Len(K0_PLN_tmpZaiko), 0)
                    Select Case sts
                        Case BtNoErr
                            Upd_Com = BtOpUpdate
                        Case BtErrKeyNotFound
                        
                            Call UniCode_Conv(PLN_tmpZaikoREC.SYUBETSU, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
                            Call UniCode_Conv(PLN_tmpZaikoREC.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                            Call UniCode_Conv(PLN_tmpZaikoREC.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                            Call UniCode_Conv(PLN_tmpZaikoREC.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                            Call UniCode_Conv(PLN_tmpZaikoREC.RIREKI_DT, StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode))
                            Call UniCode_Conv(PLN_tmpZaikoREC.DATA_KBN, StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode))
                            
                            
                            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                                    MI_QTY, _
                                                    StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode), _
                                                    StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode), _
                                                    StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)) Then
                                Call Input_UnLock
                                Exit Function
                            End If
            
                            Call UniCode_Conv(PLN_tmpZaikoREC.ST_ZAIKO_QTY, Format(SUMI_QTY + MI_QTY, "000000"))
                            Call UniCode_Conv(PLN_tmpZaikoREC.SYOUHI_QTY, "000000")
                            Call UniCode_Conv(PLN_tmpZaikoREC.NYUKA_QTY, "000000")
                            Call UniCode_Conv(PLN_tmpZaikoREC.ZAIKO_QTY, "000000")
                            Call UniCode_Conv(PLN_tmpZaikoREC.INS_TANTO, App.EXEName)
                            Call UniCode_Conv(PLN_tmpZaikoREC.Ins_DateTime, INS_NOW)
                            
                            Upd_Com = BtOpInsert
        
                        Case Else
                            
                            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                                "資材所要量確認画面中間ファイル作成　異常停止！！[展開処理]", Me.hwnd, 0)
                            
                            Call Input_UnLock
                            Call File_Error(sts, BtOpGetEqual, "資材所要量確認画面中間ファイル")
                            Exit Function
                    
                    End Select
                    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2011.12.01
                    For i = 0 To UBound(GAISOU_TBL)
                        If Trim(StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode)) = GAISOU_TBL(i) Then
                            Exit For
                        End If
                    Next i
                    
                    
                    If i <= UBound(GAISOU_TBL) Or StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_GAISOU Then
                        
                        '2012.01.07 ０割りチェック追加
                        If Val(StrConv(PLN_tmpP_COMP_REC.KO_QTY, vbUnicode)) = 0 Then
                        Else
                            Call UniCode_Conv(PLN_tmpZaikoREC.SYOUHI_QTY, Format((CLng(StrConv(PLN_tmpZaikoREC.SYOUHI_QTY, vbUnicode))) + _
                                                                                            ToRoundUp(CCur(CLng(StrConv(PLN_S_YOTEI_R.YOTEI_QTY, vbUnicode))) / _
                                                                                            CCur(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))), "000000")))
                        End If
                    Else
                        Call UniCode_Conv(PLN_tmpZaikoREC.SYOUHI_QTY, Format(CLng(StrConv(PLN_tmpZaikoREC.SYOUHI_QTY, vbUnicode)) + _
                                                                                        CLng(StrConv(PLN_S_YOTEI_R.YOTEI_QTY, vbUnicode)) * _
                                                                                        CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "000000"))
                    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2011.12.01
                    
                    sts = BTRV(Upd_Com, PLN_tmpZaiko_POS, PLN_tmpZaikoREC, Len(PLN_tmpZaikoREC), K0_PLN_tmpZaiko, Len(K0_PLN_tmpZaiko), 0)
                    Select Case sts
                        Case BtNoErr
                
                        Case Else
                
                
                            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                                "資材所要量確認画面中間ファイル作成　異常停止！！[展開処理]", Me.hwnd, 0)
                            
                            Call Input_UnLock
                            Call File_Error(sts, Upd_Com, "資材所要量確認画面中間ファイル")
                            Exit Function
                
                
                
                    End Select
'---------------------------------------------------------->>>>>    資材所要量中間ファイル作成   <<<<<
                    Call UniCode_Conv(K0_PLN_tmpP_COMP.JGYOBU, StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_PLN_tmpP_COMP.NAIGAI, StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_PLN_tmpP_COMP.HIN_GAI, StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode))
                    Call UniCode_Conv(K0_PLN_tmpP_COMP.KO_SYUBETSU, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
                    Call UniCode_Conv(K0_PLN_tmpP_COMP.KO_JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_PLN_tmpP_COMP.KO_NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_PLN_tmpP_COMP.KO_HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    Call UniCode_Conv(K0_PLN_tmpP_COMP.YOTEI_DT, StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode))


                    sts = BTRV(BtOpGetEqual, PLN_tmpP_COMP_POS, PLN_tmpP_COMP_REC, Len(PLN_tmpP_COMP_REC), K0_PLN_tmpP_COMP, Len(K0_PLN_tmpP_COMP), 0)
                    Select Case sts
                        Case BtNoErr
                            Upd_Com = BtOpUpdate
                        Case BtErrKeyNotFound
                            Upd_Com = BtOpInsert
                        
                        
                            Call UniCode_Conv(PLN_tmpP_COMP_REC.JGYOBU, StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode))
                            Call UniCode_Conv(PLN_tmpP_COMP_REC.NAIGAI, StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode))
                            Call UniCode_Conv(PLN_tmpP_COMP_REC.HIN_GAI, StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode))
                            Call UniCode_Conv(PLN_tmpP_COMP_REC.KO_SYUBETSU, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
                            Call UniCode_Conv(PLN_tmpP_COMP_REC.KO_JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                            Call UniCode_Conv(PLN_tmpP_COMP_REC.KO_NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                            Call UniCode_Conv(PLN_tmpP_COMP_REC.KO_HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                            Call UniCode_Conv(PLN_tmpP_COMP_REC.YOTEI_DT, StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode))
                        
                            Call UniCode_Conv(PLN_tmpP_COMP_REC.YOTEI_QTY, "00000000")
                            Call UniCode_Conv(PLN_tmpP_COMP_REC.KO_QTY, StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))
                            Call UniCode_Conv(PLN_tmpP_COMP_REC.USE_QTY, "000000")
                        
                            Call UniCode_Conv(PLN_tmpP_COMP_REC.DATA_KBN, StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode))
                        
                            Call UniCode_Conv(PLN_tmpP_COMP_REC.INS_TANTO, App.EXEName)
                            Call UniCode_Conv(PLN_tmpP_COMP_REC.Ins_DateTime, INS_NOW)
                        Case Else
                            
                            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                                "資材所要量確認画面中間ファイル作成　異常停止！！[展開処理]", Me.hwnd, 0)
                            
                            Call Input_UnLock
                            Call File_Error(sts, BtOpGetEqual, "資材所要量中間ファイル")
                            Exit Function
                    
                    End Select


                    Call UniCode_Conv(PLN_tmpP_COMP_REC.YOTEI_QTY, Format(CLng(StrConv(PLN_tmpP_COMP_REC.YOTEI_QTY, vbUnicode)) + _
                                                                            CLng(StrConv(PLN_S_YOTEI_R.YOTEI_QTY, vbUnicode)), "00000000"))

                    
                    
                    
                    For i = 0 To UBound(GAISOU_TBL)
                        If Trim(StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode)) = GAISOU_TBL(i) Then
                            Exit For
                        End If
                    Next i
                    
                    If i <= UBound(GAISOU_TBL) Or StrConv(PLN_tmpP_COMP_REC.DATA_KBN, vbUnicode) = P_GAISOU Then
                        '2012.01.07 ０割りチェック追加
                        If Val(StrConv(PLN_tmpP_COMP_REC.KO_QTY, vbUnicode)) = 0 Then
                            Call UniCode_Conv(PLN_tmpP_COMP_REC.USE_QTY, "00000")
                        Else
                            Call UniCode_Conv(PLN_tmpP_COMP_REC.USE_QTY, Format(Int(CLng(StrConv(PLN_tmpP_COMP_REC.YOTEI_QTY, vbUnicode)) / _
                                                                                            CDbl(StrConv(PLN_tmpP_COMP_REC.KO_QTY, vbUnicode))), "000000"))
                        End If
                    Else
                        Call UniCode_Conv(PLN_tmpP_COMP_REC.USE_QTY, Format(Round(CLng(StrConv(PLN_tmpP_COMP_REC.YOTEI_QTY, vbUnicode)) * _
                                                                                        CDbl(StrConv(PLN_tmpP_COMP_REC.KO_QTY, vbUnicode)), 1), "000000"))
                    End If
Debug.Print StrConv(PLN_tmpP_COMP_REC.USE_QTY, vbUnicode)
                    sts = BTRV(Upd_Com, PLN_tmpP_COMP_POS, PLN_tmpP_COMP_REC, Len(PLN_tmpP_COMP_REC), K0_PLN_tmpP_COMP, Len(K0_PLN_tmpP_COMP), 0)
                    Select Case sts
                        Case BtNoErr
                
                        Case Else
                
                
                            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                                "資材所要量確認画面中間ファイル作成　異常停止！！[展開処理]", Me.hwnd, 0)
                            
                            Call Input_UnLock
                            Call File_Error(sts, Upd_Com, "資材所要量中間ファイル")
                            Exit Function
                
                
                
                    End Select

'---------------------------------------------------------->>>>>    資材所要量中間ファイル作成   <<<<<
                End If
'---------------------------------------------------------->>>>>    消費集計   <<<<<
                
'---------------------------------------------------------->>>>>    入荷集計   <<<<<
                Call UniCode_Conv(K5_P_SHORDER.KAN_F, P_KAN_OFF)
                Call UniCode_Conv(K5_P_SHORDER.Y_NOUKI_DT, Format(Now, "YYYYMMDD"))
                Call UniCode_Conv(K5_P_SHORDER.ORDER_CODE, "")
                
                com = BtOpGetGreater
                                                
                Do
                    DoEvents
                
                    sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K5_P_SHORDER, Len(K5_P_SHORDER), 5)
                    Select Case sts
                        Case BtNoErr
                            
                            If StrConv(P_SHORDER_REC.KAN_F, vbUnicode) <> P_KAN_OFF Then
                                Exit Do
                            End If
                            
                            If StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode) > Format(End_Date, "YYYYMMDD") Then
                                Exit Do
                            End If
                        
                        Case BtErrEOF
                            Exit Do
                        Case Else
                            
                            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                                "資材所要量確認画面中間ファイル作成　異常停止！！[展開処理]", Me.hwnd, 0)
                            
                            Call Input_UnLock
                            Call File_Error(sts, com, "資材注文データ")
                            Exit Function
                    
                    End Select
            
                    SKIP_FLG = False
                    If StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
                        SKIP_FLG = True
                    End If
            
                    If Trim(StrConv(P_SHORDER_REC.JGYOBU, vbUnicode)) <> StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode) Or _
                        Trim(StrConv(P_SHORDER_REC.NAIGAI, vbUnicode)) <> StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode) Or _
                        Trim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)) <> Trim(StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)) Then
                        SKIP_FLG = True
                    End If
            
            
                    If Not SKIP_FLG Then
                        Call UniCode_Conv(K0_PLN_tmpZaiko.SYUBETSU, "")
                        Call UniCode_Conv(K0_PLN_tmpZaiko.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                        Call UniCode_Conv(K0_PLN_tmpZaiko.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                        Call UniCode_Conv(K0_PLN_tmpZaiko.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                        Call UniCode_Conv(K0_PLN_tmpZaiko.RIREKI_DT, StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode))
                        
                        sts = BTRV(BtOpGetEqual, PLN_tmpZaiko_POS, PLN_tmpZaikoREC, Len(PLN_tmpZaikoREC), K0_PLN_tmpZaiko, Len(K0_PLN_tmpZaiko), 0)
                        Select Case sts
                            Case BtNoErr
                                Upd_Com = BtOpUpdate
                            Case BtErrKeyNotFound
                            
                                Call UniCode_Conv(PLN_tmpZaikoREC.SYUBETSU, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
                                Call UniCode_Conv(PLN_tmpZaikoREC.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                                Call UniCode_Conv(PLN_tmpZaikoREC.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                                Call UniCode_Conv(PLN_tmpZaikoREC.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                                Call UniCode_Conv(PLN_tmpZaikoREC.RIREKI_DT, StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode))
                                Call UniCode_Conv(PLN_tmpZaikoREC.DATA_KBN, StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode))
                                
                                
                                If Zaiko_Syukei_Proc(SUMI_QTY, _
                                                        MI_QTY, _
                                                        StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode), _
                                                        StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode), _
                                                        StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)) Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                
                                Call UniCode_Conv(PLN_tmpZaikoREC.ST_ZAIKO_QTY, Format(SUMI_QTY + MI_QTY, "000000"))
                                Call UniCode_Conv(PLN_tmpZaikoREC.SYOUHI_QTY, "000000")
                                Call UniCode_Conv(PLN_tmpZaikoREC.NYUKA_QTY, "000000")
                                Call UniCode_Conv(PLN_tmpZaikoREC.ZAIKO_QTY, "000000")
                                Call UniCode_Conv(PLN_tmpZaikoREC.INS_TANTO, App.EXEName)
                                Call UniCode_Conv(PLN_tmpZaikoREC.Ins_DateTime, INS_NOW)
                                
                                Upd_Com = BtOpInsert
            
                            Case Else
                                
                                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                                    "資材所要量確認画面中間ファイル作成　異常停止！！[展開処理]", Me.hwnd, 0)
                                
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual, "資材所要量確認画面中間ファイル")
                                Exit Function
                        
                        End Select
                        Call UniCode_Conv(PLN_tmpZaikoREC.NYUKA_QTY, Format(CLng(StrConv(PLN_tmpZaikoREC.NYUKA_QTY, vbUnicode)) + _
                                                                                        (CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) - _
                                                                                        CLng(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode))), "000000"))
                        
                        sts = BTRV(Upd_Com, PLN_tmpZaiko_POS, PLN_tmpZaikoREC, Len(PLN_tmpZaikoREC), K0_PLN_tmpZaiko, Len(K0_PLN_tmpZaiko), 0)
                        Select Case sts
                            Case BtNoErr
                    
                            Case Else
                    
                    
                                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                                    "資材所要量確認画面中間ファイル作成　異常停止！！[展開処理]", Me.hwnd, 0)
                                
                                Call Input_UnLock
                                Call File_Error(sts, Upd_Com, "資材所要量確認画面中間ファイル")
                                Exit Function
                    
                        End Select
                    End If
                    '資材注文ﾃﾞｰﾀのﾙｰﾌﾟ
                    com = BtOpGetNext
                Loop
                
                '構成ﾏｽﾀのﾙｰﾌﾟ
                com = BtOpGetNext
            Loop
        End If
        
        '商品化予定ファイルのﾙｰﾌﾟ
        com = BtOpGetNext
    Loop
                                    
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<歯抜け日付分埋め込み>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "資材所要量確認画面中間ファイル作成！！[日別展開処理開始]", Me.hwnd, 0)
    
    
    
    com = BtOpGetFirst
    
    Do
        DoEvents
        sts = BTRV(com, PLN_tmpZaiko_POS, PLN_tmpZaikoREC, Len(PLN_tmpZaikoREC), K0_PLN_tmpZaiko, Len(K0_PLN_tmpZaiko), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "資材所要量確認画面中間ファイル作成　異常停止！！[日別展開処理]", Me.hwnd, 0)
                
                Call Input_UnLock
                Call File_Error(sts, BtOpGetEqual, "資材所要量確認画面中間ファイル")
                Exit Function
                                
        End Select
                                
        svSYUBETSU = StrConv(PLN_tmpZaikoREC.SYUBETSU, vbUnicode)   '種別
        svJGYOBU = StrConv(PLN_tmpZaikoREC.JGYOBU, vbUnicode)       '事業部区分
        svNAIGAI = StrConv(PLN_tmpZaikoREC.NAIGAI, vbUnicode)       '国内外
        svHIN_GAI = StrConv(PLN_tmpZaikoREC.HIN_GAI, vbUnicode)     '品番（外部）
        svDATA_KBN = StrConv(PLN_tmpZaikoREC.DATA_KBN, vbUnicode)   'ﾃﾞｰﾀ区分
                                
                                
                                
        For i = 0 To List_Day
            
            DoEvents
            
            Ing_Date = DateAdd("d", i, Start_Date)
                                            
            Call UniCode_Conv(K0_PLN_tmpZaiko.SYUBETSU, svSYUBETSU)
            Call UniCode_Conv(K0_PLN_tmpZaiko.JGYOBU, svJGYOBU)
            Call UniCode_Conv(K0_PLN_tmpZaiko.NAIGAI, svNAIGAI)
            Call UniCode_Conv(K0_PLN_tmpZaiko.HIN_GAI, svHIN_GAI)
            Call UniCode_Conv(K0_PLN_tmpZaiko.RIREKI_DT, Format(Ing_Date, "YYYYMMDD"))
                                
                                        
            sts = BTRV(BtOpGetEqual, PLN_tmpZaiko_POS, PLN_tmpZaikoREC, Len(PLN_tmpZaikoREC), K0_PLN_tmpZaiko, Len(K0_PLN_tmpZaiko), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                
                
                    Call UniCode_Conv(PLN_tmpZaikoREC.SYUBETSU, svSYUBETSU)
                    Call UniCode_Conv(PLN_tmpZaikoREC.JGYOBU, svJGYOBU)
                    Call UniCode_Conv(PLN_tmpZaikoREC.NAIGAI, svNAIGAI)
                    Call UniCode_Conv(PLN_tmpZaikoREC.HIN_GAI, svHIN_GAI)
                    Call UniCode_Conv(PLN_tmpZaikoREC.RIREKI_DT, Format(Ing_Date, "YYYYMMDD"))
                    Call UniCode_Conv(PLN_tmpZaikoREC.DATA_KBN, svDATA_KBN)
                    
                    
                    If Zaiko_Syukei_Proc(SUMI_QTY, _
                                            MI_QTY, _
                                            svJGYOBU, _
                                            svNAIGAI, _
                                            svHIN_GAI) Then
                        Call Input_UnLock
                        Exit Function
                    End If
    
                    Call UniCode_Conv(PLN_tmpZaikoREC.ST_ZAIKO_QTY, Format(SUMI_QTY + MI_QTY, "000000"))
                    Call UniCode_Conv(PLN_tmpZaikoREC.SYOUHI_QTY, "000000")
                    Call UniCode_Conv(PLN_tmpZaikoREC.NYUKA_QTY, "000000")
                    Call UniCode_Conv(PLN_tmpZaikoREC.ZAIKO_QTY, "000000")
                    Call UniCode_Conv(PLN_tmpZaikoREC.INS_TANTO, App.EXEName)
                    Call UniCode_Conv(PLN_tmpZaikoREC.Ins_DateTime, INS_NOW)
                
                
                    sts = BTRV(BtOpInsert, PLN_tmpZaiko_POS, PLN_tmpZaikoREC, Len(PLN_tmpZaikoREC), K0_PLN_tmpZaiko, Len(K0_PLN_tmpZaiko), 0)
                    Select Case sts
                        Case BtNoErr
                
                        Case Else
                            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                                "資材所要量確認画面中間ファイル作成　異常停止！！[日別展開処理]", Me.hwnd, 0)
                            
                            Call Input_UnLock
                            Call File_Error(sts, Upd_Com, "資材所要量確認画面中間ファイル")
                            Exit Function
                
                    End Select
                
                
                Case Else
                    
                    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                        "資材所要量確認画面中間ファイル作成　異常停止！！[日別展開処理]", Me.hwnd, 0)
                    
                    Call Input_UnLock
                    Call File_Error(sts, BtOpGetEqual, "資材所要量確認画面中間ファイル")
                    Exit Function
                                    
            End Select
                                        
                                        
                                        
        Next i
    
    
    
        Call UniCode_Conv(K0_PLN_tmpZaiko.SYUBETSU, svSYUBETSU)
        Call UniCode_Conv(K0_PLN_tmpZaiko.JGYOBU, svJGYOBU)
        Call UniCode_Conv(K0_PLN_tmpZaiko.NAIGAI, svNAIGAI)
        Call UniCode_Conv(K0_PLN_tmpZaiko.HIN_GAI, svHIN_GAI)
        Call UniCode_Conv(K0_PLN_tmpZaiko.RIREKI_DT, "zzzzzzzz")
    
    
        com = BtOpGetGreater
    
    Loop
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<在庫残設定処理>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "資材所要量確認画面中間ファイル作成！！[在庫残設定処理開始]", Me.hwnd, 0)


    Fast_Flg = True
    com = BtOpGetFirst
    Do
        DoEvents
    
        sts = BTRV(com, PLN_tmpZaiko_POS, PLN_tmpZaikoREC, Len(PLN_tmpZaikoREC), K0_PLN_tmpZaiko, Len(K0_PLN_tmpZaiko), 0)
        Select Case sts
            Case BtNoErr
            
                If Fast_Flg Then
                    svSYUBETSU = StrConv(PLN_tmpZaikoREC.SYUBETSU, vbUnicode)   '種別
                    svJGYOBU = StrConv(PLN_tmpZaikoREC.JGYOBU, vbUnicode)       '事業部区分
                    svNAIGAI = StrConv(PLN_tmpZaikoREC.NAIGAI, vbUnicode)       '国内外
                    svHIN_GAI = StrConv(PLN_tmpZaikoREC.HIN_GAI, vbUnicode)     '品番（外部）
                    
                    ST_ZAIKO_QTY = CLng(StrConv(PLN_tmpZaikoREC.ST_ZAIKO_QTY, vbUnicode))
                                        
                    Fast_Flg = False
                End If
                            
            Case BtErrEOF
                
                Exit Do
            
            Case Else
                
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "資材所要量確認画面中間ファイル作成　異常停止！！[在庫残設定処理開始]", Me.hwnd, 0)
                
                Call Input_UnLock
                Call File_Error(sts, BtOpGetEqual, "資材所要量確認画面中間ファイル")
                Exit Function
                                
        End Select
    
    
        If Trim(svSYUBETSU) <> Trim(StrConv(PLN_tmpZaikoREC.SYUBETSU, vbUnicode)) Or _
            svJGYOBU <> StrConv(PLN_tmpZaikoREC.JGYOBU, vbUnicode) Or _
            svNAIGAI <> StrConv(PLN_tmpZaikoREC.NAIGAI, vbUnicode) Or _
            Trim(svHIN_GAI) <> Trim(StrConv(PLN_tmpZaikoREC.HIN_GAI, vbUnicode)) Then
            
            
            svSYUBETSU = StrConv(PLN_tmpZaikoREC.SYUBETSU, vbUnicode)   '種別
            svJGYOBU = StrConv(PLN_tmpZaikoREC.JGYOBU, vbUnicode)       '事業部区分
            svNAIGAI = StrConv(PLN_tmpZaikoREC.NAIGAI, vbUnicode)       '国内外
            svHIN_GAI = StrConv(PLN_tmpZaikoREC.HIN_GAI, vbUnicode)     '品番（外部）
            
            
            ST_ZAIKO_QTY = CLng(StrConv(PLN_tmpZaikoREC.ST_ZAIKO_QTY, vbUnicode))


        End If
    
    
    
    
        ST_ZAIKO_QTY = ST_ZAIKO_QTY - CLng(StrConv(PLN_tmpZaikoREC.SYOUHI_QTY, vbUnicode)) + CLng(StrConv(PLN_tmpZaikoREC.NYUKA_QTY, vbUnicode))
                    
        If ST_ZAIKO_QTY < 0 Then
            Call UniCode_Conv(PLN_tmpZaikoREC.ZAIKO_QTY, Format(ST_ZAIKO_QTY, "00000"))
        Else
            Call UniCode_Conv(PLN_tmpZaikoREC.ZAIKO_QTY, Format(ST_ZAIKO_QTY, "000000"))
        End If
    
    
        sts = BTRV(BtOpUpdate, PLN_tmpZaiko_POS, PLN_tmpZaikoREC, Len(PLN_tmpZaikoREC), K0_PLN_tmpZaiko, Len(K0_PLN_tmpZaiko), 0)
        Select Case sts
            Case BtNoErr
                            
            
            Case Else
                
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "資材所要量確認画面中間ファイル作成　異常停止！！[在庫残設定処理開始]", Me.hwnd, 0)
                
                Call Input_UnLock
                Call File_Error(sts, BtOpGetEqual, "資材所要量確認画面中間ファイル")
                Exit Function
                                
        End Select
    
        com = BtOpGetNext
    
    Loop


                                    'ＩＮＩ処理日付出力
    If WriteIni(App.EXEName, "LAST_DATETIME", App.EXEName, Format(Now, "YYYY/MM/DD") & " " & Format(Now, "HH:MM:SS")) Then
        Beep
        MsgBox ("INIファイルの書き込みに失敗しました。" & App.EXEName & "LAST_DATETIME")
        Unload Me
    End If

    If WriteIni(App.EXEName, "LAST_DATE", App.EXEName, Text1(ptxStart_Date).Text) Then
        Beep
        MsgBox ("INIファイルの書き込みに失敗しました。" & App.EXEName & "Start_Date")
        Unload Me
    End If



    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "資材所要量確認画面中間ファイル作成　処理開始！！[展開処理終了]", Me.hwnd, 0)


    Update_Proc = False
    Call Input_UnLock
    Exit Function




End Function

Private Function List_Disp_Proc(Start_Date As String) As Integer
'----------------------------------------------------------------------------
'                   「資材所要量確認画面」表示処理
'----------------------------------------------------------------------------
Dim sts                 As Integer
Dim ans                 As Integer
Dim com                 As Integer


Dim Row                 As Long
Dim List_Day            As Long
Dim i                   As Long
Dim j                   As Long
Dim k                   As Long
Dim wkday               As Long


Dim Ing_Date            As String
'Dim Start_Date      As String
Dim End_Date            As String

Dim Yobi                As Integer
Dim Yobi_NAME           As String

Dim Z_RIREKI_Tbl()      As Z_RIREKI_Tbl_tag

Dim SKIP_FLG            As Integer

Dim c                   As String * 128


Dim TOTAL_SYOUHI_QTY    As Long
Dim TOTAL_NYUKA_QTY     As Long


    List_Disp_Proc = True

    Call Input_Lock

    List_Disp_F = True



hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "資材所要量確認画面　[検索]処理開始！！", Me.hwnd, 0)


    List_Day = List_Week * 7
    End_Date = DateAdd("d", List_Day - 1, Start_Date)

    k = -1
    Erase DATE_TBL

    For i = colDAY To List_Day + colTITLE
        
        
        Ing_Date = DateAdd("d", i - colDAY, Start_Date)
        Yobi = Weekday(Ing_Date)
        
        Select Case Yobi
            Case 1
                Yobi_NAME = "(" & "日" & ")"
            Case 2
                Yobi_NAME = "(" & "月" & ")"
            Case 3
                Yobi_NAME = "(" & "火" & ")"
            Case 4
                Yobi_NAME = "(" & "水" & ")"
            Case 5
                Yobi_NAME = "(" & "木" & ")"
            Case 6
                Yobi_NAME = "(" & "金" & ")"
            Case 7
                Yobi_NAME = "(" & "土" & ")"
        End Select
                
                
        
        TDBGrid1.Columns(i).Caption = Right(Format(Ing_Date, "YYYY/MM/DD"), 5) & "   　" & Yobi_NAME


        k = k + 1
        ReDim Preserve DATE_TBL(0 To k)
        DATE_TBL(k) = Ing_Date


    Next i
    
    
    TDBGrid1.Columns(i).Caption = "合　計"
    
    Call UniCode_Conv(K1_PLN_tmpZaiko.RIREKI_DT, Format(Start_Date, "YYYYMMDD"))
    Call UniCode_Conv(K1_PLN_tmpZaiko.SYUBETSU, "")
    Call UniCode_Conv(K1_PLN_tmpZaiko.JGYOBU, "")
    Call UniCode_Conv(K1_PLN_tmpZaiko.NAIGAI, "")
    Call UniCode_Conv(K1_PLN_tmpZaiko.HIN_GAI, "")
    
    com = BtOpGetGreater


    i = -1
    Do
        DoEvents
        
        sts = BTRV(com, PLN_tmpZaiko_POS, PLN_tmpZaikoREC, Len(PLN_tmpZaikoREC), K1_PLN_tmpZaiko, Len(K1_PLN_tmpZaiko), 1)
        Select Case sts
            Case BtNoErr
                
                If StrConv(PLN_tmpZaikoREC.RIREKI_DT, vbUnicode) > Format(End_Date, "YYYYMMDD") Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "資材所要量確認画面　異常停止！！", Me.hwnd, 0)
                
                Call Input_UnLock
                Call File_Error(sts, com, "資材所要量確認画面中間ファイル")
                Exit Function
        
        End Select
    
        SKIP_FLG = True
        '資材表示のﾁｪｯｸ
        If Check1(chkSHIZAI).Value = vbChecked Then
            For j = 0 To UBound(SHIZAI_TBL)
                If Trim(StrConv(PLN_tmpZaikoREC.SYUBETSU, vbUnicode)) = SHIZAI_TBL(j) Then
                    SKIP_FLG = False
                    Exit For
                End If
            Next j
        End If
        '個装表示のﾁｪｯｸ
        If Check1(chkKOSOU).Value = vbChecked Then
            For j = 0 To UBound(KOSOU_TBL)
                If Trim(StrConv(PLN_tmpZaikoREC.SYUBETSU, vbUnicode)) = KOSOU_TBL(j) Then
                    SKIP_FLG = False
                    Exit For
                End If
            Next j
        
            If StrConv(PLN_tmpZaikoREC.DATA_KBN, vbUnicode) = P_KOSOU Then
                SKIP_FLG = False
            End If
        End If
        '外装表示のﾁｪｯｸ
        If Check1(chkGAISOU).Value = vbChecked Then
            For j = 0 To UBound(GAISOU_TBL)
                If Trim(StrConv(PLN_tmpZaikoREC.SYUBETSU, vbUnicode)) = GAISOU_TBL(j) Then
                    SKIP_FLG = False
                    Exit For
                End If
            Next j
        
            If StrConv(PLN_tmpZaikoREC.DATA_KBN, vbUnicode) = P_GAISOU Then
                SKIP_FLG = False
            End If
        End If
        '同梱表示のﾁｪｯｸ
        If Check1(chkDOUKON).Value = vbChecked Then
            For j = 0 To UBound(DOUKON_TBL)
                If Trim(StrConv(PLN_tmpZaikoREC.SYUBETSU, vbUnicode)) = DOUKON_TBL(j) Then
                    SKIP_FLG = False
                    Exit For
                End If
            Next j
        End If
        '構成表示のﾁｪｯｸ
        If Check1(chkKOUSEI).Value = vbChecked Then
            For j = 0 To UBound(KOUSEI_TBL)
                If Trim(StrConv(PLN_tmpZaikoREC.SYUBETSU, vbUnicode)) = KOUSEI_TBL(j) Then
                    SKIP_FLG = False
                    Exit For
                End If
            Next j
        End If
    
    
        
        If Not SKIP_FLG Then
            If i = -1 Then
                i = i + 1
                ReDim Z_RIREKI_Tbl(0 To i)
                ReDim Z_RIREKI_Tbl(i).SYOUHI_QTY(0 To List_Day - 1)
                ReDim Z_RIREKI_Tbl(i).NYUKA_QTY(0 To List_Day - 1)
                ReDim Z_RIREKI_Tbl(i).ZAIKO_QTY(0 To List_Day - 1)
                For j = 0 To UBound(Z_RIREKI_Tbl(i).SYOUHI_QTY)
                    Z_RIREKI_Tbl(i).SYOUHI_QTY(j) = 0
                    Z_RIREKI_Tbl(i).NYUKA_QTY(j) = 0
                    Z_RIREKI_Tbl(i).ZAIKO_QTY(j) = 0
                Next j
                
                Z_RIREKI_Tbl(i).SYUBETSU = StrConv(PLN_tmpZaikoREC.SYUBETSU, vbUnicode)
                Z_RIREKI_Tbl(i).JGYOBU = StrConv(PLN_tmpZaikoREC.JGYOBU, vbUnicode)
                Z_RIREKI_Tbl(i).NAIGAI = StrConv(PLN_tmpZaikoREC.NAIGAI, vbUnicode)
                Z_RIREKI_Tbl(i).HIN_GAI = StrConv(PLN_tmpZaikoREC.HIN_GAI, vbUnicode)
                Z_RIREKI_Tbl(i).Start_Zaiko_QTY = CLng(StrConv(PLN_tmpZaikoREC.ST_ZAIKO_QTY, vbUnicode))
                
                wkday = DateDiff("d", Start_Date, Mid(StrConv(PLN_tmpZaikoREC.RIREKI_DT, vbUnicode), 1, 4) & "/" & _
                                                Mid(StrConv(PLN_tmpZaikoREC.RIREKI_DT, vbUnicode), 5, 2) & "/" & _
                                                Mid(StrConv(PLN_tmpZaikoREC.RIREKI_DT, vbUnicode), 7, 2))
                Z_RIREKI_Tbl(i).SYOUHI_QTY(wkday) = CLng(StrConv(PLN_tmpZaikoREC.SYOUHI_QTY, vbUnicode))
                Z_RIREKI_Tbl(i).NYUKA_QTY(wkday) = CLng(StrConv(PLN_tmpZaikoREC.NYUKA_QTY, vbUnicode))
                Z_RIREKI_Tbl(i).ZAIKO_QTY(wkday) = CLng(StrConv(PLN_tmpZaikoREC.ZAIKO_QTY, vbUnicode))
            
                Z_RIREKI_Tbl(i).DATA_KBN = StrConv(PLN_tmpZaikoREC.DATA_KBN, vbUnicode)
            
            Else
            
                For i = 0 To UBound(Z_RIREKI_Tbl)
                
                
                    If Z_RIREKI_Tbl(i).SYUBETSU = StrConv(PLN_tmpZaikoREC.SYUBETSU, vbUnicode) And _
                        Z_RIREKI_Tbl(i).JGYOBU = StrConv(PLN_tmpZaikoREC.JGYOBU, vbUnicode) And _
                        Z_RIREKI_Tbl(i).NAIGAI = StrConv(PLN_tmpZaikoREC.NAIGAI, vbUnicode) And _
                        Z_RIREKI_Tbl(i).HIN_GAI = StrConv(PLN_tmpZaikoREC.HIN_GAI, vbUnicode) Then
                        
                        Exit For
                
                    End If
                Next i
            
            
                If i > UBound(Z_RIREKI_Tbl) Then
                
                
                    ReDim Preserve Z_RIREKI_Tbl(0 To i)
                    ReDim Z_RIREKI_Tbl(i).SYOUHI_QTY(0 To List_Day - 1)
                    ReDim Z_RIREKI_Tbl(i).NYUKA_QTY(0 To List_Day - 1)
                    ReDim Z_RIREKI_Tbl(i).ZAIKO_QTY(0 To List_Day - 1)
                    For j = 0 To UBound(Z_RIREKI_Tbl(i).SYOUHI_QTY)
                        Z_RIREKI_Tbl(i).SYOUHI_QTY(j) = 0
                        Z_RIREKI_Tbl(i).NYUKA_QTY(j) = 0
                        Z_RIREKI_Tbl(i).ZAIKO_QTY(j) = 0
                    Next j
                    
                    Z_RIREKI_Tbl(i).SYUBETSU = StrConv(PLN_tmpZaikoREC.SYUBETSU, vbUnicode)
                    Z_RIREKI_Tbl(i).JGYOBU = StrConv(PLN_tmpZaikoREC.JGYOBU, vbUnicode)
                    Z_RIREKI_Tbl(i).NAIGAI = StrConv(PLN_tmpZaikoREC.NAIGAI, vbUnicode)
                    Z_RIREKI_Tbl(i).HIN_GAI = StrConv(PLN_tmpZaikoREC.HIN_GAI, vbUnicode)
                    Z_RIREKI_Tbl(i).Start_Zaiko_QTY = CLng(StrConv(PLN_tmpZaikoREC.ST_ZAIKO_QTY, vbUnicode))
                
                
                    Z_RIREKI_Tbl(i).DATA_KBN = StrConv(PLN_tmpZaikoREC.DATA_KBN, vbUnicode)
                
                
                
                End If
            
            
            
                wkday = DateDiff("d", Start_Date, Mid(StrConv(PLN_tmpZaikoREC.RIREKI_DT, vbUnicode), 1, 4) & "/" & _
                                                Mid(StrConv(PLN_tmpZaikoREC.RIREKI_DT, vbUnicode), 5, 2) & "/" & _
                                                Mid(StrConv(PLN_tmpZaikoREC.RIREKI_DT, vbUnicode), 7, 2))
                Z_RIREKI_Tbl(i).SYOUHI_QTY(wkday) = Z_RIREKI_Tbl(i).SYOUHI_QTY(wkday) + CLng(StrConv(PLN_tmpZaikoREC.SYOUHI_QTY, vbUnicode))
                Z_RIREKI_Tbl(i).NYUKA_QTY(wkday) = Z_RIREKI_Tbl(i).NYUKA_QTY(wkday) + CLng(StrConv(PLN_tmpZaikoREC.NYUKA_QTY, vbUnicode))
                Z_RIREKI_Tbl(i).ZAIKO_QTY(wkday) = Z_RIREKI_Tbl(i).ZAIKO_QTY(wkday) + CLng(StrConv(PLN_tmpZaikoREC.ZAIKO_QTY, vbUnicode))
            
            
            End If
        End If
    
        com = BtOpGetNext
    
    
    Loop

    Set ZAIKO_RIREKI = Nothing
    Row = Min_Row - 1



    If i > -1 Then
    
        For i = 0 To UBound(Z_RIREKI_Tbl)
        
        
            '種別／品番／消費／入荷／在庫
            Row = Row + 1
            ZAIKO_RIREKI.ReDim Min_Row, Row, Min_Col, Max_Col + 1
            ZAIKO_RIREKI.ReDim Min_Row, Row + 1, Min_Col, Max_Col + 1
            ZAIKO_RIREKI.ReDim Min_Row, Row + 2, Min_Col, Max_Col + 1
            
            
            Select Case Z_RIREKI_Tbl(i).DATA_KBN
                
                Case P_KOSOU
                    Call UniCode_Conv(K0_P_CODE.C_Code, KOSOU_CODE)
                
                Case P_GAISOU
                    Call UniCode_Conv(K0_P_CODE.C_Code, GAISOU_CODE)
                Case P_DOUKON
                    Call UniCode_Conv(K0_P_CODE.C_Code, Z_RIREKI_Tbl(i).SYUBETSU)
            End Select
            
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_CODEREC.C_NAME, "")
                    Call UniCode_Conv(P_CODEREC.C_RNAME, "")
                Case Else
                    
                    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                        "資材所要量確認画面　異常停止！！", Me.hwnd, 0)
                    Call Input_UnLock
                    Call File_Error(sts, BtOpGetEqual, "資材所要量確認画面中間ファイル")
                    Exit Function
            
            End Select
            '種別
            ZAIKO_RIREKI(Row, colSYUBETSU_CODE) = Z_RIREKI_Tbl(i).SYUBETSU
            If Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode)) = "" Then
                ZAIKO_RIREKI(Row, colSYUBETSU) = Trim(StrConv(P_CODEREC.C_NAME, vbUnicode))
            Else
                ZAIKO_RIREKI(Row, colSYUBETSU) = Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode))
            End If
            
            '事業部
            ZAIKO_RIREKI(Row, colJGYOBU) = Trim(Z_RIREKI_Tbl(i).JGYOBU)
            
            '品番
            ZAIKO_RIREKI(Row, colHIN_GAI) = Trim(Z_RIREKI_Tbl(i).HIN_GAI)
            
            'TITLE
            ZAIKO_RIREKI(Row, colTITLE) = "消　費"
            ZAIKO_RIREKI(Row + 1, colTITLE) = "入　荷"
            ZAIKO_RIREKI(Row + 2, colTITLE) = "在庫残"
            
            TOTAL_SYOUHI_QTY = 0
            TOTAL_NYUKA_QTY = 0
            
            '消費／入荷／在庫
            For j = colDAY To Max_Col
                ZAIKO_RIREKI(Row, j) = Format(CLng(Z_RIREKI_Tbl(i).SYOUHI_QTY(j - colDAY)), "#,###")
                TOTAL_SYOUHI_QTY = TOTAL_SYOUHI_QTY + CLng(Z_RIREKI_Tbl(i).SYOUHI_QTY(j - colDAY))
                
                ZAIKO_RIREKI(Row + 1, j) = Format(CLng(Z_RIREKI_Tbl(i).NYUKA_QTY(j - colDAY)), "#,###")
                TOTAL_NYUKA_QTY = TOTAL_NYUKA_QTY + CLng(Z_RIREKI_Tbl(i).NYUKA_QTY(j - colDAY))
                
                ZAIKO_RIREKI(Row + 2, j) = Format(CLng(Z_RIREKI_Tbl(i).ZAIKO_QTY(j - colDAY)), "#,###")
            Next j
        
        
            ZAIKO_RIREKI(Row, Max_Col + 1) = Format(TOTAL_SYOUHI_QTY, "#,###")
            ZAIKO_RIREKI(Row + 1, Max_Col + 1) = Format(TOTAL_NYUKA_QTY, "#,###")
            ZAIKO_RIREKI(Row + 2, Max_Col + 1) = ""
        
        
        
            Row = Row + 2
        
        Next i
    
    
    

        Set TDBGrid1.Array = ZAIKO_RIREKI
        TDBGrid1.ReBind
        
        TDBGrid1.Update
        TDBGrid1.MoveFirst

    End If


    List_Disp_F = False


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "資材所要量確認画面　[検索]処理終了！！", Me.hwnd, 0)



                                '前回展開年月日時分秒
    If GetIni(App.EXEName, "LAST_DATETIME", App.EXEName, c) Then
        lblLast_DATETIME.Caption = ""
        LAST_DATETIME = ""
    Else
        lblLast_DATETIME.Caption = "　前回展開日時：" & Trim(c) & "現在"
        LAST_DATETIME = Trim(c)
    End If


                                '前回展開開始年月日
    If GetIni(App.EXEName, "LAST_DATE", App.EXEName, c) Then
        lblLAST_DATE.Caption = ""
        lblLAST_DATE.Caption = ""
    Else
        lblLAST_DATE.Caption = "前回展開開始日：" & Trim(c)
        LAST_Date = Trim(c)
    End If



    Call Input_UnLock


    List_Disp_Proc = False

End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    PLN00701.MousePointer = vbHourglass

    Call Ctrl_Lock(PLN00701)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(PLN00701)


    PLN00701.MousePointer = vbDefault

End Sub


Private Sub List_Make_Proc()
'----------------------------------------------------------------------------
'                   Grid行作成
'----------------------------------------------------------------------------

Dim i       As Long
Dim NISUU   As Long


Dim TColumn     As TrueDBGrid80.Column
 



    

    NISUU = List_Week * 7


    For i = colDAY + 1 To NISUU + colTITLE
        
        Set TColumn = TDBGrid1.Columns.Add(i)
        With TColumn
            .Visible = True
            .Caption = ""
            .Width = TDBGrid1.Columns(colDAY).Width
            .Font.NAME = "ＭＳ ゴシック"
            .Font.Size = TDBGrid1.Columns(colDAY).Font.Size
            .HeadFont.NAME = "ＭＳ ゴシック"
            .HeadFont.Size = 9
            .Alignment = dbgRight
            .FetchStyle = True
        End With
    Next i


'---    2011.12.15 合計列
    Set TColumn = TDBGrid1.Columns.Add(i)
    With TColumn
        .Visible = True
        .Caption = ""
        .Width = TDBGrid1.Columns(colDAY).Width
        .Font.NAME = "ＭＳ ゴシック"
        .Font.Size = TDBGrid1.Columns(colDAY).Font.Size
        .HeadFont.NAME = "ＭＳ ゴシック"
        .HeadFont.Size = 9
        .Alignment = dbgRight
        .FetchStyle = True
        .AllowFocus = False
            
    End With
'---    2011.12.15 合計列




    Max_Col = NISUU + colTITLE

End Sub

Private Sub TDBGrid1_DblClick()
    
    If TDBGrid1.Bookmark = Null Then
        Exit Sub
    End If
    If TDBGrid1.Bookmark <= 0 Then
        Exit Sub
    End If
    
    If TDBGrid1.Col = Null Then
        Exit Sub
    End If
    
    If TDBGrid1.Col < colDAY Then
        Exit Sub
    End If
    
    If ZAIKO_RIREKI(TDBGrid1.Bookmark, TDBGrid1.Col) = 0 Or Trim(ZAIKO_RIREKI(TDBGrid1.Bookmark, TDBGrid1.Col)) = "" Then
        Exit Sub
    End If

    If Trim(ZAIKO_RIREKI(TDBGrid1.Bookmark, colHIN_GAI)) = "" Then
        Exit Sub
    End If


    DISP_KO_SYUBETSU_CODE = ZAIKO_RIREKI(TDBGrid1.Bookmark, colSYUBETSU_CODE)
    DISP_KO_JGYOBU = ZAIKO_RIREKI(TDBGrid1.Bookmark, colJGYOBU)
    DISP_KO_HIN_GAI = ZAIKO_RIREKI(TDBGrid1.Bookmark, colHIN_GAI)
    DISP_DATE = DATE_TBL(TDBGrid1.Col - colDAY)
    
    PLN00702.Show vbModal
    
    
    
End Sub

Private Sub TDBGrid1_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueDBGrid80.StyleDisp)

Dim i   As Integer



    If svBookMark = -1 Then
        svBookMark = Bookmark
    End If

    
    If Trim(ZAIKO_RIREKI(Bookmark, colHIN_GAI)) <> "" Then
        If svBookMark <> Bookmark Then
            If wkBackColor Then
                wkBackColor = False
            Else
                wkBackColor = True
            End If
        End If
        svBookMark = Bookmark
    End If

    Select Case wkBackColor
        Case False
            CellStyle = Rstyle_Blue
        Case True
            CellStyle = Rstyle_White
    End Select


    If ZAIKO_RIREKI(Bookmark, Col) < 0 Then
        CellStyle = Rstyle_Red
    Else
        CellStyle = Rstyle_Black
    End If

End Sub

' ------------------------------------------------------------------------
'       指定した精度の数値に切り上げします。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に切り上げられた数値。
' ------------------------------------------------------------------------
Private Function ToRoundUp(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    
        


    dCoef = (10 ^ iDigits)



    
    
    
    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
        Case Is > 0
            ToRoundUp = (Int(dValue * dCoef) + 1) / dCoef
        Case Is < 0
            ToRoundUp = (Fix(dValue * dCoef) - 1) / dCoef
        Case Else
            ToRoundUp = dValue
    End Select




End Function

